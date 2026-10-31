/* Native asynchronous WinHTTP adapter: no worker-thread calls into VBA. Context and
 * copied buffers survive cancellation until final HANDLE_CLOSING callbacks.
 * Operations on one context must be serialized by the calling VBA thread. */
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <winhttp.h>
#include <string.h>
#define COBJMACROS
#include <oleauto.h>
#define QUEUE_MESSAGES 32
#define QUEUE_BYTES 16777216
#define NOTIFY_MESSAGE (WM_APP + 0x139)
typedef struct Packet {
    struct Packet *next;
    DWORD length, offset, type, code;
    BYTE *data;
} Packet;
typedef struct Socket {
    HINTERNET session, connection, request, socket;
    HANDLE completed;
    volatile LONG references;
    DWORD result, count;
    WINHTTP_WEB_SOCKET_BUFFER_TYPE type;
    BYTE *buffer;
    DWORD owner;
    CRITICAL_SECTION queueLock, apiLock, sendLock;
    HANDLE readCompleted, stop, space, available, workerDone, worker;
    DWORD readResult, readCount;
    WINHTTP_WEB_SOCKET_BUFFER_TYPE readType;
    BYTE readBuffer[4096];
    Packet *head, *tail;
    DWORD queuedCount, queuedBytes, terminal, timeout;
    DWORD closeCode, closeReasonLength;
    BYTE closeReason[123];
    BOOL queueMode, dispatching, notified, terminalRead, kick, shutdownSent;
    HWND window;
    IDispatch *receiver;
    DISPID dispatchId;
    ULONG generation;
} Socket;
static void stop_listener(Socket *s);
static DWORD queued_read(Socket *s, PVOID buffer, DWORD length, DWORD *read,
                         WINHTTP_WEB_SOCKET_BUFFER_TYPE *type, DWORD timeout);
static ULONG nextGeneration;
static void release(Socket *s) {
    if (InterlockedDecrement(&s->references) == 0) {
        if (s->buffer) HeapFree(GetProcessHeap(), 0, s->buffer);
        if (s->completed) CloseHandle(s->completed);
        if (s->readCompleted) CloseHandle(s->readCompleted);
        if (s->stop) CloseHandle(s->stop);
        if (s->space) CloseHandle(s->space);
        if (s->available) CloseHandle(s->available);
        if (s->workerDone) CloseHandle(s->workerDone);
        if (s->worker) CloseHandle(s->worker);
        while (s->head) {
            Packet *packet = s->head;
            s->head = packet->next;
            if (packet->data) HeapFree(GetProcessHeap(), 0, packet->data);
            HeapFree(GetProcessHeap(), 0, packet);
        }
        DeleteCriticalSection(&s->queueLock);
        DeleteCriticalSection(&s->apiLock);
        DeleteCriticalSection(&s->sendLock);
        HeapFree(GetProcessHeap(), 0, s);
    }
}
static void CALLBACK status(HINTERNET handle, DWORD_PTR context, DWORD notification,
                           LPVOID information, DWORD length) {
    Socket *s = (Socket *)context;
    (void)handle;
    if (!s) return;
    InterlockedIncrement(&s->references);
    switch (notification) {
    case WINHTTP_CALLBACK_STATUS_HANDLE_CLOSING:
        release(s); /* Handle reference. */
        release(s); /* This callback's reference. */
        return;
    case WINHTTP_CALLBACK_STATUS_REQUEST_ERROR:
        if (s->queueMode && length == sizeof(WINHTTP_WEB_SOCKET_ASYNC_RESULT) &&
            ((WINHTTP_WEB_SOCKET_ASYNC_RESULT *)information)->Operation == WINHTTP_WEB_SOCKET_RECEIVE_OPERATION) {
            s->readResult = ((WINHTTP_ASYNC_RESULT *)information)->dwError;
            SetEvent(s->readCompleted);
            release(s);
            return;
        }
        s->result = ((WINHTTP_ASYNC_RESULT *)information)->dwError; break;
    case WINHTTP_CALLBACK_STATUS_READ_COMPLETE:
        if (s->queueMode) {
            if (length != sizeof(WINHTTP_WEB_SOCKET_STATUS)) s->readResult = ERROR_INVALID_DATA;
            else {
                s->readCount = ((WINHTTP_WEB_SOCKET_STATUS *)information)->dwBytesTransferred;
                s->readType = ((WINHTTP_WEB_SOCKET_STATUS *)information)->eBufferType;
            }
            SetEvent(s->readCompleted);
            release(s);
            return;
        }
        if (length != sizeof(WINHTTP_WEB_SOCKET_STATUS)) s->result = ERROR_INVALID_DATA;
        else {
            WINHTTP_WEB_SOCKET_STATUS *value = (WINHTTP_WEB_SOCKET_STATUS *)information;
            s->count = value->dwBytesTransferred;
            s->type = value->eBufferType;
        }
        break;
    case WINHTTP_CALLBACK_STATUS_SENDREQUEST_COMPLETE:
    case WINHTTP_CALLBACK_STATUS_HEADERS_AVAILABLE:
    case WINHTTP_CALLBACK_STATUS_WRITE_COMPLETE:
    case WINHTTP_CALLBACK_STATUS_SHUTDOWN_COMPLETE:
    case WINHTTP_CALLBACK_STATUS_CLOSE_COMPLETE: break;
    default: release(s); return;
    }
    SetEvent(s->completed);
    release(s);
}
static void close_transport(Socket *s) {
    EnterCriticalSection(&s->apiLock);
    if (s->socket) WinHttpCloseHandle(s->socket);
    s->socket = NULL;
    if (s->request) WinHttpCloseHandle(s->request);
    s->request = NULL;
    if (s->connection) WinHttpCloseHandle(s->connection);
    s->connection = NULL;
    if (s->session) WinHttpCloseHandle(s->session);
    s->session = NULL;
    LeaveCriticalSection(&s->apiLock);
}
static void destroy(Socket *s) {
    /* The export retains a reference while closing handles. Every watched
     * handle retains another until its final callback, even after return. */
    stop_listener(s);
    SetEvent(s->stop);
    close_transport(s);
    release(s);
}
static DWORD wait_result(Socket *s, DWORD timeout) {
    DWORD wait = WaitForSingleObject(s->completed, timeout);
    if (wait == WAIT_TIMEOUT) return ERROR_TIMEOUT;
    if (wait != WAIT_OBJECT_0) return GetLastError();
    return s->result;
}
__declspec(dllexport) DWORD WINAPI stdWebSocketConnect(
    LPCWSTR host, DWORD port, LPCWSTR path, DWORD flags, DWORD timeout, Socket **output) {
    Socket *s;
    DWORD result = ERROR_NOT_ENOUGH_MEMORY, size, code, index = 0;
    DWORD_PTR context;
    HMODULE module;
    if (!output) return ERROR_INVALID_PARAMETER;
    *output = NULL;
    if (!timeout || !host || !path || port < 1 || port > 65535) return ERROR_INVALID_PARAMETER;
    /* Pin the tiny DLL for process lifetime: callbacks can still be unwinding
     * when VBA releases its LoadLibrary reference after a canceled operation. */
    if (!GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_PIN,
                           (LPCWSTR)(ULONG_PTR)&status, &module)) return GetLastError();
    s = (Socket *)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(*s));
    if (!s) return result;
    s->references = 1;
    s->owner = GetCurrentThreadId();
    s->timeout = timeout;
    InitializeCriticalSection(&s->queueLock);
    InitializeCriticalSection(&s->apiLock);
    InitializeCriticalSection(&s->sendLock);
    s->completed = CreateEventW(NULL, TRUE, FALSE, NULL);
    s->readCompleted = CreateEventW(NULL, TRUE, FALSE, NULL);
    s->stop = CreateEventW(NULL, TRUE, FALSE, NULL);
    s->space = CreateEventW(NULL, FALSE, FALSE, NULL);
    s->available = CreateEventW(NULL, TRUE, FALSE, NULL);
    s->workerDone = CreateEventW(NULL, TRUE, FALSE, NULL);
    if (!s->completed || !s->readCompleted || !s->stop || !s->space || !s->available || !s->workerDone) goto failed;
    s->session = WinHttpOpen(L"stdVBA/stdWebSocket", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
                            NULL, NULL, WINHTTP_FLAG_ASYNC);
    if (!s->session) goto failed;
    if (!WinHttpSetTimeouts(s->session, timeout, timeout, timeout, timeout)) goto failed;
    s->connection = WinHttpConnect(s->session, host, (INTERNET_PORT)port, 0);
    if (!s->connection) goto failed;
    s->request = WinHttpOpenRequest(s->connection, L"GET", path, NULL, NULL, NULL, flags);
    if (!s->request) goto failed;
    context = (DWORD_PTR)s;
    if (!WinHttpSetOption(s->request, WINHTTP_OPTION_CONTEXT_VALUE, &context, sizeof(context))) goto failed;
    if (WinHttpSetStatusCallback(s->request, status, WINHTTP_CALLBACK_FLAG_ALL_NOTIFICATIONS, 0)
        == WINHTTP_INVALID_STATUS_CALLBACK) goto failed;
    InterlockedIncrement(&s->references);
    if (!WinHttpSetOption(s->request, WINHTTP_OPTION_UPGRADE_TO_WEB_SOCKET, NULL, 0)) goto failed;
    if (!WinHttpSendRequest(s->request, NULL, 0, NULL, 0, 0, context)) goto failed;
    result = wait_result(s, timeout);
    if (result) goto error;
    ResetEvent(s->completed);
    if (!WinHttpReceiveResponse(s->request, NULL)) goto failed;
    result = wait_result(s, timeout);
    if (result) goto error;
    size = sizeof(code);
    if (!WinHttpQueryHeaders(s->request, WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
                            NULL, &code, &size, &index)) goto failed;
    if (code != 101) { result = ERROR_WINHTTP_INVALID_SERVER_RESPONSE; goto error; }
    s->socket = WinHttpWebSocketCompleteUpgrade(s->request, context);
    if (!s->socket) goto failed;
    InterlockedIncrement(&s->references);
    WinHttpCloseHandle(s->request);
    s->request = NULL;
    *output = s;
    return ERROR_SUCCESS;
failed:
    result = GetLastError();
    if (!result) result = ERROR_GEN_FAILURE;
error:
    destroy(s);
    return result;
}
/* 0=receive, 1=send, 2=close, 3=query close status. Caller buffers are never
 * retained by WinHTTP. Timeout cancels async I/O and clears the caller context. */
__declspec(dllexport) DWORD WINAPI stdWebSocketOperation(
    Socket **socket, DWORD operation, PVOID buffer, DWORD length,
    DWORD *read, WINHTTP_WEB_SOCKET_BUFFER_TYPE *type, DWORD timeout, USHORT closeCode) {
    Socket *s;
    DWORD result;
    if (!socket || !*socket || !timeout || operation > 3) return ERROR_INVALID_PARAMETER;
    s = *socket;
    if (s->owner != GetCurrentThreadId()) return ERROR_INVALID_THREAD_ID;
    if (operation == 0 && s->queueMode) {
        if (s->receiver) return ERROR_BUSY;
        result = queued_read(s, buffer, length, read, type, timeout);
        if (result) { *socket = NULL; destroy(s); }
        return result;
    }
    if (operation == 3) {
        USHORT code = 0;
        if (s->queueMode) {
            EnterCriticalSection(&s->queueLock);
            *read = s->closeReasonLength; *type = (WINHTTP_WEB_SOCKET_BUFFER_TYPE)s->closeCode;
            result = length < *read ? ERROR_INSUFFICIENT_BUFFER : 0;
            if (!result && *read) memcpy(buffer, s->closeReason, *read);
            LeaveCriticalSection(&s->queueLock);
            return result;
        }
        result = WinHttpWebSocketQueryCloseStatus(s->socket, &code, buffer, length, read);
        *type = (WINHTTP_WEB_SOCKET_BUFFER_TYPE)code;
        return result;
    }
    if (s->buffer) HeapFree(GetProcessHeap(), 0, s->buffer);
    s->buffer = NULL;
    if (length) {
        s->buffer = (BYTE *)HeapAlloc(GetProcessHeap(), 0, length);
        if (!s->buffer) return ERROR_NOT_ENOUGH_MEMORY;
        if (operation != 0) memcpy(s->buffer, buffer, length);
    }
    EnterCriticalSection(&s->sendLock);
    if (operation == 2 && s->queueMode && s->shutdownSent) {
        LeaveCriticalSection(&s->sendLock);
        result = WaitForSingleObject(s->workerDone, timeout) == WAIT_OBJECT_0 ? 0 : ERROR_TIMEOUT;
        if (result) { *socket = NULL; destroy(s); }
        return result;
    }
    s->result = 0;
    s->count = 0;
    ResetEvent(s->completed);
    EnterCriticalSection(&s->apiLock);
    switch (operation) {
    case 0: result = WinHttpWebSocketReceive(s->socket, s->buffer, length, NULL, NULL); break;
    case 1: result = WinHttpWebSocketSend(s->socket, *type, s->buffer, length); break;
    case 2:
        if (s->queueMode) result = WinHttpWebSocketShutdown(s->socket, closeCode, s->buffer, length);
        else result = WinHttpWebSocketClose(s->socket, closeCode, s->buffer, length);
        break;
    default: result = ERROR_INVALID_PARAMETER;
    }
    LeaveCriticalSection(&s->apiLock);
    if (result == ERROR_IO_PENDING) result = 0;
    if (!result && operation == 2 && s->queueMode) s->shutdownSent = TRUE;
    if (!result) result = wait_result(s, timeout);
    if (!result && operation == 0) {
        if (s->count > length) result = ERROR_INVALID_DATA;
        else {
            if (s->count) memcpy(buffer, s->buffer, s->count);
            *read = s->count;
            *type = s->type;
        }
    }
    LeaveCriticalSection(&s->sendLock);
    if (!result && operation == 2 && s->queueMode &&
        WaitForSingleObject(s->workerDone, timeout) != WAIT_OBJECT_0) result = ERROR_TIMEOUT;
    if (result) {
        *socket = NULL;
        destroy(s);
    }
    return result;
}
__declspec(dllexport) void WINAPI stdWebSocketDisconnect(Socket **socket) {
    Socket *s;
    if (!socket) return;
    s = *socket;
    *socket = NULL;
    if (s) destroy(s);
}

/* Queue/window code never enters COM from WinHTTP or receive worker threads. */
static void notify_locked(Socket *s) {
    if (s->window && s->receiver && !s->notified && !s->dispatching) {
        s->notified = PostMessageW(s->window, NOTIFY_MESSAGE, s->generation, 0);
    }
}
static void free_packet(Packet *packet) {
    if (packet->data) HeapFree(GetProcessHeap(), 0, packet->data);
    HeapFree(GetProcessHeap(), 0, packet);
}
static BOOL enqueue(Socket *s, BYTE *data, DWORD length, DWORD type, DWORD code) {
    HANDLE waits[2] = {s->stop, s->space};
    Packet *packet = (Packet *)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(*packet));
    if (!packet) { if (data) HeapFree(GetProcessHeap(), 0, data); return FALSE; }
    packet->data = data; packet->length = length; packet->type = type; packet->code = code;
    for (;;) {
        EnterCriticalSection(&s->queueLock);
        if (s->queuedCount < QUEUE_MESSAGES && s->queuedBytes <= QUEUE_BYTES - length) {
            if (s->tail) s->tail->next = packet; else s->head = packet;
            s->tail = packet;
            s->queuedCount++; s->queuedBytes += length;
            if (type >= 4) s->terminal = type == 4 ? 1 : code;
            SetEvent(s->available);
            notify_locked(s);
            LeaveCriticalSection(&s->queueLock);
            return TRUE;
        }
        LeaveCriticalSection(&s->queueLock);
        if (WaitForMultipleObjects(2, waits, FALSE, INFINITE) != WAIT_OBJECT_0 + 1) {
            free_packet(packet); return FALSE;
        }
    }
}
static DWORD WINAPI receive_worker(PVOID parameter) {
    Socket *s = (Socket *)parameter;
    HANDLE waits[2] = {s->stop, s->readCompleted};
    BYTE *data = NULL;
    DWORD size = 0, capacity = 0, error = 0, family = 0, wait, timeout;
    ULONGLONG deadline = 0;
    for (;;) {
        s->readResult = 0; s->readCount = 0;
        ResetEvent(s->readCompleted);
        EnterCriticalSection(&s->apiLock);
        if (WaitForSingleObject(s->stop, 0) == WAIT_OBJECT_0) {
            LeaveCriticalSection(&s->apiLock); break;
        }
        error = WinHttpWebSocketReceive(s->socket, s->readBuffer, sizeof(s->readBuffer), NULL, NULL);
        LeaveCriticalSection(&s->apiLock);
        if (error && error != ERROR_IO_PENDING) break;
        timeout = INFINITE;
        if (deadline) {
            ULONGLONG now = GetTickCount64();
            if (now >= deadline) { error = ERROR_TIMEOUT; break; }
            timeout = (DWORD)(deadline - now);
        }
        wait = WaitForMultipleObjects(2, waits, FALSE, timeout);
        if (wait == WAIT_OBJECT_0) { error = 0; break; }
        if (wait != WAIT_OBJECT_0 + 1) { error = wait == WAIT_TIMEOUT ? ERROR_TIMEOUT : GetLastError(); break; }
        error = s->readResult;
        if (error) break;
        if (s->readType == WINHTTP_WEB_SOCKET_CLOSE_BUFFER_TYPE) {
            USHORT code = 0;
            BYTE *reason = (BYTE *)HeapAlloc(GetProcessHeap(), 0, 123);
            DWORD reasonLength = 0;
            if (!reason) { error = ERROR_NOT_ENOUGH_MEMORY; break; }
            error = WinHttpWebSocketQueryCloseStatus(s->socket, &code, reason, 123, &reasonLength);
            if (error) { HeapFree(GetProcessHeap(), 0, reason); break; }
            EnterCriticalSection(&s->sendLock);
            if (!s->shutdownSent) {
                s->result = 0; ResetEvent(s->completed);
                EnterCriticalSection(&s->apiLock);
                if (s->socket && WaitForSingleObject(s->stop, 0) != WAIT_OBJECT_0)
                    error = WinHttpWebSocketShutdown(s->socket, code, NULL, 0);
                else error = ERROR_OPERATION_ABORTED;
                LeaveCriticalSection(&s->apiLock);
                if (error == ERROR_IO_PENDING) error = 0;
                if (!error) {
                    s->shutdownSent = TRUE;
                    error = wait_result(s, s->timeout);
                }
            }
            LeaveCriticalSection(&s->sendLock);
            if (error) { HeapFree(GetProcessHeap(), 0, reason); break; }
            EnterCriticalSection(&s->queueLock);
            s->closeCode = code; s->closeReasonLength = reasonLength;
            if (reasonLength) memcpy(s->closeReason, reason, reasonLength);
            LeaveCriticalSection(&s->queueLock);
            close_transport(s);
            if (!enqueue(s, reason, reasonLength, 4, code)) error = ERROR_NOT_ENOUGH_MEMORY;
            break;
        }
        if (s->readType > WINHTTP_WEB_SOCKET_UTF8_FRAGMENT_BUFFER_TYPE ||
            s->readCount > sizeof(s->readBuffer) || size > QUEUE_BYTES - s->readCount) {
            error = ERROR_BUFFER_OVERFLOW; break;
        }
        if (!deadline) { deadline = GetTickCount64() + s->timeout; family = s->readType / 2; }
        if ((DWORD)s->readType / 2 != family) { error = ERROR_INVALID_DATA; break; }
        if (size + s->readCount > capacity) {
            BYTE *grown;
            if (!capacity) capacity = 4096;
            while (capacity < size + s->readCount) capacity *= 2;
            grown = data ? (BYTE *)HeapReAlloc(GetProcessHeap(), 0, data, capacity)
                         : (BYTE *)HeapAlloc(GetProcessHeap(), 0, capacity);
            if (!grown) { error = ERROR_NOT_ENOUGH_MEMORY; break; }
            data = grown;
        }
        if (s->readCount) memcpy(data + size, s->readBuffer, s->readCount);
        size += s->readCount;
        if (s->readType == WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE ||
            s->readType == WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE) {
            BYTE *message = data; data = NULL;
            if (!enqueue(s, message, size, family * 2, 0)) { error = ERROR_NOT_ENOUGH_MEMORY; break; }
            size = capacity = 0; deadline = 0;
        }
    }
    if (data) HeapFree(GetProcessHeap(), 0, data);
    if (error && WaitForSingleObject(s->stop, 0) != WAIT_OBJECT_0) {
        /* Cancel a fragment read that may still be pending, but do not release
         * any COM object here. All receive storage lives until HANDLE_CLOSING. */
        close_transport(s);
        enqueue(s, NULL, 0, 5, error);
        EnterCriticalSection(&s->queueLock);
        s->terminal = error;
        SetEvent(s->available); notify_locked(s);
        LeaveCriticalSection(&s->queueLock);
    }
    SetEvent(s->workerDone);
    release(s);
    return 0;
}
static void remove_head_locked(Socket *s) {
    Packet *packet = s->head;
    s->head = packet->next;
    if (!s->head) { s->tail = NULL; ResetEvent(s->available); }
    s->queuedCount--; s->queuedBytes -= packet->length;
    if (packet->type >= 4) s->terminalRead = TRUE;
    free_packet(packet);
    SetEvent(s->space);
}
static DWORD queued_read(Socket *s, PVOID buffer, DWORD length, DWORD *read,
                         WINHTTP_WEB_SOCKET_BUFFER_TYPE *type, DWORD timeout) {
    DWORD wait = WaitForSingleObject(s->available, timeout), result = 0;
    Packet *packet;
    if (wait != WAIT_OBJECT_0) return wait == WAIT_TIMEOUT ? ERROR_TIMEOUT : GetLastError();
    EnterCriticalSection(&s->queueLock);
    packet = s->head;
    if (!packet) result = s->terminal > 1 ? s->terminal : ERROR_CONNECTION_ABORTED;
    else if (packet->type == 5) result = packet->code;
    else {
        DWORD remaining = packet->length - packet->offset;
        *read = remaining < length ? remaining : length;
        if (*read) memcpy(buffer, packet->data + packet->offset, *read);
        packet->offset += *read;
        *type = (WINHTTP_WEB_SOCKET_BUFFER_TYPE)(packet->type + (packet->offset < packet->length ? 1 : 0));
        if (packet->offset == packet->length) remove_head_locked(s);
    }
    LeaveCriticalSection(&s->queueLock);
    return result;
}
static LRESULT CALLBACK notification_window(HWND window, UINT message, WPARAM token, LPARAM parameter) {
    Socket *s = (Socket *)GetWindowLongPtrW(window, GWLP_USERDATA);
    if (message == WM_NCCREATE) {
        s = (Socket *)((CREATESTRUCTW *)parameter)->lpCreateParams;
        SetWindowLongPtrW(window, GWLP_USERDATA, (LONG_PTR)s);
        InterlockedIncrement(&s->references);
    }
    if (message == WM_NCDESTROY && s) {
        SetWindowLongPtrW(window, GWLP_USERDATA, 0);
        release(s);
    } else if (message == NOTIFY_MESSAGE && s) {
        IDispatch *receiver = NULL;
        DISPID id = 0;
        HRESULT result;
        DISPPARAMS params = {0};
        EXCEPINFO exception = {0};
        VARIANT returned;
        InterlockedIncrement(&s->references);
        EnterCriticalSection(&s->queueLock);
        if (token == s->generation && !s->dispatching && s->receiver) {
            receiver = s->receiver; id = s->dispatchId;
            receiver->lpVtbl->AddRef(receiver);
            s->dispatching = TRUE; s->notified = FALSE;
            s->kick = FALSE;
        }
        LeaveCriticalSection(&s->queueLock);
        if (receiver) {
            VariantInit(&returned);
            result = receiver->lpVtbl->Invoke(receiver, id, &IID_NULL, LOCALE_USER_DEFAULT,
                                            DISPATCH_METHOD, &params, &returned, &exception, NULL);
            VariantClear(&returned);
            if (exception.pfnDeferredFillIn) exception.pfnDeferredFillIn(&exception);
            SysFreeString(exception.bstrSource); SysFreeString(exception.bstrDescription); SysFreeString(exception.bstrHelpFile);
            receiver->lpVtbl->Release(receiver);
            if (FAILED(result) && token == s->generation) stop_listener(s);
            EnterCriticalSection(&s->queueLock);
            s->dispatching = FALSE;
            if (s->kick || s->head || (s->terminal && !s->terminalRead)) notify_locked(s);
            LeaveCriticalSection(&s->queueLock);
        }
        release(s);
        return 0;
    }
    return DefWindowProcW(window, message, token, parameter);
}
static INIT_ONCE windowClassOnce = INIT_ONCE_STATIC_INIT;
static HMODULE windowModule;
static BOOL CALLBACK register_window(PINIT_ONCE once, PVOID parameter, PVOID *context) {
    WNDCLASSW windowClass = {0};
    (void)once; (void)parameter; (void)context;
    if (!GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                           (LPCWSTR)(ULONG_PTR)&notification_window, &windowModule)) return FALSE;
    windowClass.lpfnWndProc = notification_window;
    windowClass.hInstance = windowModule;
    windowClass.lpszClassName = L"stdVBA.WebSocket.Notification.1";
    return RegisterClassW(&windowClass) != 0;
}
static void stop_listener(Socket *s) {
    IDispatch *receiver;
    HWND window;
    EnterCriticalSection(&s->queueLock);
    receiver = s->receiver; window = s->window;
    s->receiver = NULL; s->window = NULL; s->notified = FALSE; s->kick = FALSE;
    s->generation = (ULONG)InterlockedIncrement((volatile LONG *)&nextGeneration);
    LeaveCriticalSection(&s->queueLock);
    if (window) DestroyWindow(window);
    if (receiver) receiver->lpVtbl->Release(receiver);
}
/* A normal VARIANT/IDispatch argument gives explicit COM lifetime ownership;
 * neither the adapter nor VBA fabricates an object from an ObjPtr address. */
__declspec(dllexport) DWORD WINAPI stdWebSocketListen(Socket *s, VARIANT *receiver) {
    LPOLESTR name = L"protAsyncDispatch";
    DISPID id;
    HRESULT result;
    HWND window;
    if (!s || s->owner != GetCurrentThreadId()) return ERROR_INVALID_THREAD_ID;
    if (!receiver || V_VT(receiver) == VT_EMPTY) { stop_listener(s); return 0; }
    EnterCriticalSection(&s->queueLock);
    if (s->terminal && s->terminalRead && !s->head) {
        LeaveCriticalSection(&s->queueLock);
        return ERROR_CONNECTION_ABORTED;
    }
    LeaveCriticalSection(&s->queueLock);
    if (V_VT(receiver) != VT_DISPATCH || !V_DISPATCH(receiver)) return ERROR_INVALID_PARAMETER;
    result = V_DISPATCH(receiver)->lpVtbl->GetIDsOfNames(V_DISPATCH(receiver), &IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &id);
    if (FAILED(result)) return ERROR_INVALID_FUNCTION;
    if (!InitOnceExecuteOnce(&windowClassOnce, register_window, NULL, NULL)) return GetLastError();
    window = CreateWindowExW(0, L"stdVBA.WebSocket.Notification.1", L"", 0, 0, 0, 0, 0,
                             HWND_MESSAGE, NULL, windowModule, s);
    if (!window) return GetLastError();
    stop_listener(s);
    EnterCriticalSection(&s->queueLock);
    s->window = window; s->receiver = V_DISPATCH(receiver); s->dispatchId = id;
    s->kick = TRUE;
    s->receiver->lpVtbl->AddRef(s->receiver);
    notify_locked(s); /* Deferred Open, even if the receive queue is empty. */
    LeaveCriticalSection(&s->queueLock);
    if (!s->queueMode) {
        s->queueMode = TRUE;
        InterlockedIncrement(&s->references);
        s->worker = CreateThread(NULL, 0, receive_worker, s, 0, NULL);
        if (!s->worker) {
            DWORD error = GetLastError();
            s->queueMode = FALSE; release(s); stop_listener(s);
            return error;
        }
    }
    return 0;
}
__declspec(dllexport) DWORD WINAPI stdWebSocketListening(Socket *s) {
    return s && s->receiver ? 1 : 0;
}
__declspec(dllexport) DWORD WINAPI stdWebSocketUnlisten(Socket *s, VARIANT *receiver) {
    IUnknown *current = NULL, *candidate = NULL;
    BOOL same;
    if (!s || s->owner != GetCurrentThreadId()) return ERROR_INVALID_THREAD_ID;
    if (!receiver || V_VT(receiver) != VT_DISPATCH || !V_DISPATCH(receiver)) return ERROR_INVALID_PARAMETER;
    if (!s->receiver) return 0;
    s->receiver->lpVtbl->QueryInterface(s->receiver, &IID_IUnknown, (void **)&current);
    V_DISPATCH(receiver)->lpVtbl->QueryInterface(V_DISPATCH(receiver), &IID_IUnknown, (void **)&candidate);
    same = current && current == candidate;
    if (current) current->lpVtbl->Release(current);
    if (candidate) candidate->lpVtbl->Release(candidate);
    if (same) stop_listener(s);
    return 0;
}
__declspec(dllexport) DWORD WINAPI stdWebSocketTerminal(Socket *s) {
    DWORD terminal;
    if (!s) return 1;
    EnterCriticalSection(&s->queueLock); terminal = s->terminal; LeaveCriticalSection(&s->queueLock);
    return terminal;
}
__declspec(dllexport) DWORD WINAPI stdWebSocketNext(Socket *s, VARIANT *data, DWORD *type, DWORD *code) {
    Packet *packet;
    SAFEARRAY *array;
    void *bytes;
    DWORD result = 0;
    if (!s || s->owner != GetCurrentThreadId()) return ERROR_INVALID_THREAD_ID;
    EnterCriticalSection(&s->queueLock);
    packet = s->head;
    if (!packet) {
        if (s->terminal > 1 && !s->terminalRead) {
            *type = 5; *code = s->terminal; s->terminalRead = TRUE; VariantInit(data);
        } else result = ERROR_NO_MORE_ITEMS;
    } else {
        array = SafeArrayCreateVector(VT_UI1, 0, packet->length - packet->offset);
        if (!array) result = ERROR_NOT_ENOUGH_MEMORY;
        else {
            SafeArrayAccessData(array, &bytes);
            if (packet->length > packet->offset) memcpy(bytes, packet->data + packet->offset, packet->length - packet->offset);
            SafeArrayUnaccessData(array);
            VariantInit(data); V_VT(data) = VT_ARRAY | VT_UI1; V_ARRAY(data) = array;
            *type = packet->type; *code = packet->code;
            remove_head_locked(s);
        }
    }
    LeaveCriticalSection(&s->queueLock);
    return result;
}
