/* Exercise the same adapter code as a real x86 and x64 native client.
 * Office tests separately exercise DLL loading, VBA declarations and events. */
#include "../../src/WIP/stdWebSocket/native/stdWebSocketNative.c"
#include <stdio.h>
#include <stdlib.h>
#define CHECK(value) do { if (!(value)) { fprintf(stderr, "Native assertion failed at line %d\n", __LINE__); return 1; } } while (0)

typedef struct TestReceiver {
    IDispatch dispatch;
    LONG references;
    Socket *socket;
    DWORD owner, messages, closes;
    BOOL failed;
} TestReceiver;
static HRESULT STDMETHODCALLTYPE test_qi(IDispatch *self, REFIID iid, void **out) {
    if (memcmp(iid, &IID_IDispatch, sizeof(*iid)) && memcmp(iid, &IID_IUnknown, sizeof(*iid))) {
        *out = NULL; return E_NOINTERFACE;
    }
    *out = self; self->lpVtbl->AddRef(self); return S_OK;
}
static ULONG STDMETHODCALLTYPE test_addref(IDispatch *self) { return ++((TestReceiver *)self)->references; }
static ULONG STDMETHODCALLTYPE test_release(IDispatch *self) { return --((TestReceiver *)self)->references; }
static HRESULT STDMETHODCALLTYPE test_count(IDispatch *self, UINT *count) { (void)self; *count = 0; return S_OK; }
static HRESULT STDMETHODCALLTYPE test_info(IDispatch *self, UINT index, LCID locale, ITypeInfo **info) {
    (void)self; (void)index; (void)locale; (void)info; return E_NOTIMPL;
}
static HRESULT STDMETHODCALLTYPE test_names(IDispatch *self, REFIID iid, LPOLESTR *names, UINT count, LCID locale, DISPID *id) {
    (void)self; (void)iid; (void)locale;
    if (count != 1 || wcscmp(names[0], L"protAsyncDispatch")) return DISP_E_UNKNOWNNAME;
    *id = 42; return S_OK;
}
static HRESULT STDMETHODCALLTYPE test_invoke(IDispatch *self, DISPID id, REFIID iid, LCID locale, WORD flags,
                                              DISPPARAMS *params, VARIANT *returned, EXCEPINFO *exception, UINT *argument) {
    TestReceiver *r = (TestReceiver *)self;
    VARIANT data, receiver;
    DWORD type, code, result;
    BYTE *bytes;
    (void)iid; (void)locale; (void)flags; (void)params; (void)returned; (void)exception; (void)argument;
    if (id != 42 || GetCurrentThreadId() != r->owner) { r->failed = TRUE; return E_FAIL; }
    VariantInit(&data);
    result = stdWebSocketNext(r->socket, &data, &type, &code);
    if (result == ERROR_NO_MORE_ITEMS) return S_OK;
    if (result) { r->failed = TRUE; return E_FAIL; }
    if (type == 0) {
        SafeArrayAccessData(V_ARRAY(&data), (void **)&bytes);
        if (bytes[0] != r->messages) r->failed = TRUE;
        SafeArrayUnaccessData(V_ARRAY(&data));
        r->messages++;
    } else if (type == 4) {
        if (code != 3004) r->failed = TRUE;
        r->closes++;
        VariantInit(&receiver); V_VT(&receiver) = VT_DISPATCH; V_DISPATCH(&receiver) = self;
        stdWebSocketUnlisten(r->socket, &receiver);
    } else r->failed = TRUE;
    VariantClear(&data);
    return S_OK;
}
static IDispatchVtbl test_vtable = {test_qi, test_addref, test_release, test_count, test_info, test_names, test_invoke};
static void pump_messages(void) {
    MSG message;
    while (PeekMessageW(&message, NULL, 0, 0, PM_REMOVE)) DispatchMessageW(&message);
}
static int async_bridge(DWORD port) {
    Socket *s = NULL;
    VARIANT argument, empty;
    TestReceiver receiver = {0};
    ULONGLONG deadline;
    DWORD count;
    int i;
    receiver.dispatch.lpVtbl = &test_vtable; receiver.references = 1; receiver.owner = GetCurrentThreadId();
    CHECK(!stdWebSocketConnect(L"127.0.0.1", port, L"/burst", 0, 2000, &s));
    receiver.socket = s;
    VariantInit(&argument); V_VT(&argument) = VT_DISPATCH; V_DISPATCH(&argument) = &receiver.dispatch;
    CHECK(!stdWebSocketListen(s, &argument));
    deadline = GetTickCount64() + 2000;
    do {
        EnterCriticalSection(&s->queueLock); count = s->queuedCount; LeaveCriticalSection(&s->queueLock);
        if (count == QUEUE_MESSAGES) break;
        Sleep(5);
    } while (GetTickCount64() < deadline);
    CHECK(count == QUEUE_MESSAGES && receiver.messages == 0);
    deadline = GetTickCount64() + 5000;
    while (!receiver.closes && GetTickCount64() < deadline) { pump_messages(); Sleep(1); }
    CHECK(receiver.messages == 100 && receiver.closes == 1 && !receiver.failed);
    CHECK(!stdWebSocketListening(s) && receiver.references == 1);
    stdWebSocketDisconnect(&s);
    /* Pending notifications must not invoke a released receiver, even if
     * Windows recycles the old HWND for the next registration. */
    VariantInit(&empty);
    for (i = 0; i < 30; i++) {
        CHECK(!stdWebSocketConnect(L"127.0.0.1", port, L"/silent", 0, 2000, &s));
        receiver.socket = s;
        CHECK(!stdWebSocketListen(s, &argument));
        CHECK(!stdWebSocketListen(s, &empty));
        CHECK(!stdWebSocketListen(s, &argument));
        stdWebSocketDisconnect(&s);
        pump_messages();
        CHECK(receiver.messages == 100 && receiver.references == 1 && !receiver.failed);
    }
    return 0;
}

static int roundtrip(DWORD port) {
    Socket *s = NULL;
    BYTE message[] = {'a', 0, 0xe2, 0x82, 0xac};
    BYTE buffer[64];
    DWORD n = 0;
    WINHTTP_WEB_SOCKET_BUFFER_TYPE type = WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE;
    CHECK(!stdWebSocketConnect(L"127.0.0.1", port, L"/echo?test=1", 0, 2000, &s));
    CHECK(!stdWebSocketOperation(&s, 1, message, sizeof(message), &n, &type, 2000, 0));
    CHECK(!stdWebSocketOperation(&s, 0, buffer, sizeof(buffer), &n, &type, 2000, 0));
    CHECK(n == sizeof(message) && type == WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE);
    CHECK(!memcmp(buffer, message, sizeof(message)));
    CHECK(!stdWebSocketOperation(&s, 2, NULL, 0, &n, &type, 2000, 1000));
    stdWebSocketDisconnect(&s);
    CHECK(!s);
    stdWebSocketDisconnect(&s);
    return 0;
}
int main(int argc, char **argv) {
    DWORD port, before, after, n = 0, error;
    WINHTTP_WEB_SOCKET_BUFFER_TYPE type;
    Socket *s = NULL;
    BYTE buffer[64];
    ULONGLONG started;
    int i;
    CHECK(argc == 2);
    port = (DWORD)atoi(argv[1]);
    CHECK(!async_bridge(port));
    for (i = 0; i < 5; i++) CHECK(!roundtrip(port));
    Sleep(200);
    CHECK(GetProcessHandleCount(GetCurrentProcess(), &before));
    for (i = 0; i < 100; i++) CHECK(!roundtrip(port));
    /* Late close callbacks must release their context/event references. */
    Sleep(500);
    CHECK(GetProcessHandleCount(GetCurrentProcess(), &after));
    CHECK(after <= before + 16);
    for (i = 0; i < 20; i++) {
        CHECK(!stdWebSocketConnect(L"127.0.0.1", port, L"/silent", 0, 2000, &s));
        started = GetTickCount64();
        error = stdWebSocketOperation(&s, 0, buffer, sizeof(buffer), &n, &type, 20, 0);
        CHECK(error == ERROR_TIMEOUT && !s);
        CHECK(GetTickCount64() - started < 2000);
    }
    Sleep(500);
    CHECK(GetProcessHandleCount(GetCurrentProcess(), &after));
    CHECK(after <= before + 16);
    printf("Passed native %u-bit lifecycle/deadline checks; handles before=%lu after=%lu\n",
           (unsigned)(sizeof(void *) * 8), before, after);
    return 0;
}
