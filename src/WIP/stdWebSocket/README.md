# Native stdWebSocket

`src/stdWebSocket.cls` replaces the Internet Explorer implementation implicated in [#139](https://github.com/sancarn/stdVBA/issues/139). `stdWebSocketNew.cls` contains the same implementation under the WIP class name; import one of them.

The rewrite also addresses the native client intent of [#35](https://github.com/sancarn/stdVBA/issues/35). It uses the **WinHTTP WebSocket API**, rather than the literal `websocket.dll` protocol component. WinHTTP owns the HTTP upgrade, frame masking, control frames, TLS and certificate validation. The lower-level protocol component would require us to supply the network and TLS transports. No client capability requested in those issues requires that additional layer. The list of server functions in #35 does not make this a WebSocket server.

[Wasabi](https://github.com/vbacollective/wasabi/blob/main/package/Wasabi.bas) was reviewed for its client URL/handshake, message, timeout and cleanup design. This implementation uses WinHTTP instead of copying Wasabi's Winsock, TLS or machine-code implementation.

## Build and import

Requires Windows 8 or later, Office VBA, and Visual Studio C++ build tools with a Windows SDK to build the helper:

```powershell
powershell -NoProfile -File src/WIP/stdWebSocket/native/build.ps1 -OutputDirectory C:\libs\stdWebSocket
```

Import `src/stdWebSocket.cls` and the generic `src/stdICallable.cls` interface. Choose the DLL matching **Office's bitness**, independently of Windows bitness, and supply its absolute path:

```vb
Dim ws As stdWebSocket
Set ws = stdWebSocket.Create("ws://127.0.0.1:8080/echo", 5000, _
                            "C:\libs\stdWebSocket\x64\stdWebSocketNative.dll")
ws.Send "hello" & vbNullChar & ChrW(&H20AC)
Debug.Print ws.Receive()
ws.CloseSocket 1000, "done"
```

The helper is a deliberate dependency. Synchronous WinHTTP WebSocket receive does not honor the session's receive timeout on the tested Windows version; the upgraded socket rejects `WINHTTP_OPTION_RECEIVE_TIMEOUT`. A VBA callback on a WinHTTP worker thread would be unsafe. The small C adapter uses asynchronous WinHTTP and native completion events, copies caller buffers, and cancels timed-out operations only after their initiating API has returned. Its reference-counted context and copied buffers survive until final `HANDLE_CLOSING` notifications. The DLL is pinned until process exit so outstanding native cancellation callbacks cannot call unloaded code. No executable memory or callback thunks are generated.

## Asynchronous client delivery

```vb
'In a standard module:
Private ws As stdWebSocket

'Keep ws in the scope that owns this connection. callback is an implementation
'of stdICallable; its Run receives (name As String, eventData As Dictionary).
Public Sub BeginSocket(ByVal callback As stdICallable)
  Set ws = stdWebSocket.Create("ws://127.0.0.1:8080/echo", 5000, _
                              "C:\libs\stdWebSocket\x64\stdWebSocketNative.dll")
  ws.StartListening callback
  ws.Send "hello"          'Synchronous completion; reader remains asynchronous.
End Sub

Public Sub ReadSynchronouslyAgain()
  ws.StopListening
  Debug.Print ws.Receive() 'Consumes the existing receive queue, then live data.
End Sub
```

`StartListening(listener As stdICallable)` returns promptly and registers exclusive read ownership. It delivers `Open`, `Message`, `Close`, and `Error` through the typed `listener.Run(name, eventData)`. It works with a private `Implements stdICallable` implementation; the callable need not expose a public late-bound `Run`. `stdWebSocket` is now non-exposed, following `stdWebView` and `stdArray`, because VBA prohibits exposing the non-exposed interface in an exposed class's public signature.

The adapter creates a message-only window on the connecting Office thread. A native receive worker assembles complete messages into a queue and posts coalesced window notifications. The WndProc invokes a retained IDispatch receiver on that same Office thread; the receiver is a separate instance of this class and calls the typed interface from VBA. VARIANT/COM AddRef/Release own its lifetime. There are no fabricated objects, weak ObjPtr references, or VBA calls from WinHTTP/receive workers. This receiver does not retain the socket owner, so ordinary owner release destroys the notification window and cancels reads. A user-supplied callable that itself retains the socket can create its own cycle; call `StopListening` or `Disconnect` to break it.

Delivery occurs when Office pumps Windows messages, including normal idle time and `DoEvents`. It **does not interrupt a busy macro**. Registration posts a deferred Open; each notification dispatches at most one message. Nested dispatch on the same connection is suppressed even if a handler calls `DoEvents`. Other applications or a busy Office UI can delay notification indefinitely; no real-time callback guarantee is made.

`StopListening` detaches delivery, releases the callable reference, and invalidates old window notifications. The native reader stays alive so subsequent synchronous `Receive` drains the same queue; two reads never compete for the socket. Replacing a listener preserves unread messages for the new listener and gives it a deferred Open. An old callback's failure cannot cancel its replacement. The current message is consumed once before its callback runs. A callback/decoder error stops that registration and is exposed in `LastListenerError`; unread messages remain available for synchronous reading. Transport errors and peer close deliver one terminal notification, detach delivery, and release native transport handles.

Buffered messages and an unread terminal notification can still be drained by a replacement listener after the peer has closed, even though `State` reports NotConnected. Once the terminal queue has been fully consumed, new registration is rejected; disconnect/reinitialize to open another connection.

The queue holds at most 32 complete messages and 16 MiB of queued payload. One additional message of at most 16 MiB can be assembled or held while waiting for space; allocation bookkeeping and geometric buffer capacity add bounded overhead. A full queue applies receive backpressure rather than dropping data. Idle asynchronous connections do not time out merely because no message arrives; once a fragmented message begins, it must complete within the configured message timeout. Stopping delivery does not disable these bounds.

`Send` retains synchronous native completion semantics. It can block under network backpressure, including when called from a listener, until completion or its deadline. Write errors raise to its caller. There is no queued `SendAsync` API. Queue-mode close can take a send/shutdown timeout plus a peer-close timeout. Local close/disconnect notifies the async callable after cleanup; when called from its own handler, Close is deferred until that user handler returns, avoiding nested user callbacks.

The existing `stdWebView` thunk implementation was inspected: its architecture-specific vtable adapters and EbMode checks do not marshal calls between threads. A direct reuse on WinHTTP workers would enter VBA on arbitrary threads. The native boundary is retained for worker queueing, cancellation, COM retention and owner-thread delivery rather than expanding generated code to implement those responsibilities. This remains a DLL-based deployment, rather than a single imported class with embedded executable code.

**VBE Reset, `End`, unloading the VBA project, and forced Office shutdown are unsupported with a live subscription.** They can skip `Class_Terminate`; neither a normal COM reference nor an EbMode check proves that reset VBA code is safe to invoke. Stop listening and disconnect first. Tests cover ordinary stop, replacement, callback errors, disconnect inside handlers, natural owner release, stale notifications and repeated lifetime stress. They do not claim safety across a forced reset or `End`.

## Event timing and migration

The names `EventRaised`, `OnOpen`, `OnMessage`, `OnClose`, `OnError`, `Listeners`, `Send`, `CloseSocket`, `Disconnect`, `handleEvent` and `url` are retained. Timing and payload shape change:

* `Create(url, timeoutMs, nativeLibraryPath)` completes the handshake or raises an error before returning. The unfinished WIP's separate port argument is removed; put the port in the URL.
* Choose `StartListening` for automatic owner-thread delivery to a typed callable, or call `Pump()`/`Receive()` explicitly for synchronous events. A synchronous read is rejected while `Listening` is True. Synchronous calls block for at most one complete message, a peer close, or a deadline, and do not run `DoEvents`.
* `Open` is delivered at the start of the first pump, allowing `WithEvents` handlers and listeners to be attached after `Create`. Sending does not implicitly pump. Message delivery is exactly once per complete message, whether consumed using `Receive` or `Pump`.
* Each event receives a late-bound `Scripting.Dictionary`. Use `eventData("data")` for a message, `eventData("code")` / `eventData("reason")` for close, and `eventData("message")` for an error. Browser `MessageEvent.data` property access must be updated. All dictionaries also contain `type`.
* For synchronous reads, `EventRaised` precedes the specific event, followed by each legacy collection listener's `Run(name, eventData)`. Individual legacy listener errors are ignored, as before. Event-handler errors propagate; an error during a synchronous receive closes that connection. Reentrant `Receive`/`Pump` is rejected. Async network notifications go exclusively to the registered typed callable, not additionally to the legacy events/collection. Local close/disconnect still raise their legacy Close events.
* `Receive` returns a String for UTF8, a Byte() for binary (including a zero-length Byte() for an empty message), and `Empty` for peer close. `Pump` returns True for any message, including empty messages, and False for peer close.
* Local `CloseSocket` and `Disconnect` deliver legacy Close immediately; async Close is deferred until an active user callback returns. `Disconnect` is an idempotent abort, reporting code 1006. Termination releases resources without raising events. Synchronous transport failures release resources, deliver legacy Error and raise a VBA error; constructor failures raise directly because handlers cannot yet be attached.

`Send(data, iType, bPartialSend)` supports text and Byte() data; the default is UTF8. Fragment types must remain consistent until the final fragment. UTF8 conversions use explicit lengths and reject invalid sequences. Each text send fragment must itself be valid Unicode. Incoming UTF8 may split within a multibyte codepoint and is decoded only after reassembly.

URLs accept ws/wss, default ports 80/443, explicit ports, bracketed IPv6, path and query. Use ASCII/Punycode hosts and percent-encoded non-ASCII paths. Credentials, fragments, control characters, whitespace and out-of-range ports are rejected. TLS uses Windows' normal trust and hostname checks, with no certificate bypass. Proxy behavior follows WinHTTP's default proxy configuration.

Send and close have a native operation deadline. Connection send/response phases each have a timeout, so total connection time can exceed one timeout interval. Receive has a deadline for the entire message, including fragments, and a 16 MiB message limit; a timeout closes the connection rather than permitting a retry on a partially consumed message. User event-handler execution time is outside network deadlines. Timeout error 1460 identifies adapter deadlines; native WinHTTP errors retain their numeric code.

## Verification

See [the local test runner](../../../tests/harnesses/stdWebSocket-tests/README.md). It uses disposable local servers and a separate hidden Excel instance. A native helper is necessary at runtime; no built DLL is committed. The x86 and x64 builds are both checked. Office runtime testing depends on the architecture installed on the test machine.
