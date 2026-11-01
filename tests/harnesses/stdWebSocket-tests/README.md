# Local WebSocket regressions

Run from the repository root:

```powershell
python tests/harnesses/stdWebSocket-tests/run.py
```

Requirements: Python with `websockets` (tested with the asyncio server API), Excel, Visual Studio C++ build tools with Windows SDK, and OpenSSL (PATH or Git for Windows' bundled executable). The runner builds x86/x64 native helpers in a temporary build directory and selects the DLL using a VBA `Win64` check in the actual Office process.

Coverage includes UTF8/non-BMP/embedded-null and empty text, exact binary bytes and lengths (including nonzero input lower bounds and empty arrays), outgoing/incoming fragmentation, split UTF8 codepoints, events and listeners, local/peer close, invalid URLs, refused connections, stalled receive/fragment/send/close/handshake, rejected upgrade, message size limit, and rejection of an untrusted local TLS certificate. The TLS fixture never installs a certificate or changes trust stores. Successful TLS negotiation against a trusted local certificate is not tested.

Async coverage uses a private `Implements stdICallable` test class with no public late-bound Run. One test returns from VBA, lets Excel idle while the PowerShell worker sleeps, and verifies deferred delivery without Pump/Receive/DoEvents in that interval. A busy macro first verifies that notifications do not interrupt it. Other checks cover callback thread IDs, exactly-once delivery, async/sync ownership switching, replacement from a callback (including an old callback failing after replacement), callback error recovery, fragmented/empty/embedded-null data, message deadlines, local/peer close, disconnect from Open/Message, non-reentrancy through DoEvents, stale HWND notifications, owner release and listener destructor counts.

Native x86 and x64 executables also exercise the message-window bridge using a reference-counted IDispatch test receiver. They fill the queue to its 32-message bound without pumping, then verify ordered delivery of 100 messages on the registering thread, peer-close acknowledgement, COM reference release, repeated HWND replacement/destruction, connection lifecycles and deadline cancellation. Office runtime tests cover the installed Office architecture; forced VBE Reset/End is unsupported and not tested.

The parent starts a modal-window watcher before Office automation and enforces a 45-second Office deadline. It identifies the owned process through Excel's window handle, rejects an existing Excel process, dismisses owned modal dialogs as test failures, minimizes owned windows, and kills only the owned process on a hang. Temporary workbooks are closed in `finally` blocks. Actual source bytes are checked for CRLF before import, including generated WIP test wrappers.

The worker records the actual Excel version's AccessVBOM value and type, and the actual VBE version's two error-trapping values and types. This runner does not enable project access or change debugging settings to start tests: it requires access already enabled and Break on Unhandled Errors already selected, and otherwise fails before importing/running VBA. It verifies/restores the original recorded values in cleanup. This keeps existing Office sessions undisturbed. On this test machine the original values were AccessVBOM DWORD 1, BreakOnAllErrors DWORD 0, BreakOnServerErrors DWORD 0; none required changes.

To run the same assertions against the WIP class identity:

```powershell
$env:STDWEBSOCKET_WIP = '1'
python tests/harnesses/stdWebSocket-tests/run.py
Remove-Item Env:STDWEBSOCKET_WIP
```

The runner generates temporary test/recorder wrappers referring to `stdWebSocketNew`, then imports the actual WIP class source. Generated wrappers are also checked for CRLF bytes immediately before import.
