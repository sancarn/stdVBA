Attribute VB_Name = "stdWebSocketAsyncTests"
Option Explicit
Public destroyedRecorders As Long
Private idleSocket As stdWebSocket
Private idleRecorder As WebSocketAsyncRecorder
Private checks As Long
Private Sub Assert(ByVal condition As Boolean, ByVal message As String)
  If Not condition Then Err.Raise 5, "Async tests", message
  checks = checks + 1
End Sub
Public Function BeginIdle(ByVal base As String, ByVal native As String) As String
  Set idleSocket = stdWebSocket.Create(base & "/idle", 1000, native)
  Set idleRecorder = New WebSocketAsyncRecorder
  Dim started As Double: started = Timer
  idleSocket.StartListening idleRecorder
  Assert Timer - started < 0.5, "Registration returns promptly"
  Do While Timer - started < 0.4: Loop
  Assert idleRecorder.calls = 0, "Does not interrupt busy VBA"
  BeginIdle = "Idle listener registered; busy macro deferred callbacks"
End Function
Public Function CheckIdle() As String
  On Error GoTo failed
  Assert idleRecorder.opens = 1 And idleRecorder.messages = 2 And idleRecorder.closes = 1, "Idle event counts"
  Assert idleRecorder.wrongThread = 0 And idleRecorder.maxDepth = 1, "Idle thread affinity and non-reentrancy"
  Assert idleRecorder.payloads(1) = "idle" & vbNullChar & ChrW(&H20AC), "Idle UTF8 explicit lengths"
  Dim bytes() As Byte: bytes = idleRecorder.lastData
  Assert UBound(bytes) = 1 And bytes(0) = 0 And bytes(1) = 255, "Async binary payload"
  Assert Not idleSocket.Listening And idleSocket.State = NotConnected, "Terminal listener cleanup"
  idleSocket.Disconnect
  Set idleSocket = Nothing: Set idleRecorder = Nothing
  CheckIdle = "Passed real idle delivery"
  Exit Function
failed:
  CheckIdle = "FAILED idle: " & Err.Description
End Function
Private Sub WaitMessages(ByVal recorder As WebSocketAsyncRecorder, ByVal count As Long)
  Dim started As Double: started = Timer
  Do While recorder.messages < count And Timer - started < 3
    DoEvents
  Loop
  Assert recorder.messages >= count, "Message delivery deadline"
End Sub
Public Function RunTests(ByVal base As String, ByVal native As String) As String
  On Error GoTo failed
  Dim ws As stdWebSocket, cb As WebSocketAsyncRecorder, other As WebSocketAsyncRecorder
  Dim value As Variant, bytes() As Byte, i As Long, previous As Long
  Set ws = stdWebSocket.Create(base & "/burst", 1000, native)
  Set cb = New WebSocketAsyncRecorder: Set cb.socket = ws: cb.stopOnMessage = True
  ws.StartListening cb
  WaitMessages cb, 1
  Assert cb.readConflict <> 0 And Not ws.Listening, "Exclusive read ownership and stop from callback"
  bytes = cb.lastData: Assert bytes(0) = 0, "First message delivered once"
  For i = 1 To 99
    value = ws.Receive(): bytes = value
    Assert bytes(0) = i, "Async to sync ordered delivery " & i
  Next
  value = ws.Receive()
  Assert IsEmpty(value), "Sync close after async queue"
  Assert cb.messages = 1 And cb.maxDepth = 1, "No duplicate or reentrant callbacks after stop"
  Set cb.socket = Nothing
  Set ws = stdWebSocket.Create(base & "/echo?test=1", 1000, native)
  Set cb = New WebSocketAsyncRecorder: cb.failOnMessage = True
  ws.StartListening cb
  ws.Send "fail"
  WaitMessages cb, 1
  Assert Not ws.Listening And ws.LastListenerError = "Deliberate listener failure", "Callback failure stops delivery"
  ws.Send "sync after failure"
  Assert ws.Receive() = "sync after failure", "Sync recovery after callback failure"
  Set other = New WebSocketAsyncRecorder
  ws.StartListening other
  ws.Send "async after sync"
  WaitMessages other, 1
  Assert other.lastData = "async after sync", "Switch sync to async"
  ws.StopListening
  Set cb = New WebSocketAsyncRecorder: Set cb.socket = ws: Set cb.replacement = other
  ws.StartListening cb
  ws.Send "replace"
  WaitMessages cb, 1
  ws.Send "new listener"
  WaitMessages other, 2
  Assert cb.messages = 1 And other.lastData = "new listener", "Replacement from callback"
  Set cb.replacement = Nothing: Set cb.socket = Nothing
  Set cb = New WebSocketAsyncRecorder: Set cb.socket = ws: Set cb.replacement = other: cb.failAfterReplacement = True
  ws.StartListening cb
  ws.Send "replace then throw"
  WaitMessages cb, 1
  Assert ws.Listening, "Stale callback error does not cancel replacement"
  ws.Send "replacement survived"
  WaitMessages other, 3
  Assert other.lastData = "replacement survived", "Replacement remains active after old callback error"
  Set cb.replacement = Nothing: Set cb.socket = Nothing
  Set cb = New WebSocketAsyncRecorder: Set cb.socket = ws: cb.disconnectOnMessage = True
  ws.StartListening cb
  ws.Send "disconnect"
  WaitMessages cb, 1
  Assert ws.State = NotConnected And cb.closes = 1 And cb.maxDepth = 1, "Disconnect from callback safely delivers Close"
  Set cb.socket = Nothing
  Set ws = stdWebSocket.Create(base & "/silent", 1000, native)
  Set cb = New WebSocketAsyncRecorder: Set cb.socket = ws: cb.disconnectOnOpen = True
  ws.StartListening cb
  Dim started As Double: started = Timer
  Do While cb.closes = 0 And Timer - started < 3: DoEvents: Loop
  Assert cb.opens = 1 And cb.closes = 1 And cb.maxDepth = 1 And ws.State = NotConnected, "Disconnect from Open"
  Set cb.socket = Nothing
  Set ws = stdWebSocket.Create(base & "/echo?test=1", 1000, native)
  Set cb = New WebSocketAsyncRecorder
  ws.StartListening cb
  ws.CloseSocket 3001, "local async close"
  Assert cb.closes = 1 And Not ws.Listening And ws.State = NotConnected, "Synchronous close while async receive pending"
  Set ws = stdWebSocket.Create(base & "/async-edge", 1000, native)
  Set cb = New WebSocketAsyncRecorder
  ws.StartListening cb
  WaitMessages cb, 3
  Assert cb.payloads(1) = String$(10000, "x") & vbNullChar & ChrW(&H20AC), "Async fragmented UTF8"
  Assert cb.payloads(2) = "", "Async empty text"
  value = cb.payloads(3)
  Assert VarType(value) = (vbArray Or vbByte) And LenB(value) = 0, "Async empty binary"
  ws.Disconnect
  Set ws = stdWebSocket.Create(base & "/trickle", 100, native)
  Set cb = New WebSocketAsyncRecorder
  ws.StartListening cb
  started = Timer
  Do While cb.errors = 0 And Timer - started < 3: DoEvents: Loop
  Assert cb.errors = 1 And ws.State = NotConnected And Not ws.Listening, "Async fragment deadline and Error cleanup"
  ws.Disconnect
  Set ws = stdWebSocket.Create(base & "/idle", 1000, native)
  Set cb = New WebSocketAsyncRecorder: Set cb.socket = ws: cb.stopOnOpen = True
  ws.StartListening cb
  started = Timer
  Do While Timer - started < 0.4: Loop
  started = Timer
  Do While cb.opens = 0 And Timer - started < 3: DoEvents: Loop
  Assert cb.opens = 1 And Not ws.Listening And ws.State = NotConnected, "Stop after peer close preserves unread queue"
  Set cb.socket = Nothing
  Set other = New WebSocketAsyncRecorder
  ws.StartListening other
  started = Timer
  Do While other.closes = 0 And Timer - started < 3: DoEvents: Loop
  Assert other.messages = 2 And other.closes = 1, "Resume async delivery of buffered terminal queue"
  On Error Resume Next
  Err.Clear
  ws.StartListening cb
  previous = Err.Number
  On Error GoTo failed
  Assert previous <> 0 And Not ws.Listening, "Fully consumed terminal queue cannot register a new listener"
  ws.Disconnect
  For i = 1 To 30
    Set ws = stdWebSocket.Create(base & "/echo?test=1", 1000, native)
    Set cb = New WebSocketAsyncRecorder
    ws.StartListening cb
    ws.StopListening
    ws.StartListening cb
    Set ws = Nothing
    previous = cb.calls
    DoEvents
    Assert cb.calls = previous, "Stale notifications after owner release"
    previous = destroyedRecorders
    Set cb = Nothing
    Assert destroyedRecorders = previous + 1, "Receiver/listener lifetime released"
  Next
  RunTests = "Passed " & checks & " asynchronous WebSocket assertions"
  Exit Function
failed:
  RunTests = "FAILED async after " & checks & ": " & Err.Description
  On Error Resume Next
  If Not ws Is Nothing Then ws.Disconnect
  If Not cb Is Nothing Then Set cb.socket = Nothing
End Function
