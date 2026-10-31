Attribute VB_Name = "stdWebSocketLocalTests"
Option Explicit
Private checks As Long
Private Sub Trace(ByVal value As String)
  Dim path As String: path = Environ$("STDWEBSOCKET_TRACE")
  If path <> "" Then
    Dim channel As Integer: channel = FreeFile
    Open path For Output As #channel
    Print #channel, value
    Close #channel
  End If
End Sub
Public Function Architecture() As String
#If Win64 Then
  Architecture = "x64"
#Else
  Architecture = "x86"
#End If
End Function
Private Sub Assert(ByVal condition As Boolean, ByVal message As String)
  If checks Mod 1000 = 0 Or checks < 15 Or checks > 10000 Then Trace CStr(checks) & ": " & message
  If Not condition Then Err.Raise 5, "WebSocket tests", message
  checks = checks + 1
End Sub
Public Function RunTests(ByVal base As String, ByVal refused As String, ByVal nativeLibrary As String, ByVal tls As String) As String
  On Error GoTo failed
  Trace "Starting synchronous tests"
  Dim ws As stdWebSocket, result As Variant, kind As EMessageType
  Set ws = stdWebSocket.Create(base & "/echo?test=1", 500, nativeLibrary)
  Assert ws.State = Connected, "Connected state"
  Assert ws.url = base & "/echo?test=1", "URL preserved"
  Dim recorder As New WebSocketRecorder
  Set recorder.socket = ws
  ws.Listeners.Add recorder
  Dim text As String
  text = "hello" & vbNullChar & ChrW(&H20AC) & ChrW(&HD83D) & ChrW(&HDE00)
  ws.Send text
  result = ws.Receive(kind)
  Assert result = text And kind = Utf8, "UTF8 explicit lengths"
  Assert recorder.opens = 1 And recorder.messages = 1 And recorder.calls = 2, "Events and listeners"
  Assert recorder.lastData = text, "Event payload"
  ws.Send ""
  Assert ws.Pump(), "Empty text pump"
  Assert recorder.lastData = "", "Empty text payload"
  Dim bytes() As Byte, returned() As Byte, i As Long
  ReDim bytes(7 To 10006)
  For i = 7 To 10006: bytes(i) = i Mod 256: Next
  ws.Send bytes, Binary
  result = ws.Receive(kind): returned = result
  Assert kind = Binary And UBound(returned) = 9999, "Exact binary length"
  For i = 0 To 9999
    Assert returned(i) = bytes(i + 7), "Binary byte " & i
  Next
  bytes = StrConv(vbNullString, vbFromUnicode)
  ws.Send bytes, Binary
  result = ws.Receive(kind)
  Assert kind = Binary And VarType(result) = (vbArray Or vbByte), "Empty binary type"
  Assert LenB(result) = 0, "Empty binary length"
  ws.Send "first", Utf8, True
  ws.Send "last", Utf8
  Assert ws.Receive() = "firstlast", "Outgoing fragments"
  ws.CloseSocket 3001, "done"
  Assert ws.State = NotConnected And recorder.closes = 1, "Local close"
  ws.Disconnect
  Assert recorder.closes = 1, "Idempotent cleanup"
  Set recorder.socket = Nothing
  Set ws = stdWebSocket.Create(base & "/fragment", 500, nativeLibrary)
  result = ws.Receive(kind)
  Assert result = String$(10000, "x") & vbNullChar & ChrW(&H20AC), "Incoming UTF8 fragments"
  result = ws.Receive()
  Assert IsEmpty(result) And ws.State = NotConnected, "Peer close"
  Set ws = stdWebSocket.Create(base & "/binary-fragment", 500, nativeLibrary)
  result = ws.Receive(kind): returned = result
  Assert kind = Binary And UBound(returned) = 5121, "Binary frame fragments exact size"
  Assert returned(5120) = 0 And returned(5121) = 255, "Binary final fragment"
  ws.Disconnect
  Set ws = stdWebSocket.Create(base & "/split-utf8", 500, nativeLibrary)
  Assert ws.Receive() = ChrW(&H20AC), "UTF8 split inside codepoint"
  ws.Disconnect
  Set ws = stdWebSocket.Create(base & "/silent", 200, nativeLibrary)
  Set recorder.socket = ws
  Dim previousCalls As Long: previousCalls = recorder.errors
  Dim started As Double: started = Timer
  On Error Resume Next
  result = ws.Receive()
  Dim errorNumber As Long: errorNumber = Err.Number
  On Error GoTo failed
  Assert errorNumber <> 0 And ws.State = NotConnected, "Receive timeout"
  Assert recorder.errors = previousCalls + 1, "Timeout Error event exactly once"
  Assert Timer - started < 3, "Bounded receive"
  Set recorder.socket = Nothing
  Dim endpoint As Variant
  For Each endpoint In Array("/trickle", "/broken-fragment", "/limit")
    Set ws = stdWebSocket.Create(base & endpoint, 3000, nativeLibrary)
    On Error Resume Next
    Err.Clear
    result = ws.Receive()
    errorNumber = Err.Number
    On Error GoTo failed
    Assert errorNumber <> 0 And ws.State = NotConnected, "Fragment failure cleanup " & endpoint
  Next
  For Each endpoint In Array("/reject", "/stall-handshake")
    On Error Resume Next
    Err.Clear
    Set ws = stdWebSocket.Create(base & endpoint, 200, nativeLibrary)
    errorNumber = Err.Number
    On Error GoTo failed
    Assert errorNumber <> 0, "Handshake rejection/deadline " & endpoint
  Next
  Set ws = stdWebSocket.Create(base & "/no-close", 200, nativeLibrary)
  started = Timer
  On Error Resume Next
  Err.Clear
  ws.CloseSocket
  errorNumber = Err.Number
  On Error GoTo failed
  Assert errorNumber <> 0 And ws.State = NotConnected, "Close timeout cleanup"
  Assert Timer - started < 3, "Bounded close"
  Set ws = stdWebSocket.Create(base & "/no-read", 200, nativeLibrary)
  ReDim bytes(0 To 16777215)
  started = Timer
  On Error Resume Next
  Err.Clear
  ws.Send bytes, Binary
  errorNumber = Err.Number
  On Error GoTo failed
  Assert errorNumber <> 0 And ws.State = NotConnected, "Send timeout cleanup"
  Assert Timer - started < 3, "Bounded send"
  On Error Resume Next
  Err.Clear
  Set ws = stdWebSocket.Create(tls & "/tls", 2000, nativeLibrary)
  errorNumber = Err.Number
  On Error GoTo failed
  Assert errorNumber <> 0, "Untrusted TLS certificate rejected"
  On Error Resume Next
  Set ws = stdWebSocket.Create(refused, 200, nativeLibrary)
  errorNumber = Err.Number
  On Error GoTo failed
  Assert errorNumber <> 0, "Connection refused"
  Dim invalid As Variant
  For Each invalid In Array("http://localhost/", "ws:///", "ws://localhost:0/", "ws://localhost:65536/", "ws://user@localhost/", "ws://localhost/#fragment")
    On Error Resume Next
    Err.Clear
    Set ws = stdWebSocket.Create(CStr(invalid), 200, nativeLibrary)
    errorNumber = Err.Number
    On Error GoTo failed
    Assert errorNumber <> 0, "Reject invalid URL " & invalid
  Next
  RunTests = "Passed " & checks & " local WebSocket assertions"
  Exit Function
failed:
  RunTests = "FAILED after " & checks & ": " & Err.Description
  On Error Resume Next
  If Not ws Is Nothing Then ws.Disconnect
End Function
