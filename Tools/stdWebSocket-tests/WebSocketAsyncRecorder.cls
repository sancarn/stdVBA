VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "WebSocketAsyncRecorder"
Attribute VB_PredeclaredId = False
Option Explicit
Implements stdICallable
#If VBA7 Then
Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
#Else
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
#End If
Public ownerThread As Long
Public wrongThread As Long
Public calls As Long
Public opens As Long
Public messages As Long
Public closes As Long
Public errors As Long
Public lastData As Variant
Public payloads As Collection
Public socket As stdWebSocket
Public stopOnMessage As Boolean
Public disconnectOnMessage As Boolean
Public failOnMessage As Boolean
Public failAfterReplacement As Boolean
Public disconnectOnOpen As Boolean
Public stopOnOpen As Boolean
Public replacement As stdICallable
Public depth As Long
Public maxDepth As Long
Public readConflict As Long
Private Sub Class_Initialize()
  ownerThread = GetCurrentThreadId()
  Set payloads = New Collection
End Sub
Private Sub Class_Terminate()
  stdWebSocketAsyncTests.destroyedRecorders = stdWebSocketAsyncTests.destroyedRecorders + 1
End Sub
Private Function stdICallable_Run(ParamArray params() As Variant) As Variant
  Handle CStr(params(0)), params(1)
End Function
Private Function stdICallable_RunEx(ByVal params As Variant) As Variant
  Handle CStr(params(0)), params(1)
End Function
Private Function stdICallable_Bind(ParamArray params() As Variant) As stdICallable
  Err.Raise 5, , "Unused"
End Function
Private Function stdICallable_SendMessage(ByVal sMessage As String, ByRef success As Boolean, ByVal params As Variant) As Variant
  success = False
End Function
Private Sub Handle(ByVal name As String, ByVal eventData As Object)
  If GetCurrentThreadId() <> ownerThread Then wrongThread = wrongThread + 1
  calls = calls + 1
  depth = depth + 1
  If depth > maxDepth Then maxDepth = depth
  Select Case name
    Case "Open"
      opens = opens + 1
      If disconnectOnOpen Then socket.Disconnect
      If stopOnOpen Then socket.StopListening
    Case "Message"
      messages = messages + 1: lastData = eventData("data")
      payloads.Add lastData
      If failOnMessage Then Err.Raise 5, , "Deliberate listener failure"
      If Not socket Is Nothing Then
        On Error Resume Next
        Dim value As Variant: value = socket.Receive()
        readConflict = Err.Number
        Err.Clear
        On Error GoTo 0
        If stopOnMessage Then socket.StopListening
        If Not replacement Is Nothing Then socket.StartListening replacement
        If failAfterReplacement Then Err.Raise 5, , "Old listener failed after replacement"
        If disconnectOnMessage Then socket.Disconnect
      End If
      DoEvents
    Case "Close": closes = closes + 1
    Case "Error": errors = errors + 1
  End Select
  depth = depth - 1
End Sub
