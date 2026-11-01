VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "WebSocketRecorder"
Attribute VB_PredeclaredId = False
Option Explicit
Public WithEvents socket As stdWebSocket
Attribute socket.VB_VarHelpID = -1
Public opens As Long
Public messages As Long
Public closes As Long
Public calls As Long
Public errors As Long
Public lastData As Variant
Private Sub socket_OnOpen(ByVal eventData As Variant)
  opens = opens + 1
End Sub
Private Sub socket_OnMessage(ByVal eventData As Variant)
  messages = messages + 1
  lastData = eventData("data")
End Sub
Private Sub socket_OnClose(ByVal eventData As Variant)
  closes = closes + 1
End Sub
Private Sub socket_OnError(ByVal eventData As Variant)
  errors = errors + 1
End Sub
Public Sub Run(ByVal name As String, ByVal eventData As Object)
  calls = calls + 1
End Sub
