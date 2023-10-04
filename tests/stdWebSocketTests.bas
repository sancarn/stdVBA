Attribute VB_Name = "stdWebSocketTests"
'@lang VBA

Public Sub testAll()
  'Create socket
  Test.Topic "stdWebSocket"
  
  Dim vRet as variant

  Dim ws As stdWebSocketNew
  Set ws = stdWebSocketNew.Create("ws://vi-server.org/", 1939)
  Call ws.Send("hello world")
  vRet = ws.Receive()
  Test.Assert "Received text correct", vRet = "hello world"

  Dim b(1 To 2) As Byte
  b(1) = 10
  b(2) = 12
  Call ws.Send(b, Binary)
  Dim iType as EMessageType: vRet = ws.Receive(iType)
  
  Test.Assert "Binary received type == Binary", iType = EMessageType.Binary
  Test.Assert "Binary 1st byte check", vRet(0) = 10
  Test.Assert "Binary 2nd byte check", vRet(1) = 12
  Test.Assert "Binary size check", (Ubound(vRet) - lbound(vRet) + 1) = (Ubound(b) - lbound(b) + 1)

  'Partial and binary
  Call ws.Send("hello", , True)
  Call ws.Send(" world", , False)
  Test.Assert "Received text correct partial", ws.Receive() = "hello world"

  'Very large string 10 chars 1000 times = 10,000 char string response should be 10,008B long
  Dim sMessage As String: sMessage = ""
  For i = 1 To 999
    Call ws.Send("1234567890", , True)
    sMessage = sMessage & "1234567890"
  Next
  Call ws.Send("1234567890", , False)
  sMessage = "MESSAGE: " & sMessage & "1234567890"
  Dim sRet As String: sRet = ws.Receive()
  Test.Assert "Test very large binary string", sRet = sMessage

  'Ensure close
  Set ws = Nothing
End Sub




