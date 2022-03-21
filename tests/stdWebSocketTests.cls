VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdWebSocketTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Sub testAll()
  'Create socket
  Test.Topic "stdWebSocket"
  
  Dim ws as stdWebSocket 
  Set ws = stdWebSocket.Create("ws://vi-server.org:1939/")
  Call ws.send("hello world")
  Test.Assert "Received text correct", ws.receive() = "MESSAGE: hello world"


  Call ws.send("hello world", utf8)
  Test.Assert "Received text correct utf8", ws.receive() = "MESSAGE: hello world"

  'Partial and binary
  Call ws.send("hello", , true)
  Call ws.send(" world", , false)
  Test.Assert "Received text correct partial", ws.receive() = "MESSAGE: hello world"

  'Very large string 10 chars 1000 times = 10,000 char string response should be 10,008B long
  Dim sMessage as string: sMessage = ""
  For i = 1 to 999
    Call ws.send("1234567890", , true)
    sMessage = sMessage & "1234567890"
  next
  Call ws.send("1234567890", , false)
  sMessage = "MESSAGE: " & sMessage & "1234567890"
  Dim sRet as string: sRet = ws.receive()
  Test.Assert "Test very large binary string", sRet = sMessage

  'Ensure close
  set ws = nothing
End Sub



