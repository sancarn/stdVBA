VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdWebSocketTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents x As stdWebSocket
Attribute x.VB_VarHelpID = -1
Dim check As Collection
Dim finished As Boolean

Public Sub testAll()
  'Create socket
  Test.Topic "stdWebSocket"
  finished = False
  Set x = stdWebSocket.Create("wss://echo.websocket.org")
  While Not finished
    DoEvents
  Wend
End Sub

Private Sub x_OnOpen(ByVal oEventData As Variant)
  Test.Assert "OnOpen Event Type", oEventData.Type = "open"
  Test.Assert "OnOpen Event Type Name", TypeName(oEventData) = "DispDOMEvent"
  Call x.Send("Hello world!")
End Sub
Private Sub x_OnMessage(ByVal oEventData As Variant)
  Test.Assert "OnMessage Event Type", oEventData.Type = "message"
  Test.Assert "OnMessage Event Type Name", TypeName(oEventData) = "DispDOMMessageEvent"
  Test.Assert "OnMessage Event Data", oEventData.data = "Hello world!"
  Call x.CloseSocket
End Sub
Private Sub x_OnClose(ByVal oEventData As Variant)
  Test.Assert "OnClose Type", oEventData.Type = "close"
  Test.Assert "OnClose Type Name", TypeName(oEventData) = "DispDOMCloseEvent"
  Call x.Disconnect
  finished = True
End Sub
Private Sub x_OnError(ByVal oEventData As Variant)
  
End Sub

