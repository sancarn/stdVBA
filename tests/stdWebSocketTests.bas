Attribute VB_Name = "stdWebSocketTests"

Dim WithEvents x as stdWebSocket
Dim check as collection
Sub testAll()
  'Create socket
  set x = stdWebSocket.Create("wss://echo.websocket.org")
End Sub

Sub x_OnOpen(oEventData)
  Debug.Assert False
  Call x.send("Hello world!")
End Sub
Sub x_OnMessage(oEventData)
  Debug.Assert False 'Event.Data
  Call x.Close()
End Sub
Sub x_OnClose(oEventData)
  Debug.Assert False 'Event.Data
  Call x.Disconnect()
End Sub
Sub x_OnError(oEventData)
  
End Sub
