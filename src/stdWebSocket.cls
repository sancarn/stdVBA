VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdWebSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Spec: 
'

'Functions implemented on this class:
'=============================================
'  CONSTRUCTORS
'    [X] Create

'
'  INSTANCE METHODS
'  Many methods were inspired by those in Ruby's Enumerable: https://ruby-doc.org/core-2.7.2/Enumerable.html
'    [X] send(data)
'    [X] close(code,reason)
'    [X] disconnect()
'    [X] Get url()

'  PROTECTED INSTANCE METHODS
'    [X] handleEvent()

'  EVENTS
'    [X] OnOpen(data)
'    [X] OnClose(data)
'    [X] OnError(data)
'    [X] OnMessage(data)
'    [X] EventRaised(name,data)

'Event raised when any of the below events are raised
'@param sEventName as "Open"|"Close"|"Error"|"Message" - The name of the event.
'@param eventData - Data to forward to event handler.
Public Event EventRaised(ByVal sEventName As String, ByVal eventData As Variant)

'Event raised when the websocket is opened
'@param eventData - Data to forward to event handler.
Public Event OnOpen(ByVal eventData As Variant)

'Event raised when the websocket is closed
'@param eventData - Data to forward to event handler.
Public Event OnClose(ByVal eventData As Variant)

'Event raised when the websocket has an error
'@param eventData - Data to forward to event handler.
Public Event OnError(ByVal eventData As Variant)

'Event raised when the websocket receives a message
'@param eventData - Data to forward to event handler.
Public Event OnMessage(ByVal eventData As Variant)

Private ie As Object
Private document As Object
Private pURL As String
Public Listeners As Collection 'Collection<stdICallable>

'Creates and returns a new WebSocket object
'@constructor
'@param url - The websocket url to track.
'@returns - The new WebSocket object.
'@example stdWebSocket.Create("wss://echo.websocket.org/")
Public Function Create(ByVal url As String) As stdWebSocket
  Set Create = New stdWebSocket
  Call Create.protInit(url)
End Function

'Initializes the object
'@protected
'@param url - The websocket url to track.
'@example `ws.protInit("wss://echo.websocket.org/")`
Public Sub protInit(ByVal url As String)
  'Create IE object
  Set ie = CreateObject("InternetExplorer.Application")

  'Navigate to empty page
  Call ie.navigate("about:<!DOCTYPE html><meta http-equiv='X-UA-Compatible' content='IE=edge'><body></body>")
  While ie.readyState < 4
    DoEvents
  Wend

  'Save document
  Set document = ie.document
  
  'Create bind function on document object
  Dim scpt As Object
  Set scpt = document.createElement("script")
  scpt.text = "document.bind = function(name,data){window[name]=data}"
  Call document.body.appendChild(scpt)
  
  'Bind browser helper objects
  With document
    Call .Bind("vb_obj", Me)
    Call .Bind("websocket_url", url)
  End With

  'Script BHO possible alternative:
  'set .vb_event = stdCallback.CreateFromObjectMethod(me,"handleEvent")
  'vb_event('EventName',event)
  
  pURL = url
  Set Listeners = New Collection
  
  'Create and run javascript bindings to websocket
  Set scpt = document.createElement("script")
  scpt.text = "window.ws = new WebSocket(websocket_url);" & _
    "ws.onopen = function(event){ vb_obj.handleEvent('Open', event); };" & _
    "ws.onclose = function(event){ vb_obj.handleEvent('Close', event); };" & _
    "ws.onerror = function(event){ vb_obj.handleEvent('Error', event); };" & _
    "ws.onmessage = function(event){ vb_obj.handleEvent('Message', event); };" & _
    "document.ws = window.ws;"
  Call document.body.appendChild(scpt)
End Sub

'Gets the url of the websocket
'@returns - The url of the websocket
Public Property Get url() As String
  url = pURL
End Property

'Sends data through the WebSocket
'@param data - The data to send
Public Sub Send(ByVal data As Variant)
  Call document.ws.Send(data)
End Sub

'Closes the WebSocket connection
'@param Code - A numeric value indicating the status code explaining why the connection is being closed. If this parameter is not specified, a default value of 1005 is assumed. See the [list of status](https://developer.mozilla.org/en-US/docs/Web/API/CloseEvent#status_codes) code of CloseEvent for permitted values.
'@param Reason - A human-readable string explaining why the connection is closing.
Public Sub CloseSocket(Optional ByVal Code As Long = 1000, Optional ByVal reason As String = "")
  Call document.ws.Close(Code, reason)
End Sub

'Closes and deletes the WebSocket, removing references so the object can be garbage collected
Public Sub Disconnect()
  Call CloseSocket
  Call document.Bind("vb_obj", 0)
  Call ie.Quit
End Sub


'Handles socket events
'@protected
'@param sEventName - The name of the event ("Open", "Close", "Error" or "Message")
'@param oEventData - Data to forward to event handler.
Public Sub handleEvent(ByVal sEventName As String, ByVal oEventData As Object)
  RaiseEvent EventRaised(sEventName, oEventData)
  Select Case sEventName
    Case "Open":    RaiseEvent OnOpen(oEventData)
    Case "Close":   RaiseEvent OnClose(oEventData)
    Case "Error":   RaiseEvent OnError(oEventData)
    Case "Message": RaiseEvent OnMessage(oEventData)
  End Select
  
  On Error Resume Next
    Dim listener As Object
    For Each listener In Listeners
      Call listener.Run(sEventName, oEventData)
    Next
  On Error GoTo 0
End Sub

