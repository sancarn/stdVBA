VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdWebSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Spec:
'

'Functions implemented on this class:
'=============================================
'  CONSTRUCTORS
'    [X] Create

'
'  INSTANCE METHODS
'  Many methods were inspired by those in Ruby's Enumerable: https://ruby-doc.org/core-2.7.2/Enumerable.html
'    [ ] send(data)
'    [ ] close(code,reason)
'    [ ] disconnect()
'    [X] Get url() 

'  PROTECTED INSTANCE METHODS
'    [X] handleEvent()

'  EVENTS
'    [X] OnOpen(data)
'    [X] OnClose(data)
'    [X] OnError(data)
'    [X] OnMessage(data)
'    [X] EventRaised(name,data)
'  
'  WHAT WE WON'T DO:
'    with_index    'this can't be easily done, so instead implement methods like `forEach(cb,withIndex?=false)

'  TODO: Things we can't do (yet)
'    take <-- can't do this unless we implement IEnumVARIANT and call Next() method

Public Event EventRaised(ByVal sEventName as string, ByVal eventData as variant)
Public Event OnOpen(ByVal eventData as variant)
Public Event OnClose(ByVal eventData as variant)
Public Event OnError(ByVal eventData as variant)
Public Event OnMessage(ByVal eventData as variant)

Private ie as object
Private document as object
Private pURL as string
Public Listeners as Collection 'Collection<stdICallable>

'Creates and returns a new WebSocket object
'@constructor
'@param {String} url - The websocket url to track.
'@example stdWebSocket.Create("wss://echo.websocket.org/")
Public Function Create(ByVal url as string) as stdWebSocket
  set Create = new stdWebSocket
  Call Create.Init(url)
End Function
Public Sub Init(ByVal url as string)
  'Create IE object
  set ie = CreateObject("InternetExplorer.Application")

  'Navigate to empty page
  Call ie.navigate("about:<!DOCTYPE html><meta http-equiv='X-UA-Compatible' content='IE=edge'><body></body>")
  While ie.readyState < 4
    DoEvents
  Wend

  'Save document
  set document = ie.document

  'Bind browser helper objects
  with document.parentWindow
    set .vb_obj = Me
    .websocket_url = url
  end with

  'Script BHO possible alternative:
  'set .vb_event = stdCallback.CreateFromObjectMethod(me,"handleEvent")
  'vb_event('EventName',event)

  'Create and run javascript bindings to websocket
  Dim scpt as object: set scpt = document.createElement("script")
  scpt.txt = "window.ws = new WebSocket(websocket_url);" & _ 
    "ws.onopen = function(event){ vb_obj.handleEvent('Open', event); };" & _
    "ws.onclose = function(event){ vb_obj.handleEvent('Close', event); };" & _ 
    "ws.onerror = function(event){ vb_obj.handleEvent('Error', event); };" & _ 
    "ws.onmessage = function(event){ vb_obj.handleEvent('Message', event); };"
  Call document.body.appendChild(scpt)

  pURL = url
  set Listeners = new Collection
End Sub

'Gets the url of the websocket
Public Property Get url() as string
  url = pURL
End Property

'Sends data through the WebSocket
'@param {String} data - The data to send
Public Sub Send(ByVal data as variant)
  Call document.parentWindow.ws.send(data)
End Sub

'Closes the WebSocket connection
'@param {Long} Code - A numeric value indicating the status code explaining why the connection is being closed. If this parameter is not specified, a default value of 1005 is assumed. See the [list of status](https://developer.mozilla.org/en-US/docs/Web/API/CloseEvent#status_codes) code of CloseEvent for permitted values.
'@param {String} Reason - A human-readable string explaining why the connection is closing.
Public Sub Close(Optional ByVal code as long = 1000, Optional ByVal reason as string = "")
  Call document.parentWindow.ws.close(code,reason)
End Sub

'Closes and deletes the WebSocket, removing references so the object can be garbage collected
Public Sub Disconnect()
  Call Close()
  set ie = nothing
End Sub


'Handles socket events
'@protected
'@param {String} sEventName - The name of the event ("Open", "Close", "Error" or "Message")
'@param {Object} oEventData - Data to forward to event handler.
Public Sub handleEvent(ByVal sEventName as string, ByVal oEventData as object)
  RaiseEvent EventRaised(sEventName, oEventData)
  select case sEventName
    case "Open":    RaiseEvent OnOpen(oEventData)
    case "Close":   RaiseEvent OnClose(oEventData)
    case "Error":   RaiseEvent OnError(oEventData)
    case "Message": RaiseEvent OnMessage(oEventData)
  end select
  
  On Error Resume Next
    Dim listener as Object
    For each listener in Listeners
      Call listener.run(sEventName,oEventData)
    next
  On Error GoTo 0
End Sub

