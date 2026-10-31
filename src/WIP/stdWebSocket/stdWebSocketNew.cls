VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdWebSocketNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@module
'@description Native Windows WebSocket client using WinHTTP, replacing Internet Explorer (#139, #35).
'@requires Windows 8 or later, stdICallable, and stdWebSocketNative.dll built for the Office process architecture.
'@remark StartListening delivers to a typed callable through an owner-thread message window. Pump/Receive provide synchronous events when no async listener owns reads. No VBA runs on WinHTTP or receive worker threads.
'@devNote WinHTTP supplies HTTP upgrade, framing and TLS. Wasabi was reviewed as a design reference: https://github.com/vbacollective/wasabi/blob/main/package/Wasabi.bas

Option Explicit

#if VBA7 then
  Private Declare PtrSafe Function LoadLibraryW Lib "kernel32" (ByVal path As LongPtr) As LongPtr
  Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal library As LongPtr) As Long
  Private Declare PtrSafe Function stdWebSocketConnect Lib "stdWebSocketNative.dll" (ByVal host As LongPtr, ByVal port As Long, ByVal path As LongPtr, ByVal flags As Long, ByVal timeout As Long, ByRef socket As LongPtr) As Long
  Private Declare PtrSafe Sub stdWebSocketDisconnect Lib "stdWebSocketNative.dll" (ByRef socket As LongPtr)
  Private Declare PtrSafe Function stdWebSocketListen Lib "stdWebSocketNative.dll" (ByVal socket As LongPtr, ByRef receiver As Variant) As Long
  Private Declare PtrSafe Function stdWebSocketUnlisten Lib "stdWebSocketNative.dll" (ByVal socket As LongPtr, ByRef receiver As Variant) As Long
  Private Declare PtrSafe Function stdWebSocketListening Lib "stdWebSocketNative.dll" (ByVal socket As LongPtr) As Long
  Private Declare PtrSafe Function stdWebSocketTerminal Lib "stdWebSocketNative.dll" (ByVal socket As LongPtr) As Long
  Private Declare PtrSafe Function stdWebSocketNext Lib "stdWebSocketNative.dll" (ByVal socket As LongPtr, ByRef data As Variant, ByRef kind As Long, ByRef code As Long) As Long
  Private Declare PtrSafe Function stdWebSocketOperation Lib "stdWebSocketNative.dll" (ByRef socket As LongPtr, ByVal operation As Long, ByVal buffer As LongPtr, ByVal length As Long, ByRef bytesRead As Long, ByRef bufferType As Long, ByVal timeout As Long, ByVal closeCode As Integer) As Long


  'UTF8
  Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cchMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long ) As Long
  Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr ) As Long

#else
  Private Declare Function LoadLibraryW Lib "kernel32" (ByVal path As Long) As Long
  Private Declare Function FreeLibrary Lib "kernel32" (ByVal library As Long) As Long
  Private Declare Function stdWebSocketConnect Lib "stdWebSocketNative.dll" (ByVal host As Long, ByVal port As Long, ByVal path As Long, ByVal flags As Long, ByVal timeout As Long, ByRef socket As Long) As Long
  Private Declare Sub stdWebSocketDisconnect Lib "stdWebSocketNative.dll" (ByRef socket As Long)
  Private Declare Function stdWebSocketListen Lib "stdWebSocketNative.dll" (ByVal socket As Long, ByRef receiver As Variant) As Long
  Private Declare Function stdWebSocketUnlisten Lib "stdWebSocketNative.dll" (ByVal socket As Long, ByRef receiver As Variant) As Long
  Private Declare Function stdWebSocketListening Lib "stdWebSocketNative.dll" (ByVal socket As Long) As Long
  Private Declare Function stdWebSocketTerminal Lib "stdWebSocketNative.dll" (ByVal socket As Long) As Long
  Private Declare Function stdWebSocketNext Lib "stdWebSocketNative.dll" (ByVal socket As Long, ByRef data As Variant, ByRef kind As Long, ByRef code As Long) As Long
  Private Declare Function stdWebSocketOperation Lib "stdWebSocketNative.dll" (ByRef socket As Long, ByVal operation As Long, ByVal buffer As Long, ByVal length As Long, ByRef bytesRead As Long, ByRef bufferType As Long, ByVal timeout As Long, ByVal closeCode As Integer) As Long


  'UTF8
  Private Declare Function MultiByteToWideChar Lib "kernel32" ( ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
  Private Declare Function WideCharToMultiByte Lib "kernel32" ( ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long ) As Long

#end if

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

Public Enum EMessageType
  Binary
  Utf8
End Enum
Public Enum EWebSocketState
  NotConnected
  Connecting
  Connected
  Disconnecting
End Enum
Private Type TThis
#If VBA7 Then
  library As LongPtr
  hWebSocket As LongPtr
#Else
  library As Long
  hWebSocket As Long
#End If
  url As String
  State As EWebSocketState
  timeout As Long
  pumping As Boolean
  opened As Boolean
  sendingFragment As Boolean
  sendingType As EMessageType
  receiver As stdWebSocketNew
  listener As stdICallable
  isReceiver As Boolean
  lastListenerError As String
  callbackActive As Boolean
  pendingClose As Boolean
  pendingCode As Long
  pendingReason As String
End Type
Private This As TThis
Public Listeners As Collection

'Creates a connected native Windows WebSocket (Windows 8 or later).
'@constructor
'@param url - ws:// or wss:// URL; ASCII/Punycode host, optional port and path/query.
'@param nativeLibraryPath - Path to the matching Office-bitness stdWebSocketNative.dll; see native/build.ps1.
'@param timeoutMs - Positive network timeout in milliseconds.
'@returns - The new WebSocket object.
'@remark Use StartListening for deferred owner-thread delivery, or Pump/Receive for synchronous reads and events. Open is deferred so handlers can be attached after Create.
'@example ```vb
'Dim ws As stdWebSocketNew
'Set ws = stdWebSocketNew.Create("ws://127.0.0.1:8080/echo", 5000, "C:\libs\stdWebSocketNative.dll")
'ws.Send "hello"
'Debug.Print ws.Receive()
'ws.CloseSocket
'```
Public Function Create(ByVal url As String, Optional ByVal timeoutMs As Long = 5000, Optional ByVal nativeLibraryPath As String = "stdWebSocketNative.dll") As stdWebSocketNew
  Set Create = New stdWebSocketNew
  Create.protInit url, timeoutMs, nativeLibraryPath
End Function

'Initializes the object and completes the upgrade with normal TLS certificate validation.
'@protected
'@param url - WebSocket URL.
'@param nativeLibraryPath - Path to the matching Office-bitness stdWebSocketNative.dll; see native/build.ps1.
'@param timeoutMs - Positive per-operation network timeout in milliseconds.
'@example ```vb
'Dim ws As New stdWebSocketNew
'ws.protInit "ws://127.0.0.1:8080/echo", 5000, "C:\libs\stdWebSocketNative.dll"
'```
Public Sub protInit(ByVal url As String, Optional ByVal timeoutMs As Long = 5000, Optional ByVal nativeLibraryPath As String = "stdWebSocketNative.dll")
  If This.State <> NotConnected Then Err.Raise 5, "stdWebSocketNew", "Already connected"
  If timeoutMs < 1 Then Err.Raise 5, "stdWebSocketNew", "Timeout must be positive"
  Dim host As String, path As String, port As Long, flags As Long
  ParseURL url, host, port, path, flags
  If Listeners Is Nothing Then Set Listeners = New Collection
  This.url = url
  This.timeout = timeoutMs
  This.State = Connecting
  On Error GoTo failed
  If This.library = 0 Then This.library = LoadLibraryW(StrPtr(nativeLibraryPath))
  If This.library = 0 Then Err.Raise 53, "stdWebSocketNew", "Build/load stdWebSocketNative.dll for your Office architecture; pass its absolute path as nativeLibraryPath"
  Dim result As Long
  result = stdWebSocketConnect(StrPtr(host), port, StrPtr(path), flags, timeoutMs, This.hWebSocket)
  If result <> 0 Then Err.Raise vbObjectError + 139, "stdWebSocketNew.protInit", "Connect failed: " & result
  This.State = Connected
  Exit Sub
failed:
  Dim description As String: description = Err.Description
  ReleaseHandles
  Err.Raise vbObjectError + 139, "stdWebSocketNew.protInit", description
End Sub

'Gets the original WebSocket URL.
'@returns - The URL supplied to Create.
Public Property Get url() As String
  url = This.url
End Property
'Gets the current connection state.
'@returns - Connection state, including a terminal state detected by the async reader before the notification is delivered.
Public Property Get State() As EWebSocketState
  State = This.State
  If This.hWebSocket <> 0 And Not This.isReceiver Then
    If stdWebSocketTerminal(This.hWebSocket) <> 0 Then State = NotConnected
  End If
End Property

'Registers an asynchronous receive listener, returning without waiting for data.
'@param listener as stdICallable<(name as String, eventData as Dictionary)=>Variant> - Receives Open, Message, Close and Error on the registering Office thread.
'@remark Delivery requires the Office thread to pump Windows messages; it never interrupts a busy macro. This listener exclusively owns reads until StopListening. Each registration receives a deferred Open. Legacy events/Listeners remain the synchronous API. The class is non-exposed, matching stdWebView, so it can publish the non-exposed stdICallable type.
'@remark StopListening or Disconnect before VBE Reset, End, project unload or forced Office shutdown. Those paths can skip VBA destructors and are unsupported with a live subscription. Ordinary owner release is supported unless the supplied callable itself retains the socket, creating a user-owned cycle.
'@remark Queue limits are 32 messages and 16 MiB queued data, with up to one additional 16 MiB message under assembly/backpressure. Send remains synchronous and can block up to its deadline.
'@remark A stopped listener can be replaced to drain buffered messages and a pending terminal notification after peer close; registration is rejected after the terminal queue has been fully consumed.
Public Sub StartListening(ByVal listener As stdICallable)
  If listener Is Nothing Then Err.Raise 5, "stdWebSocketNew.StartListening", "Listener is Nothing"
  If This.pumping Or This.isReceiver Then Err.Raise 5, , "Read operation in progress"
  If This.State <> Connected Or This.hWebSocket = 0 Then Err.Raise 5, , "Not connected"
  Dim receiver As New stdWebSocketNew
  receiver.protInitReceiver listener
  receiver.protContext = This.hWebSocket
  Dim argument As Variant: Set argument = receiver
  Dim result As Long: result = stdWebSocketListen(This.hWebSocket, argument)
  If result <> 0 Then Err.Raise vbObjectError + 139, "stdWebSocketNew.StartListening", "Listener registration failed: " & result
  Set This.receiver = receiver
End Sub

'Stops asynchronous delivery without discarding unread messages or closing the socket.
'@remark A pending native read continues into the bounded queue. Subsequent Receive consumes that same queue, avoiding competing reads or duplicate delivery. Safe inside a listener; queued notifications from old registrations are invalidated.
Public Sub StopListening()
  If This.hWebSocket <> 0 Then
    Dim argument As Variant
    Dim result As Long: result = stdWebSocketListen(This.hWebSocket, argument)
    If result <> 0 Then Err.Raise vbObjectError + 139, , "StopListening failed: " & result
  End If
  If Not This.receiver Is Nothing Then This.receiver.protClearListener
End Sub

'Gets whether the asynchronous listener currently owns reads.
'@returns - True while a listener is registered.
Public Property Get Listening() As Boolean
  If This.hWebSocket <> 0 Then Listening = stdWebSocketListening(This.hWebSocket) <> 0
End Property

'Gets the last VBA error raised by the asynchronous listener or its decoder.
'@returns - Error description, or an empty string. A callback error stops delivery and leaves unread messages available to Receive.
Public Property Get LastListenerError() As String
  If Not This.receiver Is Nothing Then LastListenerError = This.receiver.protListenerError
End Property

'Initializes a separate receiver that does not retain the socket owner.
'@protected
'@param listener - Typed callable retained for this registration.
Public Sub protInitReceiver(ByVal listener As stdICallable)
  This.isReceiver = True
  Set This.listener = listener
End Sub

'Sets the borrowed native context of the bridge receiver. Never call directly.
'@protected
'@param value - Native context, valid only while the adapter retains this receiver.
#If VBA7 Then
Public Property Let protContext(ByVal value As LongPtr)
#Else
Public Property Let protContext(ByVal value As Long)
#End If
  This.hWebSocket = value
End Property

'Releases the typed listener reference after delivery is stopped.
'@protected
Public Sub protClearListener()
  Set This.listener = Nothing
End Sub

'Gets a bridge receiver error for its owning socket.
'@protected
'@returns - Last callback/decoding error.
Public Property Get protListenerError() As String
  protListenerError = This.lastListenerError
End Property

'Native message-window entry point, invoked through retained IDispatch on its owner thread.
'@protected
'@remark Dispatches at most one queued message. The native adapter prevents nested dispatch even when user code runs DoEvents. No WinHTTP worker invokes this member.
Public Sub protAsyncDispatch()
  If Not This.isReceiver Or This.listener Is Nothing Then Exit Sub
  On Error GoTo failed
  If Not This.opened Then
    This.opened = True
    AsyncEmit "Open"
    Exit Sub
  End If
  Dim data As Variant, kind As Long, code As Long, result As Long
  result = stdWebSocketNext(This.hWebSocket, data, kind, code)
  If result = 259 Then Exit Sub
  If result <> 0 Then Err.Raise vbObjectError + 139, , "Queue read failed: " & result
  Dim bytes() As Byte, value As Variant
  Select Case kind
    Case 0
      AsyncEmit "Message", data
    Case 2
      bytes = data: value = Decode(bytes, ByteCount(bytes))
      AsyncEmit "Message", value
    Case 4
      bytes = data
      AsyncEmit "Close", Empty, code, Decode(bytes, ByteCount(bytes))
      StopReceiver
    Case 5
      AsyncEmit "Error", "Receive failed: " & code
      StopReceiver
    Case Else
      Err.Raise 5, , "Invalid queue message type"
  End Select
  Exit Sub
failed:
  This.lastListenerError = Err.Description
  On Error Resume Next
  StopReceiver
End Sub

Private Sub StopReceiver()
  Dim argument As Variant: Set argument = Me
  Dim result As Long: result = stdWebSocketUnlisten(This.hWebSocket, argument)
  Set This.listener = Nothing
End Sub

Private Sub AsyncEmit(ByVal name As String, Optional ByVal data As Variant, Optional ByVal code As Long, Optional ByVal reason As String)
  Dim eventData As Object: Set eventData = CreateEventData(name, data, code, reason)
  Dim listener As stdICallable: Set listener = This.listener
  If Not listener Is Nothing Then
    This.callbackActive = True
    On Error GoTo failed
    listener.Run name, eventData
    This.callbackActive = False
    If This.pendingClose Then
      This.pendingClose = False
      Set eventData = CreateEventData("Close", Empty, This.pendingCode, This.pendingReason)
      listener.Run "Close", eventData
      Set This.listener = Nothing
    End If
  End If
  Exit Sub
failed:
  This.callbackActive = False
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'Delivers a local close after cleanup, deferring it until the active user callback returns.
'@protected
'@param code - Local close status.
'@param reason - Local close reason.
Public Sub protLocalClose(ByVal code As Long, ByVal reason As String)
  If This.callbackActive Then
    This.pendingClose = True: This.pendingCode = code: This.pendingReason = reason
  Else
    On Error GoTo failed
    AsyncEmit "Close", Empty, code, reason
    Set This.listener = Nothing
  End If
  Exit Sub
failed:
  This.lastListenerError = Err.Description
  Set This.listener = Nothing
End Sub

'Sends text or binary data, optionally as a message fragment.
'@param data - String for Utf8 or Byte() for Binary; empty and embedded nulls are preserved.
'@param iType - Utf8 or Binary.
'@param bPartialSend - True until the final fragment. Fragment types must match.
'@remark Send always waits for native completion, including when an async reader is active. Network backpressure can block the caller until the configured deadline; write failures raise a synchronous VBA error.
Public Sub Send(ByVal data As Variant, Optional ByVal iType As EMessageType = Utf8, Optional ByVal bPartialSend As Boolean = False)
  If This.State <> Connected Then Err.Raise 5, "stdWebSocketNew.Send", "Not connected"
  If iType <> Binary And iType <> Utf8 Then Err.Raise 5, "stdWebSocketNew.Send", "Invalid message type"
  If This.sendingFragment And iType <> This.sendingType Then Err.Raise 5, "stdWebSocketNew.Send", "Fragment type changed"
  Dim bytes() As Byte, n As Long, kind As Long
  If iType = Utf8 Then
    bytes = Encode(CStr(data))
    kind = 2
  Else
    If VarType(data) <> (vbArray Or vbByte) Then Err.Raise 5, "stdWebSocketNew.Send", "Expected Byte()"
    bytes = data
  End If
  n = ByteCount(bytes)
  If bPartialSend Then kind = kind + 1
  Dim result As Long
  If n = 0 Then
    result = stdWebSocketOperation(This.hWebSocket, 1, 0, 0, n, kind, This.timeout, 0)
  Else
    result = stdWebSocketOperation(This.hWebSocket, 1, VarPtr(bytes(LBound(bytes))), n, n, kind, This.timeout, 0)
  End If
  If result <> 0 Then NativeFailure "Send", result
  This.sendingFragment = bPartialSend
  This.sendingType = iType
End Sub

'Receives one complete message and delivers Open/Message/Close events and Listeners synchronously.
'@param iType - Receives the message type.
'@remark Rejected while Listening is True. After StopListening it consumes the same bounded native queue, preserving unread messages and exclusive ownership.
'@returns - String, Byte(), or Empty when the peer closes. Empty binary is a zero-length Byte().
'@remark A native deadline cancels a stalled receive; a total message deadline and 16 MiB limit apply. Timeout/transport errors disconnect and raise Error then a VBA error. No background VBA callbacks or DoEvents are used.
Public Function Receive(Optional ByRef iType As EMessageType) As Variant
  If Listening Then Err.Raise 5, "stdWebSocketNew.Receive", "Async listener owns reads; call StopListening first"
  If This.pumping Then Err.Raise 5, "stdWebSocketNew.Receive", "Reentrant receive"
  If This.State <> Connected Then Err.Raise 5, "stdWebSocketNew.Receive", "Not connected"
  This.pumping = True
  On Error GoTo failed
  If Not This.opened Then
    This.opened = True
    Emit "Open"
  End If
  If This.State <> Connected Then GoTo done
  Dim chunk(0 To 4095) As Byte, bytes() As Byte, n As Long, total As Long, capacity As Long
  Dim kind As Long, firstKind As Long, result As Long, i As Long
  Dim started As Double: started = Timer
  firstKind = -1
  Do
    Dim elapsed As Double: elapsed = Timer - started
    If elapsed < 0 Then elapsed = elapsed + 86400
    Dim remaining As Long: remaining = This.timeout - CLng(elapsed * 1000)
    If remaining < 1 Then Err.Raise 5, , "Message deadline exceeded"
    result = stdWebSocketOperation(This.hWebSocket, 0, VarPtr(chunk(0)), 4096, n, kind, remaining, 0)
    If result <> 0 Then Err.Raise vbObjectError + 139, "stdWebSocketNew.Receive", "Receive failed: " & result
    If kind = 4 Then
      Dim code As Long, reason(0 To 122) As Byte, reasonSize As Long
      result = stdWebSocketOperation(This.hWebSocket, 3, VarPtr(reason(0)), 123, reasonSize, code, remaining, 0)
      If result <> 0 Then Err.Raise vbObjectError + 139, , "Close status failed: " & result
      Dim closeReason As String: closeReason = Decode(reason, reasonSize)
      Dim closeType As Long
      result = stdWebSocketOperation(This.hWebSocket, 2, 0, 0, n, closeType, remaining, CInt(code))
      If result <> 0 Then Err.Raise vbObjectError + 139, , "Close acknowledgement failed: " & result
      ReleaseHandles
      Emit "Close", Empty, CLng(code), closeReason
      GoTo done
    End If
    If kind < 0 Or kind > 3 Then Err.Raise 5, , "Invalid buffer type"
    If firstKind = -1 Then firstKind = kind \ 2
    If kind \ 2 <> firstKind Then Err.Raise 5, , "Fragment type changed"
    If n < 0 Or n > 4096 Or total > 16777216 - n Then Err.Raise 5, , "Message exceeds 16 MiB"
    If n > 0 Then
      If total + n > capacity Then
        If capacity = 0 Then capacity = 4096
        Do While capacity < total + n: capacity = capacity * 2: Loop
        ReDim Preserve bytes(0 To capacity - 1)
      End If
      For i = 0 To n - 1: bytes(total + i) = chunk(i): Next
      total = total + n
    End If
    If kind = 0 Or kind = 2 Then Exit Do
  Loop
  iType = IIf(firstKind = 0, Binary, Utf8)
  If iType = Utf8 Then
    Receive = Decode(bytes, total)
  Else
    If total = 0 Then
      bytes = StrConv(vbNullString, vbFromUnicode)
    Else
      ReDim Preserve bytes(0 To total - 1)
    End If
    Receive = bytes
  End If
  Emit "Message", Receive
done:
  This.pumping = False
  Exit Function
failed:
  Dim description As String: description = Err.Description
  ReleaseHandles
  This.pumping = False
  Emit "Error", description
  Err.Raise vbObjectError + 139, "stdWebSocketNew.Receive", description
End Function

'Pumps one complete message through the existing events and Listeners.
'@returns - True for a message (including empty messages), False for peer close.
'@remark This is a bounded blocking pump, not a readiness poll. Open is delivered before waiting for a message.
Public Function Pump() As Boolean
  Dim value As Variant: value = Receive()
  Pump = Not IsEmpty(value)
End Function

'Closes the WebSocket and releases all handles even if the handshake fails.
'@param Code - 1000 or application close status 3000 through 4999.
'@param reason - UTF8 reason, at most 123 bytes.
Public Sub CloseSocket(Optional ByVal Code As Long = 1000, Optional ByVal reason As String = "")
  If Code <> 1000 And (Code < 3000 Or Code > 4999) Then Err.Raise 5, , "Invalid close code"
  Dim bytes() As Byte, n As Long, result As Long, kind As Long
  bytes = Encode(reason): n = ByteCount(bytes)
  If n > 123 Then Err.Raise 5, , "Close reason exceeds 123 bytes"
  If This.State = NotConnected Then Exit Sub
  This.State = Disconnecting
  If n = 0 Then
    result = stdWebSocketOperation(This.hWebSocket, 2, 0, 0, n, kind, This.timeout, CInt(Code))
  Else
    result = stdWebSocketOperation(This.hWebSocket, 2, VarPtr(bytes(0)), n, n, kind, This.timeout, CInt(Code))
  End If
  Dim receiver As stdWebSocketNew: Set receiver = This.receiver
  ReleaseHandles
  If result <> 0 Then NativeFailure "Close", result
  Emit "Close", Empty, Code, reason
  If Not receiver Is Nothing Then receiver.protLocalClose Code, reason
End Sub

'Releases the connection immediately without waiting for a close handshake. Safe to repeat.
Public Sub Disconnect()
  Dim active As Boolean: active = This.State <> NotConnected
  Dim receiver As stdWebSocketNew: Set receiver = This.receiver
  ReleaseHandles
  If active Then Emit "Close", Empty, 1006, "Disconnected"
  If active And Not receiver Is Nothing Then receiver.protLocalClose 1006, "Disconnected"
End Sub

'Handles an event on the calling VBA thread, forwarding the same data to events and Listeners.
'@protected
'@param sEventName - Open, Close, Error or Message.
'@param oEventData - Dictionary event object; Message has data, Close has code/reason, Error has message.
Public Sub handleEvent(ByVal sEventName As String, ByVal oEventData As Object)
  RaiseEvent EventRaised(sEventName, oEventData)
  Select Case sEventName
    Case "Open": RaiseEvent OnOpen(oEventData)
    Case "Close": RaiseEvent OnClose(oEventData)
    Case "Error": RaiseEvent OnError(oEventData)
    Case "Message": RaiseEvent OnMessage(oEventData)
  End Select
  On Error Resume Next
  Dim listener As Object
  For Each listener In Listeners
    listener.Run sEventName, oEventData
  Next
  On Error GoTo 0
End Sub

Private Sub Emit(ByVal name As String, Optional ByVal data As Variant, Optional ByVal code As Long, Optional ByVal reason As String)
  Dim eventData As Object: Set eventData = CreateEventData(name, data, code, reason)
  handleEvent name, eventData
End Sub
Private Function CreateEventData(ByVal name As String, ByVal data As Variant, ByVal code As Long, ByVal reason As String) As Object
  Dim eventData As Object: Set eventData = CreateObject("Scripting.Dictionary")
  eventData.Add "type", name
  eventData.Add "data", data
  eventData.Add "message", data
  eventData.Add "code", code
  eventData.Add "reason", reason
  Set CreateEventData = eventData
End Function
Private Sub CheckHandle(ByVal success As Boolean)
  If Not success Then Err.Raise vbObjectError + 139, "stdWebSocketNew", "Invalid UTF8/Unicode conversion"
End Sub
Private Sub NativeFailure(ByVal operation As String, ByVal code As Long)
  ReleaseHandles
  Emit "Error", operation & " failed: " & code
  Err.Raise vbObjectError + 139, "stdWebSocketNew", operation & " failed: " & code
End Sub
Private Sub ReleaseHandles()
  If This.hWebSocket <> 0 Then stdWebSocketDisconnect This.hWebSocket
  Set This.receiver = Nothing
  This.State = NotConnected
  This.sendingFragment = False
  This.opened = False
End Sub
Private Sub Class_Terminate()
  If This.isReceiver Then Exit Sub
  ReleaseHandles
  If This.library <> 0 Then FreeLibrary This.library
  This.library = 0
End Sub
Private Function ByteCount(ByRef bytes() As Byte) As Long
  On Error GoTo emptyArray
  ByteCount = UBound(bytes) - LBound(bytes) + 1
emptyArray:
End Function
Private Function Encode(ByVal value As String) As Byte()
  Dim bytes() As Byte, n As Long
  If Len(value) = 0 Then
    bytes = StrConv(vbNullString, vbFromUnicode)
  Else
    n = WideCharToMultiByte(65001, &H80, StrPtr(value), Len(value), 0, 0, 0, 0)
    CheckHandle n > 0
    ReDim bytes(0 To n - 1)
    CheckHandle WideCharToMultiByte(65001, &H80, StrPtr(value), Len(value), VarPtr(bytes(0)), n, 0, 0) = n
  End If
  Encode = bytes
End Function
Private Function Decode(ByRef bytes() As Byte, ByVal n As Long) As String
  If n = 0 Then Exit Function
  Dim chars As Long
  chars = MultiByteToWideChar(65001, 8, VarPtr(bytes(LBound(bytes))), n, 0, 0)
  CheckHandle chars > 0
  Decode = String$(chars, vbNullChar)
  CheckHandle MultiByteToWideChar(65001, 8, VarPtr(bytes(LBound(bytes))), n, StrPtr(Decode), chars) = chars
End Function
Private Sub ParseURL(ByVal value As String, ByRef host As String, ByRef port As Long, ByRef path As String, ByRef flags As Long)
  Dim scheme As String, rest As String, authority As String, at As Long, i As Long, portText As String
  at = InStr(value, "://")
  If at = 0 Then GoTo invalid
  scheme = LCase$(Left$(value, at - 1))
  Select Case scheme
    Case "ws": port = 80
    Case "wss": port = 443: flags = &H800000
    Case Else: GoTo invalid
  End Select
  rest = Mid$(value, at + 3)
  If InStr(rest, "#") Or InStr(rest, "\") Then GoTo invalid
  at = Len(rest) + 1
  For i = 1 To Len(rest)
    If Mid$(rest, i, 1) = "/" Or Mid$(rest, i, 1) = "?" Then
      at = i: Exit For
    End If
  Next
  authority = Left$(rest, at - 1): path = Mid$(rest, at)
  If InStr(authority, "@") Then GoTo invalid
  If path = "" Then path = "/"
  If Left$(path, 1) = "?" Then path = "/" & path
  If Left$(authority, 1) = "[" Then
    at = InStr(authority, "]")
    If at < 3 Then GoTo invalid
    host = Mid$(authority, 2, at - 2)
    If InStr(host, ":") = 0 Then GoTo invalid
    If Len(authority) > at Then
      If Mid$(authority, at + 1, 1) <> ":" Then GoTo invalid
      portText = Mid$(authority, at + 2)
      If portText = "" Then GoTo invalid
    End If
  Else
    at = InStr(authority, ":")
    If at Then
      host = Left$(authority, at - 1): portText = Mid$(authority, at + 1)
      If portText = "" Then GoTo invalid
    Else
      host = authority
    End If
  End If
  If host = "" Then GoTo invalid
  If InStr(host, "[") Or InStr(host, "]") Then GoTo invalid
  If portText <> "" Then
    If Len(portText) > 5 Or portText Like "*[!0-9]*" Then GoTo invalid
    port = CLng(portText)
    If port < 1 Or port > 65535 Then GoTo invalid
  End If
  For i = 1 To Len(value)
    If AscW(Mid$(value, i, 1)) <= 32 Or AscW(Mid$(value, i, 1)) > 126 Then GoTo invalid
  Next
  Exit Sub
invalid:
  Err.Raise 5, "stdWebSocketNew.Create", "Invalid WebSocket URL (use ASCII/Punycode and percent-encoded path)"
End Sub
