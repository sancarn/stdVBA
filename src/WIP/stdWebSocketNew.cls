VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdWebSocketNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'Spec: 
'
'Functions implemented on this class:
'=============================================
'  CONSTRUCTORS
'    [ ] Create
'
'  INSTANCE METHODS
'  Many methods were inspired by those in Ruby's Enumerable: https://ruby-doc.org/core-2.7.2/Enumerable.html
'    [ ] send(data)
'    [ ] close(code,reason)
'    [ ] disconnect()
'    [ ] Get url()
'  PROTECTED INSTANCE METHODS
'    [ ] handleEvent()
'  EVENTS
'    [ ] OnOpen(data)
'    [ ] OnClose(data)
'    [ ] OnError(data)
'    [ ] OnMessage(data)
'    [ ] EventRaised(name,data)
'TODO: IDEALLY WE'D DO EVERYTHING ASYNCHRONOUSLY WITH EVENTS, HOWEVER THIS ISN'T DOABLE UNTIL WE GET A OBJECT CALLER THUNK. EVENTS WE'D SUPPORT:
'  OnOpen(ByVal eventData As Variant)
'  OnClose(ByVal eventData As Variant)
'  OnError(ByVal eventData As Variant)
'  OnMessage(ByVal eventData As Variant)

' Acknowledgements
' * This class, written by Sancarn, was an entire re-write of https://github.com/EagleAglow/vba-websocket-class





'See parameters of https://docs.microsoft.com/en-us/windows/win32/api/winhttp/nf-winhttp-winhttpopen
Private Enum EHttpProxyType
  DEFAULT_PROXY = 0
  NO_PROXY = 1
  NAMED_PROXY = 3
  AUTOMATIC_PROXY = 4
End enum

'WinHTTP WebSocket Buffer types
Private Enum EBufferType
  BinaryMessage = 0
  BinaryFragment = 1
  Utf8Message = 2
  Utf8Fragment = 3
  CloseMessage = 4
End Enum


#if VBA7 then
  Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long

  Private Declare PtrSafe Function WinHttpOpen Lib "winhttp" (ByVal pszAgentW As LongPtr, ByVal dwAccessType As Long, ByVal pszProxyW As LongPtr, ByVal pszProxyBypassW As LongPtr, ByVal dwFlags As Long ) As LongPtr
  Private Declare PtrSafe Function WinHttpConnect Lib "winhttp" (ByVal hSession As LongPtr, ByVal pswzServerName As LongPtr, ByVal nServerPort As Long, ByVal dwReserved As Long ) As LongPtr
  Private Declare PtrSafe Function WinHttpOpenRequest Lib "winhttp" ( ByVal hConnect As LongPtr, ByVal pwszVerb As LongPtr, ByVal pwszObjectName As LongPtr, ByVal pwszVersion As LongPtr, ByVal pwszReferrer As LongPtr, ByVal ppwszAcceptTypes As LongPtr, ByVal dwFlags As Long ) As LongPtr
  Private Declare PtrSafe Function WinHttpSetOption Lib "winhttp" ( ByVal hInternet As LongPtr, ByVal dwOption As Long, ByVal lpBuffer As LongPtr, ByVal dwBufferLength As Long ) As Long
  Private Declare PtrSafe Function WinHttpSendRequest Lib "winhttp" ( ByVal hRequest As LongPtr, ByVal lpszHeaders As LongPtr, ByVal dwHeadersLength As Long, ByVal lpOptional As LongPtr, ByVal dwOptionalLength As Long, ByVal dwTotalLength As Long, ByVal dwContext As Long ) As Long
  Private Declare PtrSafe Function WinHttpReceiveResponse Lib "winhttp" ( ByVal hRequest As LongPtr, ByVal lpReserved As LongPtr ) As Long
  Private Declare PtrSafe Function WinHttpWebSocketCompleteUpgrade Lib "winhttp" ( ByVal hRequest As LongPtr, ByVal pContext As LongPtr ) As LongPtr
  Private Declare PtrSafe Function WinHttpCloseHandle Lib "winhttp" ( ByVal hRequest As LongPtr ) As Long
  Private Declare PtrSafe Function WinHttpWebSocketSend Lib "winhttp" ( ByVal hWebSocket As LongPtr, ByVal eEBufferType As Long, ByVal pvBuffer As LongPtr, ByVal dwBufferLength As Long ) As Long
  Private Declare PtrSafe Function WinHttpWebSocketReceive Lib "winhttp" ( ByVal hWebSocket As LongPtr, ByRef pvBuffer As Any, ByVal dwBufferLength As Long, ByRef pdwBytesRead As Long, ByRef peEBufferType As Long ) As Long
  Private Declare PtrSafe Function WinHttpWebSocketClose Lib "winhttp" ( ByVal hWebSocket As LongPtr, ByVal usStatus As Integer, ByVal pvReason As LongPtr, ByVal dwReasonLength As Long ) As Long
  Private Declare PtrSafe Function WinHttpWebSocketQueryCloseStatus Lib "winhttp" ( ByVal hWebSocket As LongPtr, ByRef usStatus As Integer, ByRef pvReason As Any, ByVal dwReasonLength As Long, ByRef pdwReasonLengthConsumed As LongPtr ) As Long
  Private Declare PtrSafe Function WinHttpQueryHeaders Lib "winhttp" ( ByVal hRequest As LongPtr, ByVal dwInfoLevel As Long, ByVal pwszName As LongPtr, ByRef lpBuffer As Long, ByRef lpdwBufferLength As Long, ByRef lpdwIndex As Long ) As Long
  
  'Not currently used. Cannot really use callbacks yet as would need a BAS module anyway. Not sure what a web socket action is.
  'For status callback we'd also need `Private Const WINHTTP_FLAG_ASYNC = &H10000000`
  'Private Declare PtrSafe Function WinHttpSetStatusCallback Lib "winhttp" ( ByVal hWebSocket As LongPtr, ByVal lpfnInternetCallback As LongPtr, ByVal dwNotificationFlags As Long, ByVal dwReserved As LongPtr ) As Long
  'Private Declare PtrSafe Function WebSocketGetAction Lib "websocket" ( ByVal hWebSocket As LongPtr, ByVal eActionQueue As Long, ByVal pDataBuffers As LongPtr, ByRef pulDataBufferCount As Long, ByRef pAction As Long, ByRef pEBufferType As Long, ByRef pvApplicationContext As LongPtr, ByRef pvActionContext As LongPtr ) As Long

  'UTF8
  Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cchMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long ) As Long
  Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long ) As Long

  'Properties
  Private hConnection as LongPtr
  Private hSession as LongPtr
  Private hRequest as LongPtr
  Private hWebSocket as LongPtr
#else
  Private Declare Function GetLastError Lib "kernel32" () As Long

  Private Declare Function WinHttpOpen Lib "winhttp" (ByVal pszAgentW As Long, ByVal dwAccessType As Long, ByVal pszProxyW As Long, ByVal pszProxyBypassW As Long, ByVal dwFlags As Long ) As Long
  Private Declare Function WinHttpConnect Lib "winhttp" (ByVal hSession As Long, ByVal pswzServerName As Long, ByVal nServerPort As Long, ByVal dwReserved As Long ) As Long
  Private Declare Function WinHttpOpenRequest Lib "winhttp" ( ByVal hConnect As Long, ByVal pwszVerb As Long, ByVal pwszObjectName As Long, ByVal pwszVersion As Long, ByVal pwszReferrer As Long, ByVal ppwszAcceptTypes As Long, ByVal dwFlags As Long ) As Long
  Private Declare Function WinHttpSetOption Lib "winhttp" ( ByVal hInternet As Long, ByVal dwOption As Long, ByVal lpBuffer As Long, ByVal dwBufferLength As Long ) As Long
  Private Declare Function WinHttpSendRequest Lib "winhttp" ( ByVal hRequest As Long, ByVal lpszHeaders As Long, ByVal dwHeadersLength As Long, ByVal lpOptional As Long, ByVal dwOptionalLength As Long, ByVal dwTotalLength As Long, ByVal dwContext As Long ) As Long
  Private Declare Function WinHttpReceiveResponse Lib "winhttp" ( ByVal hRequest As Long, ByVal lpReserved As Long ) As Long
  Private Declare Function WinHttpWebSocketCompleteUpgrade Lib "winhttp" ( ByVal hRequest As Long, ByVal pContext As Long ) As Long
  Private Declare Function WinHttpCloseHandle Lib "winhttp" ( ByVal hRequest As Long ) As Long
  Private Declare Function WinHttpWebSocketSend Lib "winhttp" ( ByVal hWebSocket As Long, ByVal eEBufferType As Long, ByVal pvBuffer As Long, ByVal dwBufferLength As Long ) As Long
  Private Declare Function WinHttpWebSocketReceive Lib "winhttp" ( ByVal hWebSocket As Long, ByRef pvBuffer As Any, ByVal dwBufferLength As Long, ByRef pdwBytesRead As Long, ByRef peEBufferType As Long ) As Long
  Private Declare Function WinHttpWebSocketClose Lib "winhttp" ( ByVal hWebSocket As Long, ByVal usStatus As Integer, ByVal pvReason As Long, ByVal dwReasonLength As Long ) As Long
  Private Declare Function WinHttpWebSocketQueryCloseStatus Lib "winhttp" ( ByVal hWebSocket As Long, ByRef usStatus As Integer, ByRef pvReason As Any, ByVal dwReasonLength As Long, ByRef pdwReasonLengthConsumed As Long ) As Long
  Private Declare Function WinHttpQueryHeaders Lib "winhttp" ( ByVal hRequest As Long, ByVal dwInfoLevel As Long, ByVal pwszName As Long, ByRef lpBuffer As Long, ByRef lpdwBufferLength As Long, ByRef lpdwIndex As Long ) As Long
  
  'Not currently used
  'Private Declare Function WinHttpSetStatusCallback Lib "winhttp" ( ByVal hWebSocket As Long, ByVal lpfnInternetCallback As Long, ByVal dwNotificationFlags As Long, ByVal dwReserved As Long ) As Long
  'Private Declare Function WebSocketGetAction Lib "websocket" ( ByVal hWebSocket As Long, ByVal eActionQueue As Long, ByVal pDataBuffers As Long, ByRef pulDataBufferCount As Long, ByRef pAction As Long, ByRef pEBufferType As Long, ByRef pvApplicationContext As Long, ByRef pvActionContext As Long ) As Long

  'UTF8
  Private Declare Function MultiByteToWideChar Lib "kernel32" ( ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
  Private Declare Function WideCharToMultiByte Lib "kernel32" ( ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long ) As Long

  'Properties
  Private hConnection as Long
  Private hSession as Long
  Private hRequest as Long
  Private hWebSocket as Long
#end if

Public Enum EMessageType
  Binary
  Utf8
End Enum

'WebSocket State
Public Enum EWebSocketState
  NotConnected
  Connecting
  Connected
  Disconnecting
End Enum

Private pState as EWebSocketState
Private pURL as string
Private pAgent as string
Private pPort as long


'Creates and returns an active WebSocket object
'@constructor
'@param {String} url - The websocket server url to connect to
'@returns {stdWebSocket} - The created websocket
'@example stdWebSocket.Create("wss://ws.vi-server.org/mirror",1939)
Public Function Create(ByVal url As String, ByVal port as string, optional byval agent as string = "stdVBA::stdWebSocket") As stdWebSocketNew
  Set Create = New stdWebSocketNew
  Call Create.protInit(url,port,agent)
End Function
Public Sub protInit(ByVal url As String,ByVal port as string, ByVal agent as string)
  pUrl = url
  pPort = port
  pAgent = agent
  Dim sError as string: sError =  Connect()
  if sError <> "" then
    Err.raise 1, "stdWebsocket::protInit", sError
  end if
End Sub

'Gets the url of the websocket
'@returns {string}
Public Property Get url() As String
  url = pURL
End Property

'Gets the port that the websocket is connected to
'@returns {long}
Public Property Get port() As long
  port = pPort
End Property

'Obtain the state of the websocket object
'@returns {EWebSocketState} The state of the websocket
Public Property Get State() as EWebSocketState
  State = pState
End Property

'Sends data through the WebSocket
'@param {String | Bytes} data - The data to send
'@param {EMessageType=Utf8} Type - Type of data sent. Either Binary or UTF8
'@param {Boolean=False} PartialSend - set to true if you only wnat to send a partial message (i.e. a fragment), server will wait until this param is false again
Public Sub Send(ByVal data as variant, optional ByVal iType as EMessageType = EMessageType.Utf8, Optional ByVal bPartialSend as boolean = false)
  if pState = EWebSocketState.Connected then
    Dim bin() as byte, iBufferType as EBufferType
    select case iType
      case EMessageType.Binary
        iBufferType = iif(bPartialSend, EBufferType.BinaryFragment, EBufferType.BinaryMessage)
        bin = data
      case EMessageType.Utf8
        iBufferType = iif(bPartialSend, EBufferType.Utf8Fragment, EBufferType.Utf8Message)
        Dim sData as string: sData = data
        bin = Utf8BytesFromString(sData)
      case else
        Err.Raise 1, "stdWebSocket::Send", "Unknown message type."
    end select

    'Raise error when iError not successful (0).
    Dim iLength as long: iLength = UBound(bin) - Lbound(bin) + 1
    Dim iError as long:  iError = WinHttpWebSocketSend(hWebSocket, iBufferType, VarPtr(bin(LBound(bin))), iLength)
    if iError <> 0 then Err.Raise 1, "stdWebSocket::Send" , "Send Error " & iError & " (" & GetLastError() & ")"
  else
    Err.Raise 1, "stdWebSocket::Send", "Must be connected to send data"
  end if
End Sub

'Receive data from a websocket
'@returns {String | Byte()} data - The data sent by the server
'@remark Will hang until server sends a response. Unfortunately can't peak for a response
'@remark A list of WebSocket close statuses can be found at https://docs.microsoft.com/en-us/windows/win32/api/winhttp/ne-winhttp-winhttp_web_socket_close_status 
Public Function Receive(Optional ByRef iType as EMessageType) as Variant
  if pState = EWebSocketState.Connected then
    Dim buff() as Byte: Redim buff(1023)
    Dim iBuffLen as long: iBuffLen = 1024
    Dim dwBytesTransferred as long
    Dim iEBufferType as long
    Dim iStatus as long
    iStatus = WinHttpWebSocketReceive(hWebSocket, buff(0), 1024, dwBytesTransferred, iEBufferType)
    if iStatus = 0 then
      Dim iPos as long
      select case iEBufferType 
        case EBufferType.BinaryMessage
          'Buffer contains entire binary message (or last part of the message)
          iType = EMessageType.Binary
          Redim Preserve buff(cInt(dwBytesTransferred))
          Receive = buff
        case EBufferType.BinaryFragment
          'Buffer contains part of a binary message, rest of binary message is to follow, continue calling `WinHttpWebSocketReceive` until receiving `EBufferType.BinaryMessage`
          iPos = 0
          While iEBufferType = EBufferType.BinaryFragment
            iPos = iPos + dwBytesTransferred
            redim preserve buff(iPos + 1024)
            iStatus = WinHttpWebSocketReceive(hWebSocket, buff(iPos), 1024, dwBytesTransferred, iEBufferType)
            Debug.Assert False 'TODO: Remove
          Wend
          iType = EMessageType.Binary
          redim preserve buff(iPos + dwBytesTransferred)
          Receive = buff
        case EBufferType.Utf8Message
          'Buffer contains entire Utf8 Message
          iType = EMessageType.Utf8
          Redim Preserve buff(cInt(dwBytesTransferred))
          Receive = Utf8BytesToString(buff)
        case EBufferType.Utf8Fragment
          'Buffer contains part of a Utf8 message, rest of Utf8 message is to follow, continue calling `WinHttpWebSocketReceive` until receiving `EBufferType.Utf8Message`
          iPos = 0
          While iEBufferType = EBufferType.Utf8Fragment
            iPos = iPos + dwBytesTransferred
            redim preserve buff(iPos + 1024)
            iStatus = WinHttpWebSocketReceive(hWebSocket, buff(iPos), 1024, dwBytesTransferred, iEBufferType)
            Debug.Assert False 'TODO: Remove
          Wend
          iType = EMessageType.Utf8
          redim preserve buff(iPos + dwBytesTransferred)
          Receive = Utf8BytesToString(buff)
        case EBufferType.CloseMessage
          'Socket closure, call disconnect
          Dim sError as string: sError = Disconnect(true)
          Err.Raise 1, "stdWebsocket::Receive", "Server requested socket closure."
      end select
    else
      Err.Raise 1, "stdWebsocket::Receive", "Receive errored with error " & iStatus
    end if
  else
    Err.Raise 1, "stdWebsocket::Receive", "Must be connected to send data"
  end if
End Function


'Closes the WebSocket connection
'TODO: '@param {Long = 1000} Code - A numeric value indicating the status code explaining why the connection is being closed. If this parameter is not specified, a default value of 1005 is assumed. See the [list of status](https://developer.mozilla.org/en-US/docs/Web/API/CloseEvent#status_codes) code of CloseEvent for permitted values.
'TODO: '@param {String = ""} Reason - A human-readable string explaining why the connection is closing.
Public Sub CloseSocket()
  Dim sError as string: sError = Disconnect()
  if sError <> "" then
    Err.Raise 1, "stdWebsocket::CloseSocket", sError
  end if
End Sub

'Connect to the server/port and begin websocket connection
'@returns {string} Any errors that occurred. "" if no errors
Private Function Connect() as string
  Dim sErrorText as string: sErrorText = ""
  If pState <> EWebSocketState.NotConnected then GoTo Fail_ActiveSocket
  if pURL = "" then GoTo Fail_Parameter

  'TODO: For a given URL e.g. "ws://vi-server.org:1939/mirror" we need to convert this into:
  '  sUrl  = "vi-server.org"  '<-- Note this URL must be in "PunyCode". Can make a call to `IdnToAscii` of `Normaliz.dll` for internationalisation. 
  '  iPort = 1939
  '  sPath = "mirro"
  Dim sUrl as string:  sUrl = "vi-server.org"
  Dim iPort as long:   iPort = 1939
  Dim sPath as string: sPath = ""


  'Flag that we are connecting
  pState = EWebSocketState.Connecting
  
  'Create a session with agent string
  Const WINHTTP_FLAG_SYNC = &H0
  hSession = WinHttpOpen(StrPtr(pAgent), EHttpProxyType.DEFAULT_PROXY, 0, 0, WINHTTP_FLAG_SYNC)
  If hSession = 0 Then GoTo Fail_Session

  'Connect to server
  hConnection = WinHttpConnect(hSession, StrPtr(sURL), iPort, 0)
  if hConnection = 0 then GoTo Fail_Connection

  'Open GET request
  Dim sMethod as string: sMethod = "GET" ' ALWAYS FOR WEBSOCKET
  hRequest = WinHttpOpenRequest(hConnection, StrPtr(sMethod), StrPtr(sPath), 0, 0, 0, 0)
  if hRequest = 0 then GoTo Fail_Request

  'Upgrade request to websocket
  Const WINHTTP_OPTION_UPGRADE_TO_WEB_SOCKET = 114
  Dim iStatus as long
  iStatus = WinHttpSetOption(hRequest, WINHTTP_OPTION_UPGRADE_TO_WEB_SOCKET, 0, 0)
  if iStatus = 0 then GoTo Fail_SetOptionUpgrade

  'Websocket handshake - send request
  Dim sHeader as string: sHeader = ""
  sHeader = sHeader & "Host: " & pURL & vbCrLf                          ' may be redundant or unnecessary
  sHeader = sHeader & "Sec-WebSocket-Version: 13" & vbCrLf              ' may be redundant or unnecessary
  sHeader = sHeader & "Sec-Websocket-Protocol: echo-protocol"           ' subprotocol
  Const WINHTTP_NO_REQUEST_DATA = 0
  iStatus = WinHttpSendRequest(hRequest, StrPtr(sHeader), -1, WINHTTP_NO_REQUEST_DATA, 0, 0, 0)
  if iStatus = 0 then GoTo Fail_SendHandshakeRequest
  
  'Websocket handshake - receive response
  iStatus = WinHttpReceiveResponse(hRequest, 0)
  if iStatus = 0 then GoTo Fail_RecvHandshakeResponse

  ' See if the HTTP Response confirms the upgrade, with HTTP status code 101.
  Const WINHTTP_QUERY_STATUS_CODE = 19  ' special: part of status line
  Const WINHTTP_QUERY_FLAG_NUMBER = &H20000000   ' bit flag to get result as number
  Const WINHTTP_HEADER_NAME_BY_INDEX = 0
  Const WINHTTP_NO_HEADER_INDEX = 0
  Dim dwStatusCode as long, sizeStatusCode As Long: sizeStatusCode = 4 'for HTTP result request (4 bytes long)
  iStatus = WinHttpQueryHeaders(hRequest, (WINHTTP_QUERY_STATUS_CODE Or WINHTTP_QUERY_FLAG_NUMBER), WINHTTP_HEADER_NAME_BY_INDEX, dwStatusCode, sizeStatusCode, WINHTTP_NO_HEADER_INDEX)
  if iStatus = 0 then GoTo Fail_QueryHeadersForStatusCode
  if dwStatusCode <> 101 then GoTo Fail_ServerRefusedProtocolSwitch

  ' Obtain Websocket handle
  hWebSocket = WinHttpWebSocketCompleteUpgrade(hRequest, 0)
  if hWebSocket = 0 then GoTo Fail_ObtainWebsocketPointer
  Call WinHttpCloseHandle(hRequest)    'TODO: Check this doesn't have any negative affects
  Call WinHttpCloseHandle(hConnection) 'TODO: Check this doesn't have any negative affects
  Call WinHttpCloseHandle(hSession)    'TODO: Check this doesn't have any negative affects

  pState = EWebSocketState.Connected
  Connect = ""
  Exit Function


Fail_ObtainWebsocketPointer:
  if sErrorText = "" then sErrorText = "Failed to complete websocket upgrade and obtain the websocket pointer"  
Fail_ServerRefusedProtocolSwitch:
  if sErrorText = "" then sErrorText = "Server refused to switch to websocket protocol"
Fail_QueryHeadersForStatusCode:
  if sErrorText = "" then sErrorText = "Failed to query response headers for status code"
Fail_RecvHandshakeResponse:
  if sErrorText = "" then sErrorText = "Failed to receive handshake response"
Fail_SendHandshakeRequest:
  if sErrorText = "" then sErrorText = "Failed in sending Websocket handshake request"
Fail_SetOptionUpgrade:
  if sErrorText = "" then sErrorText = "Failed to set Winsocket Upgrade option"
  Call WinHttpCloseHandle(hRequest)
Fail_Request:
  if sErrorText = "" then sErrorText = "Failed to create HTTP GET Request"
  Call WinHttpCloseHandle(hConnection)
Fail_Connection:
  if sErrorText = "" then sErrorText = "Failed to reach """ & pUrl & ":" & pPort & """"
  Call WinHttpCloseHandle(hSession)
Fail_Session:
  if sErrorText = "" then sErrorText = "Failed to create HTTP session with agent """ & pAgent & """"
  pState = EWebSocketState.NotConnected
  Connect = sErrorText
  Exit Function
Fail_Parameter:
  Connect = "Missing server url"
  Exit Function
Fail_ActiveSocket:
  Connect = "Must be disconnected before attempting to connect"
  Exit Function
End Function

'Closes and deletes the WebSocket, removing references so the object can be garbage collected
'@remark A list of WebSocket close statuses can be found at https://docs.microsoft.com/en-us/windows/win32/api/winhttp/ne-winhttp-winhttp_web_socket_close_status 
Private Function Disconnect(optional byval bForce as boolean = false) as string
  Const SUCCESS_CLOSE_STATUS = 1000
  Dim sErrorText as string: sErrorText = ""
  if pState <> EWebSocketState.Connected then GoTo Fail_SocketNotActive
  pState = EWebSocketState.Disconnecting

  'Close websocket
  Dim iStatus as Long
  iStatus = WinHttpWebSocketClose(hWebsocket, SUCCESS_CLOSE_STATUS, 0, 0)
  if iStatus <> 0 then GoTo Fail_WebSocketCloseFail

  'Check close status returned by the server
  Dim rgbCloseReasonBuffer(1 to 123) As Byte
  Dim dwCloseReasonLength As Long: dwCloseReasonLength = 0
  Dim usStatus As Integer:  usStatus = 0
  iStatus = WinHttpWebSocketQueryCloseStatus(hWebSocketHandle, usStatus, rgbCloseReasonBuffer(1), 123, CLngPtr(dwCloseReasonLength))
  if iStatus = 0 then
    If usStatus = SUCCESS_CLOSE_STATUS then
      pCloseReason = "Closed by host."
    else
      Dim sChars as string: sChars = rgbCloseReasonBuffer
      pCloseReason = "The server closed the connection with status code: " & usStatus & vbCrLf & "The reason for closure was: " & Left(sChars, dwCloseReasonLength) 
    end if
  else
    pCloseReason = "UNKNOWN - Could not query close reason."
  end if

ForceClose:
  'Cleanup
  if hWebsocket <> 0 then Call WinHttpCloseHandle(hWebSocket)
  if hRequest <> 0 then Call WinHttpCloseHandle(hRequest)
  if hConnection <> 0 then Call WinHttpCloseHandle(hConnection)
  if hSession <> 0 then Call WinHttpCloseHandle(hSession)

  'Return true
  pState = EWebSocketState.NotConnected
  Disconnect = ""
  Exit Function

Fail_WebSocketCloseFail:
  pState = EWebSocketState.Connected
  Disconnect = "Could not close websocket now"
  if bForce then Goto ForceClose:
  Exit Function
Fail_SocketNotActive:
  Connect = "Must be connected in order to disconnect"
  Exit Function
End Function

'Get a string from an array of utf8 bytes
'@params {Byte()} UTF8 Bytes to cast to a string
'@returns {string} UTF8 bytes in string representation
Private Function Utf8BytesToString(abUtf8Array() As Byte) as string
  ' CodePage constant for UTF-8
  Const CP_UTF8 = 65001

  ' Catch uninitialized input array
  if UBound(abUtf8Array) - LBound(abUtf8Array) + 1 > 0 then
    ' Get number of characters in output string
    Dim nChars As Long: nChars = MultiByteToWideChar(CP_UTF8, 0&, VarPtr(abUtf8Array(0)), -1, 0&, 0&)
    ' Dimension output buffer to receive string
    Dim strOut As String: strOut = String(nChars-1, 0)
    nChars = MultiByteToWideChar(CP_UTF8, 0&, VarPtr(abUtf8Array(0)), -1, StrPtr(strOut), nChars)
    Utf8BytesToString = Left$(strOut, nChars)
  else
    Utf8BytesToString = ""
  end if
End Function

'Get a string from an array of utf8 bytes
'@params {Byte()} string to cast to a UTF8 bytes
'@returns {string} string as an array of UTF8 bytes
Private Function Utf8BytesFromString(strInput As String) As Byte()
    ' CodePage constant for UTF-8
    Const CP_UTF8 = 65001

    Dim nBytes As Long
    Dim buff() As Byte
    
    'Catch empty string
    Utf8BytesFromString = vbNullString
    If Len(strInput) < 1 Then Exit Function

    ' Get length in bytes *including* terminating null
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, StrPtr(strInput), -1, 0&, 0&, 0&, 0&)

    ' We don't want the terminating null in our byte array, so ask for `nBytes-1` bytes
    ReDim buff(nBytes - 2)  ' NB ReDim with one less byte than you need
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, StrPtr(strInput), -1, VarPtr(buff(0)), nBytes - 1, 0&, 0&)
    Utf8BytesFromString = buff
End Function


