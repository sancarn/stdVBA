VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdHTTPServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'  Dim x As stdHTTPServer
'  Set x = stdHTTPServer.Create(8091)
'  Dim bshit As Boolean: bshit = True
'  While bshit
'    DoEvents
'    Debug.Print x.Status
'    Select Case x.Status
'      Case EHTTPServerStatus.AwaitingRequestForData
'        data = x.getRequestData()
'        Debug.Print StrConv(data, vbUnicode)
'      Case EHTTPServerStatus.AwaitingResponse
'        x.sendResponse "Hello world"
'        If StrConv(data, vbUnicode) Like "GET /stop*" Then bshit = False
'    End Select
'  Wend

#If VBA7 = 0 Then
    Enum LongPtr
        [_]
    End Enum
#End If

Public Enum EHTTPServerStatus
  [_Unknown] = -1
  WaitingForConnection = 0
  AwaitingRequestForData = 1
  AwaitingResponse = 2
End Enum

Private Enum addressFamily
    AF_UNSPEC = 0
    AF_INET = 2
    AF_IPX = 6
    AF_APPLETALK = 16
    AF_NETBIOS = 17
    AF_INET6 = 23
    AF_IRDA = 26
    AF_BTH = 32
End Enum
Private Enum TypeSpecification
    SOCK_STREAM = 1
    SOCK_DGRAM = 2
    SOCK_RAW = 3
    SOCK_RDM = 4
    SOCK_SEQPACKET = 5
End Enum
Private Enum protocol
    IPPROTO_IP = 0
    IPPROTO_ICMP = 1
    IPPROTO_IGMP = 2
    BTHPROTO_RFCOMM = 3
    IPPROTO_TCP = 6
    IPPROTO_UDP = 17
    IPPROTO_ICMPV6 = 58
    IPPROTO_RM = 113
End Enum

Private Type wsadata
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 255) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As LongPtr
End Type

Private Type WinsockAddress
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero(0 To 7) As Byte
End Type

Private Type timeval
    tv_sec As Long
    tv_usec As Long
End Type

Private Enum WSAErrorCodes
    WSANOTINITIALISED = 10093
    WSAEFAULT = 10014
    WSAENETDOWN = 10050
    WSAEINVAL = 10022
    WSAEINTR = 10004
    WSAEINPROGRESS = 10036
    WSAENOTSOCK = 10038
End Enum

Private Const FD_SETSIZE = 64
Private Type FD_SET
    fd_count As Integer
    fd_array(FD_SETSIZE) As LongPtr
End Type


#If VBA7 Then
    Private Declare PtrSafe Function socket Lib "wsock32.dll" (ByVal addressFamily As Long, ByVal typeSpec As Long, ByVal protocol As Long) As LongPtr
    Private Declare PtrSafe Function bind Lib "wsock32.dll" (ByVal socketHandle As LongPtr, ByRef socketAddress As WinsockAddress, ByVal addressLength As Long) As Long
    Private Declare PtrSafe Function listen Lib "wsock32.dll" (ByVal socketHandle As LongPtr, ByVal backlog As Long) As Long
    Private Declare PtrSafe Function accept Lib "wsock32.dll" (ByVal socketHandle As LongPtr, ByRef socketAddress As WinsockAddress, ByRef addressLength As Long) As LongPtr
    Private Declare PtrSafe Function recv Lib "wsock32.dll" (ByVal socketHandle As LongPtr, ByVal buffer As LongPtr, ByVal bufferLength As Long, ByVal flags As Long) As Long
    Private Declare PtrSafe Function send Lib "wsock32.dll" (ByVal socketHandle As LongPtr, ByVal buffer As LongPtr, ByVal bufferLength As Long, ByVal flags As Long) As Long
    Private Declare PtrSafe Function closesocket Lib "wsock32.dll" (ByVal socketHandle As LongPtr) As Long
    Private Declare PtrSafe Function WSAGetLastError Lib "wsock32.dll" () As Long
    Private Declare PtrSafe Function WSACleanup Lib "wsock32.dll" () As Long
    Private Declare PtrSafe Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequested As Long, ByRef wsadata As wsadata) As Long
    Private Declare PtrSafe Function htons Lib "wsock32.dll" (ByVal hostshort As Long) As Integer
    Private Declare PtrSafe Function WSASelect Lib "wsock32.dll" Alias "select" (ByVal nfds As Long, ByRef readFDSet As FD_SET, ByRef writeFDSet As FD_SET, ByRef errorFDSet As FD_SET, ByRef timeout As timeval) As Long
#Else
    Private Declare Function socket lib "wsock32.dll" (ByVal addressFamily As Long, ByVal typeSpec As Long, ByVal protocol As Long) As LongPtr
    Private Declare Function bind lib "wsock32.dll" (ByVal socketHandle As LongPtr, ByRef socketAddress As WinsockAddress, ByVal addressLength As Long) As Long
    Private Declare Function listen lib "wsock32.dll" (ByVal socketHandle As LongPtr, ByVal backlog As Long) As Long
    Private Declare Function accept lib "wsock32.dll" (ByVal socketHandle As LongPtr, ByRef socketAddress As WinsockAddress, ByRef addressLength As Long) As LongPtr
    Private Declare Function recv lib "wsock32.dll" (ByVal socketHandle As LongPtr, ByVal buffer As LongPtr, ByVal bufferLength As Long, ByVal flags As Long) As Long
    Private Declare Function send lib "wsock32.dll" (ByVal socketHandle As LongPtr, ByVal buffer As LongPtr, ByVal bufferLength As Long, ByVal flags As Long) As Long
    Private Declare Function closesocket lib "wsock32.dll" (ByVal socketHandle As LongPtr) As Long
    Private Declare Function WSAGetLastError lib "wsock32.dll" () As Long
    Private Declare Function WSACleanup lib "wsock32.dll" () As Long
    Private Declare Function WSAStartup lib "wsock32.dll" (ByVal wVersionRequested As Long, ByRef wsadata As WSADATA) As Long
    Private Declare Function htons lib "wsock32.dll" (ByVal hostshort As Long) As Integer
    Private Declare Function WSASelect Lib "wsock32.dll" Alias "select" (ByVal nfds As Long, ByRef socketToRead As LongPtr, ByRef socketToWrite as LongPtr, ByRef socketForError As LongPtr, ByRef timeout As timeval) As Long
#End If

#If Win64 Then
    Const NullPtr As LongLong = 0^
#Else
    Const NullPtr As Long = 0&
#End If



Private Type TThis
    winsockData As wsadata
    sockaddr As WinsockAddress
    listenHandle As LongPtr
    readFDSet As FD_SET
    
    acceptedClient As LongPtr
    OverrideStatus As EHTTPServerStatus
End Type
Private This As TThis

'Create a HTTP server
'@paran port - The port to listen on
'@param maxQueue - The maximum number of connections to allow to queue.
Public Function Create(ByVal port As Long, Optional ByVal maxQueue As Long = 10) As stdHTTPServer
    Set Create = New stdHTTPServer
    Call Create.protInit(port, maxQueue)
End Function

Public Sub protInit(ByVal port As Long, ByVal maxQueue As Long)
    ' Initialize Winsock
    Dim success As Long: success = WSAStartup(&H202, This.winsockData)
    If success <> 0 Then Err.Raise vbObjectError + 1, "stdHTTPServer.protInit", "WSAStartup failed with error " & success
    
    ' Create a socket for listening
    This.listenHandle = socket(addressFamily.AF_INET, TypeSpecification.SOCK_STREAM, protocol.IPPROTO_TCP)
    If This.listenHandle = 0 Then
        Err.Raise vbObjectError + 2, "stdHTTPServer.protInit", "socket failed with error " & WSAGetLastError
        WSACleanup
    End If

    ' Prepare the sockaddr structure
    With This.sockaddr
        .sin_family = addressFamily.AF_INET
        .sin_port = htons(port)
    End With

    ' Bind the socket
    success = bind(This.listenHandle, This.sockaddr, LenB(This.sockaddr))
    If success = -1 Then
        Err.Raise vbObjectError + 3, "stdHTTPServer.protInit", "bind failed with error " & WSAGetLastError
        WSACleanup
    End If

    ' Listen on the socket
    success = listen(This.listenHandle, maxQueue)
    If success = -1 Then
        Err.Raise vbObjectError + 4, "stdHTTPServer.protInit", "listen failed with error " & WSAGetLastError
        WSACleanup
    End If

    This.readFDSet.fd_array(0) = This.listenHandle
    This.readFDSet.fd_count = 1
    This.OverrideStatus = EHTTPServerStatus.[_Unknown]
End Sub

Function Status() As EHTTPServerStatus
    If This.OverrideStatus = EHTTPServerStatus.[_Unknown] Then
      Static emptyFDSet As FD_SET
      Dim timeout As timeval
      Dim success As Long
      This.readFDSet.fd_count = 1
      Status = WSASelect(0, This.readFDSet, emptyFDSet, emptyFDSet, timeout)
    Else
      Status = This.OverrideStatus
    End If
End Function

Function getRequestData(Optional ByRef bytesReceived As Long = 0) As Byte()
  This.acceptedClient = accept(This.listenHandle, This.sockaddr, LenB(This.sockaddr))
  If This.acceptedClient = 0 Then Err.Raise vbObjectError + 5, "stdHTTPServer.protInit", "accept failed with error " & WSAGetLastError
  
  Dim buffer() As Byte: ReDim buffer(0 To 1024)
  bytesReceived = recv(This.acceptedClient, VarPtr(buffer(0)), 1024, 0)
  getRequestData = buffer
  This.OverrideStatus = EHTTPServerStatus.AwaitingResponse
End Function

Sub sendResponse(ByVal data As String, Optional ByVal ContentType As String = "text/plain")
  Dim sendData As String
  sendData = "HTTP/1.1 200 OK" & vbCrLf & _
             "Content-Type: " & ContentType & vbCrLf & _
             vbCrLf & _
             data
  Dim buffer() As Byte: buffer = StrConv(sendData, vbFromUnicode)
  x = send(This.acceptedClient, VarPtr(buffer(0)), UBound(buffer) + 1, 0)
  Call closesocket(This.acceptedClient)
  This.OverrideStatus = EHTTPServerStatus.[_Unknown]
End Sub

Private Sub Class_Terminate()
    Call closesocket(This.listenHandle)
    Call WSACleanup
End Sub
