VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdHTTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Enum EHTTPSynchronisity
  HTTPSync
  HTTPAsync
End Enum
Public Enum EHTTPContentType
  ContentType_HTML = 1
  ContentType_Plain = 2
  ContentType_JSON = 4
  ContentType_XML = 8
  ContentType_CSV = 16
  ContentType_ZIP = 32
  ContentType_Binary = 64
End Enum
Public Enum EHTTPProxyType
  ProxyTypeAuto
  ProxyTypeManual
  ProxyTypeNone
End Enum


#If Win64 Then
  Private Const PTR_SIZE = 8
#Else
  Private Const PTR_SIZE = 4
#End If

Private Enum SetCredentialsType
  ForServer = 0
  ForProxy = 1
End Enum

Private Enum ProxySetting
  default = 0
  PreConfig = 0
  Direct = 1
  Proxy = 2
End Enum

'WINHTTP Errors
'@devRemark source https://raw.githubusercontent.com/libgit2/libgit2/main/deps/winhttp/winhttp.h
Private Enum WINHTTP_ERROR
  WINHTTP_ERROR_BASE = 12000
  ERROR_WINHTTP_OUT_OF_HANDLES = (WINHTTP_ERROR_BASE + 1)
  ERROR_WINHTTP_TIMEOUT = (WINHTTP_ERROR_BASE + 2)
  ERROR_WINHTTP_INTERNAL_ERROR = (WINHTTP_ERROR_BASE + 4)
  ERROR_WINHTTP_INVALID_URL = (WINHTTP_ERROR_BASE + 5)
  ERROR_WINHTTP_UNRECOGNIZED_SCHEME = (WINHTTP_ERROR_BASE + 6)
  ERROR_WINHTTP_NAME_NOT_RESOLVED = (WINHTTP_ERROR_BASE + 7)
  ERROR_WINHTTP_INVALID_OPTION = (WINHTTP_ERROR_BASE + 9)
  ERROR_WINHTTP_OPTION_NOT_SETTABLE = (WINHTTP_ERROR_BASE + 11)
  ERROR_WINHTTP_SHUTDOWN = (WINHTTP_ERROR_BASE + 12)
  ERROR_WINHTTP_LOGIN_FAILURE = (WINHTTP_ERROR_BASE + 15)
  ERROR_WINHTTP_OPERATION_CANCELLED = (WINHTTP_ERROR_BASE + 17)
  ERROR_WINHTTP_INCORRECT_HANDLE_TYPE = (WINHTTP_ERROR_BASE + 18)
  ERROR_WINHTTP_INCORRECT_HANDLE_STATE = (WINHTTP_ERROR_BASE + 19)
  ERROR_WINHTTP_CANNOT_CONNECT = (WINHTTP_ERROR_BASE + 29)
  ERROR_WINHTTP_CONNECTION_ERROR = (WINHTTP_ERROR_BASE + 30)
  ERROR_WINHTTP_RESEND_REQUEST = (WINHTTP_ERROR_BASE + 32)
  ERROR_WINHTTP_SECURE_CERT_DATE_INVALID = (WINHTTP_ERROR_BASE + 37)
  ERROR_WINHTTP_SECURE_CERT_CN_INVALID = (WINHTTP_ERROR_BASE + 38)
  ERROR_WINHTTP_CLIENT_AUTH_CERT_NEEDED = (WINHTTP_ERROR_BASE + 44)
  ERROR_WINHTTP_SECURE_INVALID_CA = (WINHTTP_ERROR_BASE + 45)
  ERROR_WINHTTP_SECURE_CERT_REV_FAILED = (WINHTTP_ERROR_BASE + 57)
  ERROR_WINHTTP_CANNOT_CALL_BEFORE_OPEN = (WINHTTP_ERROR_BASE + 100)
  ERROR_WINHTTP_CANNOT_CALL_BEFORE_SEND = (WINHTTP_ERROR_BASE + 101)
  ERROR_WINHTTP_CANNOT_CALL_AFTER_SEND = (WINHTTP_ERROR_BASE + 102)
  ERROR_WINHTTP_CANNOT_CALL_AFTER_OPEN = (WINHTTP_ERROR_BASE + 103)
  ERROR_WINHTTP_HEADER_NOT_FOUND = (WINHTTP_ERROR_BASE + 150)
  ERROR_WINHTTP_INVALID_SERVER_RESPONSE = (WINHTTP_ERROR_BASE + 152)
  ERROR_WINHTTP_INVALID_HEADER = (WINHTTP_ERROR_BASE + 153)
  ERROR_WINHTTP_INVALID_QUERY_REQUEST = (WINHTTP_ERROR_BASE + 154)
  ERROR_WINHTTP_HEADER_ALREADY_EXISTS = (WINHTTP_ERROR_BASE + 155)
  ERROR_WINHTTP_REDIRECT_FAILED = (WINHTTP_ERROR_BASE + 156)
  ERROR_WINHTTP_SECURE_CHANNEL_ERROR = (WINHTTP_ERROR_BASE + 157)
  ERROR_WINHTTP_BAD_AUTO_PROXY_SCRIPT = (WINHTTP_ERROR_BASE + 166)
  ERROR_WINHTTP_UNABLE_TO_DOWNLOAD_SCRIPT = (WINHTTP_ERROR_BASE + 167)
  ERROR_WINHTTP_SECURE_INVALID_CERT = (WINHTTP_ERROR_BASE + 169)
  ERROR_WINHTTP_SECURE_CERT_REVOKED = (WINHTTP_ERROR_BASE + 170)
  ERROR_WINHTTP_NOT_INITIALIZED = (WINHTTP_ERROR_BASE + 172)
  ERROR_WINHTTP_SECURE_FAILURE = (WINHTTP_ERROR_BASE + 175)
  ERROR_WINHTTP_AUTO_PROXY_SERVICE_ERROR = (WINHTTP_ERROR_BASE + 178)
  ERROR_WINHTTP_SECURE_CERT_WRONG_USAGE = (WINHTTP_ERROR_BASE + 179)
  ERROR_WINHTTP_AUTODETECTION_FAILED = (WINHTTP_ERROR_BASE + 180)
  ERROR_WINHTTP_HEADER_COUNT_EXCEEDED = (WINHTTP_ERROR_BASE + 181)
  ERROR_WINHTTP_HEADER_SIZE_OVERFLOW = (WINHTTP_ERROR_BASE + 182)
  ERROR_WINHTTP_CHUNKED_ENCODING_HEADER_SIZE_OVERFLOW = (WINHTTP_ERROR_BASE + 183)
  ERROR_WINHTTP_RESPONSE_DRAIN_OVERFLOW = (WINHTTP_ERROR_BASE + 184)
  ERROR_WINHTTP_CLIENT_CERT_NO_PRIVATE_KEY = (WINHTTP_ERROR_BASE + 185)
  ERROR_WINHTTP_CLIENT_CERT_NO_ACCESS_PRIVATE_KEY = (WINHTTP_ERROR_BASE + 186)
  WINHTTP_ERROR_LAST = (WINHTTP_ERROR_BASE + 186)
  ERROR_WINHTTP_INVALID_PARAMETER = 87  ' Standard Windows API error code for invalid parameter
End Enum

Private Enum WinHttpRequestOption
    UserAgentString = 0
    Url = 1
    URLCodePage = 2
    EscapePercentInURL = 3
    SslErrorIgnoreFlags = 4
    SelectCertificate = 5
    EnableRedirects = 6
    UrlEscapeDisable = 7
    UrlEscapeDisableQuery = 8
    SecureProtocols = 9
    EnableTracing = 10
    RevertImpersonationOverSsl = 11
    EnableHttpsToHttpRedirects = 12
    EnablePassportAuthentication = 13
    MaxAutomaticRedirects = 14
    MaxResponseHeaderSize = 15
    MaxResponseDrainSize = 16
    EnableHttp1_1 = 17
    EnableCertificateRevocationCheck = 18
End Enum

#If Mac Then
#ElseIf VBA7 Then
  Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDest As LongPtr, ByVal lpSource As LongPtr, ByVal iCount As Long)
  Private Declare PtrSafe Function SysAllocString Lib "oleaut32" (ByVal pwsz As LongPtr) As LongPtr
  Private Declare PtrSafe Sub SysFreeString Lib "oleaut32" (ByVal pwsz As LongPtr)
  Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal p As LongPtr) As LongPtr
  Private Declare PtrSafe Function WinHttpGetIEProxyConfigForCurrentUser Lib "WinHTTP.dll" (ByRef proxyConfig As IE_PROXY_CONFIG) As Long
  Private Declare PtrSafe Function WinHttpGetProxyForUrl Lib "WinHTTP.dll" (ByVal hSession As LongPtr, ByVal pszUrl As LongPtr, ByRef pAutoProxyOptions As AUTOPROXY_OPTIONS, ByRef pProxyInfo As PROXY_INFO) As Long
  Private Declare PtrSafe Function WinHttpOpen Lib "WinHTTP.dll" (ByVal pszUserAgent As LongPtr, ByVal dwAccessType As Long, ByVal pszProxyName As LongPtr, ByVal pszProxyBypass As LongPtr, ByVal dwFlags As Long) As LongPtr
  Private Declare PtrSafe Function WinHttpCloseHandle Lib "WinHTTP.dll" (ByVal hInternet As LongPtr) As Long
  
  Private Type IE_PROXY_CONFIG
      fAutoDetect As Long
      lpszAutoConfigUrl As LongPtr
      lpszProxy As LongPtr
      lpszProxyBypass As LongPtr
  End Type
  Private Type AUTOPROXY_OPTIONS
      dwFlags As Long
      dwAutoDetectFlags As Long
      lpszAutoConfigUrl As LongPtr
      lpvReserved As LongPtr
      dwReserved As Long
      fAutoLogonIfChallenged As Long
  End Type
  Private Type PROXY_INFO
      dwAccessType As Long
      lpszProxy As LongPtr
      lpszProxyBypass As LongPtr
  End Type

#Else

  Private Enum LongPtr
    [_]
  End Enum
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDest As LongPtr, ByVal lpSource As LongPtr, ByVal iCount As Long)
  Private Declare Function SysAllocString Lib "oleaut32" (ByVal pwsz As LongPtr) As LongPtr
  Private Declare Sub SysFreeString Lib "oleaut32" (ByVal pwsz As LongPtr)
  Private Declare Function GlobalFree Lib "kernel32" (ByVal p As LongPtr) As LongPtr
  Private Declare Function WinHttpGetIEProxyConfigForCurrentUser Lib "WinHTTP.dll" (ByRef proxyConfig As IE_PROXY_CONFIG) As Long
  Private Declare Function WinHttpGetProxyForUrl Lib "WinHTTP.dll" (ByVal hSession As LongPtr, ByVal pszUrl As LongPtr, ByRef pAutoProxyOptions As AUTOPROXY_OPTIONS, ByRef pProxyInfo As PROXY_INFO) As Long
  Private Declare Function WinHttpOpen Lib "WinHTTP.dll" (ByVal pszUserAgent As LongPtr, ByVal dwAccessType As Long, ByVal pszProxyName As LongPtr, ByVal pszProxyBypass As LongPtr, ByVal dwFlags As Long) As LongPtr
  Private Declare Function WinHttpCloseHandle Lib "WinHTTP.dll" (ByVal hInternet As LongPtr) As Long
  Private Type IE_PROXY_CONFIG
      fAutoDetect As Long
      lpszAutoConfigUrl As LongPtr
      lpszProxy As LongPtr
      lpszProxyBypass As LongPtr
  End Type
  Private Type AUTOPROXY_OPTIONS
      dwFlags As Long
      dwAutoDetectFlags As Long
      lpszAutoConfigUrl As LongPtr
      lpvReserved As LongPtr
      dwReserved As Long
      fAutoLogonIfChallenged As Long
  End Type
  Private Type PROXY_INFO
      dwAccessType As Long
      lpszProxy As LongPtr
      lpszProxyBypass As LongPtr
  End Type
#End If

Private Type ErrorInfo
  isSet As Boolean
  isRecoverable As Boolean
  Num As Long
  Msg As String
End Type

Private Type ProxyInfoEx
  ProxyServer As String
  ProxyBypass As String
  Error As ErrorInfo
End Type

Private Type TThis
  HTTP as object
  ResponseMapper as stdICallable
End Type
Private This As TThis


'Create a HTTP request
'@param sURL - The url to query
'@param RequestMethod - Typically "GET", "POST", "PUT", "DELETE", but can be any string which the server responds to
'@param ThreadingStyle - The Threading approach. Either HTTPSync (Synchronous) or HTTPAsync (Asynchronous).
'@param options as Object<Dictionary> - Use `stdHTTP.CreateOptions()` to generate these.
'@param Authenticator as stdICallable<(HTTP: Object, RequestMethod: String, URL: String, ThreadingStyle: Long, options: Object)=>Void> - A callable object which will be called before the request is sent. This can be used to add authentication headers, etc.
'@returns - A http request object
Public Function Create(ByVal sURL As String, Optional ByVal RequestMethod As String = "GET", Optional ByVal ThreadingStyle As EHTTPSynchronisity = HTTPSync, Optional ByVal options As Object = Nothing, Optional ByVal Authenticator As stdICallable = Nothing) As stdHTTP
  Set Create = New stdHTTP
  Call Create.protInit(sURL, RequestMethod, ThreadingStyle, options, Authenticator)
End Function

'Create HTTP request headers
'@param v - Headers to load into the http request in the form key1, value1, key2, value2, ...
'@returns Object<Dictionary> - A dictionary of headers
Public Function CreateHeaders(ParamArray v()) As Object
  Dim oRet As Object
  Set oRet = CreateObject("Scripting.Dictionary")
  oRet.CompareMode = vbTextCompare
  
  Dim i As Long
  For i = 0 To UBound(v) Step 2
    oRet(v(i)) = v(i + 1)
  Next
  
  Set CreateHeaders = oRet
End Function

'Create HTTP request options
'@param Body - The body of the request
'@param Headers as Object<Dictionary> - The headers of the request
'@param Cookies as Object<Dictionary> - The cookies of the request
'@param ContentTypeDefault - The default content type of the request
'@param Insecure - Whether to ignore SSL errors
'@param EnableRedirects - Whether to follow redirects
'@param TimeoutMS - The timeout of the request in milliseconds
'@param AutoProxy - Whether to use the system proxy settings
'@returns Object<Dictionary> - A dictionary of options
Public Function CreateOptions(Optional Body As String = "", Optional Headers As Object = Nothing, Optional Cookies As Object = Nothing, Optional ByVal ContentTypeDefault As EHTTPContentType, Optional Insecure As Boolean = False, Optional EnableRedirects As Boolean = True, Optional ByVal TimeoutMS As Long = 5000, Optional ByVal Proxy As Object = Nothing)
  Dim oRet As Object
  Set oRet = CreateObject("Scripting.Dictionary")
  oRet.CompareMode = vbTextCompare
  
  If Headers Is Nothing Then Set Headers = CreateHeaders()
  If Not Headers.Exists("Content-Type") Then
    Dim sContentType As String
    If (ContentTypeDefault And ContentType_HTML) = ContentType_HTML Then sContentType = sContentType & "; text/html"
    If (ContentTypeDefault And ContentType_Plain) = ContentType_Plain Then sContentType = sContentType & "; text/plain"
    If (ContentTypeDefault And ContentType_JSON) = ContentType_JSON Then sContentType = sContentType & "; application/json"
    If (ContentTypeDefault And ContentType_XML) = ContentType_XML Then sContentType = sContentType & "; application/xml; text/xml"
    If (ContentTypeDefault And ContentType_CSV) = ContentType_CSV Then sContentType = sContentType & "; text/csv"
    If (ContentTypeDefault And ContentType_ZIP) = ContentType_ZIP Then sContentType = sContentType & "; application/zip"
    If (ContentTypeDefault And ContentType_Binary) = ContentType_Binary Then sContentType = sContentType & "; application/octet-stream"
    If sContentType <> "" Then Headers("Content-Type") = Mid(sContentType, 3)
  End If
  Set oRet("Headers") = Headers
  
  If Cookies Is Nothing Then Set Cookies = CreateHeaders()
  Set oRet("Cookies") = Cookies
  
  if Proxy is nothing then 
    set oRet("Proxy") = CreateProxy(ProxyTypeAuto)
  else
    set oRet("Proxy") = Proxy
  end if

  oRet("TimeoutMS") = TimeoutMS
  oRet("Insecure") = Insecure
  oRet("EnableRedirects") = EnableRedirects
  oRet("Body") = Body
  Set CreateOptions = oRet
End Function

'Obtain proxy information
'@param proxyType - 
'@param server - the IP address (or hostname) of the proxy server
'@param bypass - A proxy bypass list is not always needed, but it's useful. It specifies URLs or IP ranges that should be accessed directly, bypassing the proxy. This is often used for local or intranet addresses.
'@param user - Username for proxy server authentication
'@param pass - Password for proxy server authentication
'@returns - Proxy information object
Public Function CreateProxy(Byval proxyType as EHTTPProxyType, Optional ByVal Server as string, Optional ByVal ByPass as String, Optional ByVal User as string, Optional ByVal Pass as string)
  Dim oRet As Object
  Set oRet = CreateObject("Scripting.Dictionary")
  oRet.CompareMode = vbTextCompare
  
  select case proxyType
    case ProxyTypeAuto
      'Loads the system proxy settings for a given URL at runtime
      oRet("hasAutoProxy") = true
    case ProxyTypeManual
      oRet("hasAutoProxy") = false
      oRet("hasProxy") = true
      oRet("server") = server
      oRet("bypass") = bypass
      oRet("user") = user
      oRet("pass") = pass
    case ProxyTypeNone
      oRet("hasAutoProxy") = false
      oRet("hasProxy") = false
  end select

  set CreateProxy = oRet
End Function

'Await for the request to finish. This is only necessary for asynchronous requests.
'@returns - The request/response object
Public Function Await() As stdHTTP
  Set Await = Me
  While Not isFinished
    VBA.DoEvents
  Wend
End Function

'Get whether the request has finished. This is only necessary for asynchronous requests.
'@returns - `True` if the request has finished, `False` otherwise.
Public Property Get isFinished() As Boolean
  On Error GoTo Skip 'Sometimes `responseBody` and `status` isn't ready yet. So will skip if so
  if This.HTTP.status = 202 then 'No content
    isFinished = true
  else
    isFinished = This.HTTP.WaitForResponse(0.001) And This.HTTP.ResponseBody()(0) >= 0
  end if
  Exit Property
Skip:
  Debug.Print "Error in stdHTTP#isFinished: '" & err.Description & "'"
End Property

'Get the response status code
'@returns - The response status code
Public Property Get ResponseStatus()
  On Error GoTo Processing
  ResponseStatus = This.HTTP.status
  Exit Property
Processing:
  ResponseStatus = 0
End Property

'Get the response status text
'@returns - The response status text
Public Property Get ResponseStatusText()
  On Error GoTo Skip
  ResponseStatusText = This.HTTP.StatusText
  Exit Property
Skip:
  Dim sErrDescription As String: sErrDescription = Err.Description
  Select Case sErrDescription
    Case "The data necessary to complete this operation is not yet available." & vbCrLf: ResponseStatusText = "Awaiting Response"
    Case Else: ResponseStatusText = "Error: " & sErrDescription
  End Select
End Property

'Get the response text
'@returns - The response text
Public Property Get ResponseText()
  ResponseText = This.HTTP.ResponseText
End Property

'Get the response body
'@returns - The response body
Public Property Get ResponseBody()
  ResponseBody = This.HTTP.ResponseBody
End Property

'Get a response header by name
'@param sHeader - The name of the header to get
'@returns - The response header
Public Property Get ResponseHeader(ByVal sHeader As String) As Variant
  If IsObject(This.HTTP.GetResponseHeader(sHeader)) Then
    Set ResponseHeader = This.HTTP.GetResponseHeader(sHeader)
  Else
    ResponseHeader = This.HTTP.GetResponseHeader(sHeader)
  End If
End Property

'Get all response headers
'@returns - All response headers
Public Property Get ResponseHeaders() As Object
  Set ResponseHeaders = This.HTTP.GetAllResponseHeaders
End Property

'Get the mapped response
Public Property Get ResponseMapped() as variant
  if This.ResponseMapper is nothing then Err.Raise 1, "", "No response mapper provided"
  set ResponseMapped = This.ResponseMapper.Run(Me)
End Property

'Get/Set the response mapper. This callback will run when ResponeMapped() is called. This would usually be used to convert the response to a JSON/XML VBA object.
Public Property Get ResponseMapper() as stdICallable
  set ResponseMapper = this.ResponseMapper
End Property
Public Property Set ResponseMapper(Byval v as stdICallable)
  set This.ResponseMapper = v
End Property

'Initialize the request
'@protected
'@param sURL - The url to query
'@param RequestMethod - Typically "GET", "POST", "PUT", "DELETE", but can be any string which the server responds to
'@param ThreadingStyle - The Threading approach. Either HTTPSync (Synchronous) or HTTPAsync (Asynchronous).
'@param Options as Object<Dictionary> - Use `stdHTTP.CreateOptions()` to generate these.
'@param Authenticator as stdICallable<(HTTP: Object, RequestMethod: String, URL: String, ThreadingStyle: Long, options: Object)=>Void> - A callable object which will be called before the request is sent. This can be used to add authentication headers, etc.
Friend Sub protInit(ByVal sURL As String, Optional ByVal RequestMethod As String = "GET", Optional ByVal ThreadingStyle As EHTTPSynchronisity, Optional ByVal options As Object = Nothing, Optional ByVal Authenticator As stdICallable = Nothing)
  If options Is Nothing Then Set options = CreateOptions()
  Set This.HTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
  This.HTTP.Open RequestMethod, sURL, ThreadingStyle = HTTPAsync
  This.HTTP.SetTimeouts options("TimeoutMS") / 4, options("TimeoutMS") / 4, options("TimeoutMS") / 4, options("TimeoutMS") / 4
  
  'Set proxy info
  With options("Proxy")
    If .item("hasAutoProxy") then Call LoadAutoProxy(sURL, options("Proxy"))
    If .item("hasProxy") Then
      'See http://msdn.microsoft.com/en-us/library/windows/desktop/aa384059(v=vs.85).aspx for details
      This.HTTP.SetProxy ProxySetting.Proxy, .item("Server"), .item("Bypass")
      If .item("User") <> vbNullString then
        This.HTTP.SetCredentials .item("User"), .item("Pass"), SetCredentialsType.ForProxy
      end if
    else
      'Use Proxycfg.exe
      This.HTTP.SetProxy ProxySetting.PreConfig
    end if
  end with

  If options("Insecure") Then
    This.HTTP.Option(WinHttpRequestOption.EnableCertificateRevocationCheck) = False
    This.HTTP.Option(WinHttpRequestOption.SslErrorIgnoreFlags) = 13056
    This.HTTP.Option(WinHttpRequestOption.EnableHttpsToHttpRedirects) = True
  Else
    This.HTTP.Option(WinHttpRequestOption.EnableCertificateRevocationCheck) = True
    This.HTTP.Option(WinHttpRequestOption.SslErrorIgnoreFlags) = 0
    This.HTTP.Option(WinHttpRequestOption.EnableHttpsToHttpRedirects) = False
  End If
  This.HTTP.Option(WinHttpRequestOption.EnableRedirects) = options("EnableRedirects")
  
  Dim vKey
  For Each vKey In options("Headers").keys()
    This.HTTP.SetRequestHeader vKey, options("Headers")(vKey)
  Next
  For Each vKey In options("Cookies").keys()
    This.HTTP.SetRequestHeader "Cookie", vKey & "=" & options("Cookies")(vKey)
  Next
  
  'Allow authenticators
  If Not Authenticator Is Nothing Then Call Authenticator.Run(This.HTTP, RequestMethod, sURL, ThreadingStyle, options)
  
  'Execute request
  This.HTTP.Send options("Body")
End Sub

'Get the HSesssion
'@protected
'@returns - The HSesssion
#If VBA7 then
Friend Property Get protHSession() As LongPtr
#else
Friend Property Get protHSession() As Long
#end if
  If Me Is stdHTTP Then
    Static hSession As LongPtr: If hSession = 0 Then hSession = WinHttpOpen(0, 1, 0, 0, 0)
    protHSession = hSession
  Else
    protHSession = stdHTTP.protHSession
  End If
End Property

'Encode a string for use in a URL
'@param strInput - The string to encode
'@returns - The encoded string
Public Function URIEncode(ByVal strInput As String) As String
  Dim bytInput() As Byte
  Dim strOutput As String
  Dim i As Long

  bytInput = StrConv(strInput, vbFromUnicode)

  strOutput = ""
  For i = 0 To UBound(bytInput)
    If (bytInput(i) >= 48 And bytInput(i) <= 57) Or _
       (bytInput(i) >= 65 And bytInput(i) <= 90) Or _
       (bytInput(i) >= 97 And bytInput(i) <= 122) Then
      strOutput = strOutput & Chr(bytInput(i))
    Else
      strOutput = strOutput & "%" & right("0" & Hex(bytInput(i)), 2)
    End If
  Next i

  URIEncode = strOutput
End Function

'Decode a string from a URL
'@param strInput - The string to decode
'@returns - The decoded string
Public Function URIDecode(ByVal strInput As String) As String
    Dim strOutput As String: strOutput = ""
    Dim intPos As Long: intPos = 1
    Do Until intPos > Len(strInput)
        If Mid$(strInput, intPos, 1) = "%" Then
            strOutput = strOutput & Chr(CLng("&H" & Mid$(strInput, intPos + 1, 2)))
            intPos = intPos + 3
        Else
            strOutput = strOutput & Mid$(strInput, intPos, 1)
            intPos = intPos + 1
        End If
    Loop
    URIDecode = strOutput
End Function












'  t = t & "^(?:(http[s]?|ftp):/)?/?"  ' METHOD
'  t = t & "([^:^/^?^#\\s]+)"          ' HOSTNAME
'  t = t & "(?::(\\d+))?"              ' PORT
'  t = t & "([^?^#.*]+)?"              ' PATH
'  t = t & "(\\?[^#.]*)?"              ' QUERY
'  t = t & "(#[\\w\\-]+)?$"            ' ID

'Load the system proxy settings for a given URL
'@param sURL - The url to query
'@param oProxy - The proxy object to load the settings into
Private Sub LoadAutoProxy(ByVal sURL as string, ByVal oProxy as Object)
  Dim info As ProxyInfoEx: info = GetProxyCurrentUser(sURL)
  If Not info.Error.isSet Then
    oProxy("hasProxy") = true
    oProxy("server") = info.ProxyServer
    oProxy("bypass") = info.ProxyBypass
    oProxy("user") = Empty  'TODO: Check if user/pass credentials with auto-proxy will ever be needed
    oProxy("pass") = Empty  
  else
    if info.Error.isRecoverable then
      'Assume no proxy
      oProxy("hasProxy") = false
    else
      Err.Raise info.Error.Num, "stdHTTP", info.Error.Msg 
    end if
  end if
End Sub

'Get the system proxy settings for a given URL
'@param sURL - The url to query
'@returns - The proxy information struct
Private Function GetProxyCurrentUser(ByVal sURL As String) As ProxyInfoEx
  ' Constants for dwFlags of AUTOPROXY_OPTIONS
  Const AUTO_DETECT = 1
  Const CONFIG_URL = 2
  
  ' Constants for dwAutoDetectFlags
  Const DETECT_TYPE_DHCP = 1
  Const DETECT_TYPE_DNS = 2
  
  Dim AutoProxyOptions As AUTOPROXY_OPTIONS
  AutoProxyOptions.fAutoLogonIfChallenged = 1
  
  ' WinHttpGetProxyForUrl returns unexpected errors if Url is empty
  If sURL = "" Then sURL = " "
  
  On Error GoTo Cleanup
  
  'Check IE's proxy configuration
  Dim IEProxyConfig As IE_PROXY_CONFIG, doAutoProxy As Boolean
  If WinHttpGetIEProxyConfigForCurrentUser(IEProxyConfig) > 0 Then
    'If IE is configured to auto-detect, then we will too.
    If IEProxyConfig.fAutoDetect <> 0 Then
      With AutoProxyOptions
        .dwFlags = AUTO_DETECT
        .dwAutoDetectFlags = DETECT_TYPE_DHCP Or DETECT_TYPE_DNS
      End With
      doAutoProxy = True
    End If
    
    'If IE is configured to use an auto-config script, then we will use it too
    If (IEProxyConfig.lpszAutoConfigUrl <> 0) Then
      With AutoProxyOptions
        .dwFlags = .dwFlags Or CONFIG_URL
        .lpszAutoConfigUrl = IEProxyConfig.lpszAutoConfigUrl
      End With
      doAutoProxy = True
    End If
  Else
    ' If the IE proxy config is not available, then we will try auto-detection
    With AutoProxyOptions
      .dwFlags = AUTO_DETECT
      .dwAutoDetectFlags = DETECT_TYPE_DHCP Or DETECT_TYPE_DNS
    End With
    doAutoProxy = True
  End If
  
  If doAutoProxy Then
    On Error GoTo TryIEFallback
    
    ' Get proxy information
    Dim ProxyInfo As PROXY_INFO
    If WinHttpGetProxyForUrl(protHSession, StrPtr(sURL), AutoProxyOptions, ProxyInfo) > 0 Then
      Dim ProxyStringPtr As LongPtr
      ProxyStringPtr = ProxyInfo.lpszProxy
    Else
      Dim ProxyErrorNum As WINHTTP_ERROR, ProxyErrorMsg As String
      ProxyErrorNum = Err.LastDllError
      GetProxyCurrentUser.Error.isSet = True
      GetProxyCurrentUser.Error.Num = ProxyErrorNum
      Select Case ProxyErrorNum
        Case ERROR_INVALID_PARAMETER, ERROR_WINHTTP_AUTO_PROXY_SERVICE_ERROR, ERROR_WINHTTP_AUTODETECTION_FAILED, ERROR_WINHTTP_BAD_AUTO_PROXY_SCRIPT, _
             ERROR_WINHTTP_LOGIN_FAILURE, ERROR_WINHTTP_OPERATION_CANCELLED, ERROR_WINHTTP_TIMEOUT, ERROR_WINHTTP_UNABLE_TO_DOWNLOAD_SCRIPT, ERROR_WINHTTP_UNRECOGNIZED_SCHEME
          GetProxyCurrentUser.Error.isRecoverable = True
      End Select
      Select Case ProxyErrorNum
        Case ERROR_WINHTTP_AUTODETECTION_FAILED
          GetProxyCurrentUser.Error.Msg = "WPAD detection failed"
        Case ERROR_WINHTTP_UNABLE_TO_DOWNLOAD_SCRIPT
          GetProxyCurrentUser.Error.Msg = "Unable to download proxy auto-config script"
        Case ERROR_WINHTTP_BAD_AUTO_PROXY_SCRIPT
          GetProxyCurrentUser.Error.Msg = "Error in proxy auto-config script"
        Case ERROR_WINHTTP_AUTO_PROXY_SERVICE_ERROR
          GetProxyCurrentUser.Error.Msg = "No proxy can be located for the specified URL"
        Case ERROR_WINHTTP_INVALID_URL, ERROR_WINHTTP_UNRECOGNIZED_SCHEME
          GetProxyCurrentUser.Error.Msg = "Specified URL is not valid"
        Case Else
          GetProxyCurrentUser.Error.Msg = "Unknown error while detecting proxy"
      End Select
    End If
  End If
  
TryIEFallback:
  On Error GoTo Cleanup
  
  'If we don't have a proxy server from WinHTTPGetProxyForUrl then pick one up from the IE proxy config (if given)
  If ProxyStringPtr = 0 Then ProxyStringPtr = IEProxyConfig.lpszProxy
  
  Dim ptr As LongPtr
  
  'If theres a proxy string, convert it to a VBA string
  If ProxyStringPtr <> 0 Then
    ptr = SysAllocString(ProxyStringPtr)
    Call CopyMemory(VarPtr(GetProxyCurrentUser.ProxyServer), VarPtr(ptr), PTR_SIZE)
    Call SysFreeString(ptr)
  End If
  
  'Pick up any bypass string from IEProxyConfig
  If IEProxyConfig.lpszProxyBypass <> 0 Then
    ptr = SysAllocString(IEProxyConfig.lpszProxyBypass)
    Call CopyMemory(VarPtr(GetProxyCurrentUser.ProxyBypass), VarPtr(ptr), PTR_SIZE)
    Call SysFreeString(ptr)
  End If
    
Cleanup:
  'Cleanup proxy config
  With IEProxyConfig
    If .lpszAutoConfigUrl <> 0 Then Call GlobalFree(.lpszAutoConfigUrl)
    If .lpszProxy <> 0 Then Call GlobalFree(.lpszProxy)
    If .lpszProxyBypass <> 0 Then Call GlobalFree(.lpszProxyBypass)
  End With
End Function
