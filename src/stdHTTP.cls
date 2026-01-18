VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdHTTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@module
'@description A class used to make HTTP requests.
'@example A simple synchronous HTTP request:
'```vb
'Dim req as stdHTTP: set req = stdHTTP.Create("http://www.google.com")
'Dim resp as string: resp = req.ResponseText
'```
'@example A simple asynchronous HTTP request:
'```vb
'Dim req as stdHTTP: set req = stdHTTP.Create("http://www.google.com", ThreadingStyle:=HTTPAsync)
'req.Await
'Dim resp as string: resp = req.ResponseText
'```
'@example A simple HTTP request with a timeout:
'```vb
'Dim req as stdHTTP: set req = stdHTTP.Create("http://www.google.com", TimeoutMS:=1000)
'Dim resp as string: resp = req.ResponseText
'```

'@example A simple HTTP request with authentication:
'```vb
'Sub main()
'  Dim req as stdHTTP: set req = stdHTTP.Create( _
'    "http://www.google.com", _
'    Authenticator:=stdCallback.CreateFromModule("mMain", "HttpBasicAuthenticator") _
'      .bind("username","password") _
'  )
'  Debug.Print req.ResponseText
'End Sub
'Public Sub HttpBasicAuthenticator(ByVal Username As String, ByVal Password As String, ByVal pHTTP As Object, ByVal RequestMethod As String, ByVal sURL As String, ByVal ThreadingStyle As Long, ByVal options As Object)
'  Const SetCredentialsType_ForServer = 0
'  pHTTP.SetCredentials Username, Password, SetCredentialsType_ForServer
'End Sub
'```
'@example A simple POST request with a payload:
'```vb
'Dim req as stdHTTP: set req = stdHTTP.Create( _
'  "http://www.google.com", _
'  Options:=stdHTTP.CreateOptions( _
'    Body:="My payload here", _
'  ) _
')  
'Dim resp as string: resp = req.ResponseText
'```
'@example A simple HTTP request with a mapped response:
'```vb
'Dim req as stdHTTP: set req = stdHTTP.Create( "http://www.google.com", ThreadingStyle:=HTTPAsync)
'set req.ResponseMapper = stdLambda.Create("stdJSON.CreateFromString($1.ResponseText)").bindGlobal("stdJSON", stdJSON)
'
'Dim resp as stdJSON: set resp = req.Await.ResponseMapped
'```
'@example A simple HTTP request with a proxy:
'```vb
'Dim req as stdHTTP: set req = stdHTTP.Create( _
'  "http://www.google.com", _
'  Options:=stdHTTP.CreateOptions( _
'    Proxy:=stdHTTP.CreateProxy( _
'      ProxyTypeAuto, _
'      Server:="127.0.0.1", _
'      ByPass:="localhost", _
'      User:="user", _
'      Pass:="pass" _
'    ) _ 
'  ) _
')  
'Dim resp as string: resp = req.ResponseText
'```

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

Public Enum stdHTTPOptions
  stdHTTPOptionBody = 0
  stdHTTPOptionHeaders
  stdHTTPOptionCookies
  stdHTTPOptionProxy
  stdHTTPOptionTimeoutMS
  stdHTTPOptionInsecure
  stdHTTPOptionEnableRedirects
  [_stdHTTPOptionLast] = stdHTTPOptionEnableRedirects
End Enum

Public Enum stdHTTPProxy
  stdHTTPProxyHasAutoProxy = 0
  stdHTTPProxyHasProxy
  stdHTTPProxyServer
  stdHTTPProxyBypass
  stdHTTPProxyUser
  stdHTTPProxyPass
  [_stdHTTPProxyLast] = stdHTTPProxyPass
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
  #If VBA7 Then
    Private Declare PtrSafe Function web_popen Lib "/usr/lib/libc.dylib" Alias "popen" (ByVal web_Command As String, ByVal web_Mode As String) As LongPtr
    Private Declare PtrSafe Function web_pclose Lib "/usr/lib/libc.dylib" Alias "pclose" (ByVal web_File As LongPtr) As LongPtr
    Private Declare PtrSafe Function web_fread Lib "/usr/lib/libc.dylib" Alias "fread" (ByVal web_OutStr As String, ByVal web_Size As LongPtr, ByVal web_Items As LongPtr, ByVal web_Stream As LongPtr) As LongPtr
    Private Declare PtrSafe Function web_feof Lib "/usr/lib/libc.dylib" Alias "feof" (ByVal web_File As LongPtr) As LongPtr
  #Else
    Private Declare Function web_popen Lib "libc.dylib" Alias "popen" (ByVal web_Command As String, ByVal web_Mode As String) As Long
    Private Declare Function web_pclose Lib "libc.dylib" Alias "pclose" (ByVal web_File As Long) As Long
    Private Declare Function web_fread Lib "libc.dylib" Alias "fread" (ByVal web_OutStr As String, ByVal web_Size As Long, ByVal web_Items As Long, ByVal web_Stream As Long) As Long
    Private Declare Function web_feof Lib "libc.dylib" Alias "feof" (ByVal web_File As Long) As Long
  #End If
#Else
  #If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDest As LongPtr, ByVal lpSource As LongPtr, ByVal iCount As Long)
    Private Declare PtrSafe Function SysAllocString Lib "oleaut32" (ByVal pwsz As LongPtr) As LongPtr
    Private Declare PtrSafe Sub SysFreeString Lib "oleaut32" (ByVal pwsz As LongPtr)
    Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal p As LongPtr) As LongPtr
    Private Declare PtrSafe Function WinHttpGetIEProxyConfigForCurrentUser Lib "WinHTTP.dll" (ByRef proxyConfig As IE_PROXY_CONFIG) As Long
    Private Declare PtrSafe Function WinHttpGetProxyForUrl Lib "WinHTTP.dll" (ByVal hSession As LongPtr, ByVal pszUrl As LongPtr, ByRef pAutoProxyOptions As AUTOPROXY_OPTIONS, ByRef pProxyInfo As PROXY_INFO) As Long
    Private Declare PtrSafe Function WinHttpOpen Lib "WinHTTP.dll" (ByVal pszUserAgent As LongPtr, ByVal dwAccessType As Long, ByVal pszProxyName As LongPtr, ByVal pszProxyBypass As LongPtr, ByVal dwFlags As Long) As LongPtr
    Private Declare PtrSafe Function WinHttpCloseHandle Lib "WinHTTP.dll" (ByVal hInternet As LongPtr) As Long
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
  #End If
#End If

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

#If Mac Then
  Private Type HeaderPair
    Key As String
    Value As String
  End Type

  Private Type TResponse
    Status As Long
    StatusText As String
    Text As String
    Body() As Byte
    Headers() As HeaderPair
    HeadersCount As Long
    isComplete As Boolean
  End Type
#End If

Private Type TThis
  HTTP As Object
  ResponseMapper As stdICallable
  Created as Date
  TimeoutMS as Long
  #If Mac Then
    #If VBA7 Then
      CurlProcess As LongPtr
    #Else
      CurlProcess As Long
    #End If
    Response As TResponse
    curlCommand As String
  #End If
End Type
Private This As TThis


'Create a HTTP request
'@constructor
'@param sURL - The url to query
'@param RequestMethod - Typically "GET", "POST", "PUT", "DELETE", but can be any string which the server responds to
'@param ThreadingStyle - The Threading approach. Either HTTPSync (Synchronous) or HTTPAsync (Asynchronous).
'@param options as Variant<Array> - Use `stdHTTP.CreateOptions()` to generate these. Array indexed by stdHTTPOptions enum.
'@param Authenticator as stdICallable<(HTTP: Object, RequestMethod: String, URL: String, ThreadingStyle: Long, options: Variant)=>Void> - A callable object which will be called before the request is sent. This can be used to add authentication headers, etc.
'@returns - A http request object
Public Function Create(ByVal sURL As String, Optional ByVal RequestMethod As String = "GET", Optional ByVal ThreadingStyle As EHTTPSynchronisity = HTTPSync, Optional ByVal options As Variant = Nothing, Optional ByVal Authenticator As stdICallable = Nothing) As stdHTTP
  With stdError.getSentry("stdHTTP#Create", "sURL", sURL, "RequestMethod", RequestMethod, "ThreadingStyle", ThreadingStyle, "options", options, "Authenticator", Authenticator)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Set Create = New stdHTTP
      Call Create.protInit(sURL, RequestMethod, ThreadingStyle, options, Authenticator)
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Create HTTP request headers
'@static
'@param v - Headers to load into the http request in the form key1, value1, key2, value2, ...
'@returns Array - An array of headers in the form Array(key1, value1, key2, value2, ...)
Public Function CreateHeaders(ParamArray v()) As Variant
  With stdError.getSentry("stdHTTP#CreateHeaders")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      CreateHeaders = v
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Create HTTP request options
'@static
'@param Body - The body of the request
'@param Headers as Variant<Array> - The headers of the request (array format: key1, value1, key2, value2, ...)
'@param Cookies as Variant<Array> - The cookies of the request (array format: key1, value1, key2, value2, ...)
'@param ContentTypeDefault - The default content type of the request
'@param Insecure - Whether to ignore SSL errors
'@param EnableRedirects - Whether to follow redirects
'@param TimeoutMS - The timeout of the request in milliseconds
'@param Proxy as Variant<Array> - Proxy configuration array (use `stdHTTP.CreateProxy()` to generate)
'@returns Variant<Array> - A fixed-width array indexed by stdHTTPOptions enum
Public Function CreateOptions(Optional Body As String = "", Optional Headers As Variant = Empty, Optional Cookies As Variant = Empty, Optional ByVal ContentTypeDefault As EHTTPContentType, Optional Insecure As Boolean = False, Optional EnableRedirects As Boolean = True, Optional ByVal TimeoutMS As Long = 5000, Optional ByVal Proxy As Variant = Empty)
  With stdError.getSentry("stdHTTP#CreateOptions", "Body", Body, "Headers", Headers, "Cookies", Cookies, "ContentTypeDefault", ContentTypeDefault, "Insecure", Insecure, "EnableRedirects", EnableRedirects, "TimeoutMS", TimeoutMS, "Proxy", Proxy)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim options() As Variant
      ReDim options(0 To [_stdHTTPOptionLast])
      
      If IsEmpty(Headers) Then Headers = CreateHeaders()
      If Not HeaderArrayExists(Headers, "Content-Type") Then
        Dim sContentType As String
        If (ContentTypeDefault And ContentType_HTML) = ContentType_HTML Then sContentType = sContentType & "; text/html"
        If (ContentTypeDefault And ContentType_Plain) = ContentType_Plain Then sContentType = sContentType & "; text/plain"
        If (ContentTypeDefault And ContentType_JSON) = ContentType_JSON Then sContentType = sContentType & "; application/json"
        If (ContentTypeDefault And ContentType_XML) = ContentType_XML Then sContentType = sContentType & "; application/xml; text/xml"
        If (ContentTypeDefault And ContentType_CSV) = ContentType_CSV Then sContentType = sContentType & "; text/csv"
        If (ContentTypeDefault And ContentType_ZIP) = ContentType_ZIP Then sContentType = sContentType & "; application/zip"
        If (ContentTypeDefault And ContentType_Binary) = ContentType_Binary Then sContentType = sContentType & "; application/octet-stream"
        If sContentType <> "" Then Call HeaderArraySet(Headers, "Content-Type", Mid(sContentType, 3))
      End If
      
      If IsEmpty(Cookies) Then Cookies = CreateHeaders()
      
      If IsEmpty(Proxy) Then Proxy = CreateProxy(ProxyTypeAuto)
      
      'Set options using enum indices
      options(stdHTTPOptionBody) = Body
      options(stdHTTPOptionHeaders) = Headers
      options(stdHTTPOptionCookies) = Cookies
      options(stdHTTPOptionProxy) = Proxy
      options(stdHTTPOptionTimeoutMS) = TimeoutMS
      options(stdHTTPOptionInsecure) = Insecure
      options(stdHTTPOptionEnableRedirects) = EnableRedirects
      
      CreateOptions = options
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Obtain proxy information
'@static
'@param proxyType - The type of proxy configuration
'@param server - the IP address (or hostname) of the proxy server
'@param bypass - A proxy bypass list is not always needed, but it's useful. It specifies URLs or IP ranges that should be accessed directly, bypassing the proxy. This is often used for local or intranet addresses.
'@param user - Username for proxy server authentication
'@param pass - Password for proxy server authentication
'@returns Variant<Array> - A fixed-width array indexed by stdHTTPProxy enum
Public Function CreateProxy(ByVal proxyType As EHTTPProxyType, Optional ByVal Server As String, Optional ByVal ByPass As String, Optional ByVal User As String, Optional ByVal Pass As String)
  With stdError.getSentry("stdHTTP#CreateProxy", "proxyType", proxyType, "Server", Server, "ByPass", ByPass, "User", User, "Pass", Pass)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim proxy() As Variant
      ReDim proxy(0 To [_stdHTTPProxyLast])
      
      Select Case proxyType
        Case ProxyTypeAuto
          'Loads the system proxy settings for a given URL at runtime
          CreateProxy = Array(True, False, "", "", "", "")
        Case ProxyTypeManual
          CreateProxy = Array(False, True, Server, ByPass, User, Pass)
        Case ProxyTypeNone
          CreateProxy = Array(False, False, "", "", "", "")
      End Select
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Await for the request to finish. This is only necessary for asynchronous requests.
'@returns - The request/response object
Public Function Await() As stdHTTP
  With stdError.getSentry("stdHTTP#Await")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Set Await = Me
      While Not isFinished
        if This.TimeoutMS > 0 Then
          If DateDiff("s", This.Created, Now()) > This.TimeoutMS / 1000 Then
            Err_Raise 10000, "stdHTTP", "Timeout"
            Exit Function
          End If
        End If
        VBA.DoEvents
      Wend
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Get whether the request has finished. This is only necessary for asynchronous requests.
'@returns - `True` if the request has finished, `False` otherwise.
Public Property Get isFinished() As Boolean
  With stdError.getSentry("stdHTTP#isFinished[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      #If Mac Then
        If Not This.Response.isComplete And This.CurlProcess <> 0 Then
          'Check if process is still running
          If web_feof(This.CurlProcess) <> 0 Then
            'Process finished, read response
            Call ReadCurlResponse
          End If
        End If
        isFinished = This.Response.isComplete
      #Else
        On Error GoTo Skip 'Sometimes `responseBody` and `status` isn't ready yet. So will skip if so
        If This.HTTP.status = 202 Then 'No content
          isFinished = True
        Else
          isFinished = This.HTTP.WaitForResponse(0.001) And This.HTTP.ResponseBody()(0) >= 0
        End If
        Exit Property
    Skip:
        If Err.Description Like "The data necessary to complete this operation is not yet available*" Then Exit Property
        
        Debug.Print "Error in stdHTTP#isFinished: '" & Err.Description & "'"
      #End If
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Get the response status code
'@returns - The response status code
Public Property Get ResponseStatus()
  With stdError.getSentry("stdHTTP#ResponseStatus[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      #If Mac Then
        ResponseStatus = This.Response.Status
      #Else
        On Error GoTo Processing
        ResponseStatus = This.HTTP.status
        Exit Property
    Processing:
        ResponseStatus = 0
      #End If
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Get the response status text
'@returns - The response status text
Public Property Get ResponseStatusText()
  With stdError.getSentry("stdHTTP#ResponseStatusText[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      #If Mac Then
        If Not This.Response.isComplete Then
          ResponseStatusText = "Awaiting Response"
        Else
          ResponseStatusText = This.Response.StatusText
        End If
      #Else
        On Error GoTo Skip
        ResponseStatusText = This.HTTP.StatusText
        Exit Property
    Skip:
        Dim sErrDescription As String: sErrDescription = Err.Description
        Select Case sErrDescription
          Case "The data necessary to complete this operation is not yet available." & vbCrLf: ResponseStatusText = "Awaiting Response"
          Case Else: ResponseStatusText = "Error: " & sErrDescription
        End Select
      #End If
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Get the response text
'@returns - The response text
Public Property Get ResponseText()
  With stdError.getSentry("stdHTTP#ResponseText[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      #If Mac Then
        ResponseText = This.Response.Text
      #Else
        ResponseText = This.HTTP.ResponseText
      #End If
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Get the response body
'@returns - The response body
Public Property Get ResponseBody()
  With stdError.getSentry("stdHTTP#ResponseBody[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      #If Mac Then
        ResponseBody = This.Response.Body
      #Else
        ResponseBody = This.HTTP.ResponseBody
      #End If
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Get a response header by name
'@param sHeader - The name of the header to get
'@returns - The response header
Public Property Get ResponseHeader(ByVal sHeader As String) As Variant
  With stdError.getSentry("stdHTTP#ResponseHeader[Get]", "sHeader", sHeader)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      #If Mac Then
        ResponseHeader = GetMacHeader(sHeader)
      #Else
        If IsObject(This.HTTP.GetResponseHeader(sHeader)) Then
          Set ResponseHeader = This.HTTP.GetResponseHeader(sHeader)
        Else
          ResponseHeader = This.HTTP.GetResponseHeader(sHeader)
        End If
      #End If
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Get all response headers
'@returns - All response headers as a JSON string in the format {"header1": "value1", ...}
Public Property Get ResponseHeaders() As String
  With stdError.getSentry("stdHTTP#ResponseHeaders[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim sHeaders as string: sHeaders = "{"
      #If Mac Then
        Dim i as long
        For i = 0 To This.Response.HeadersCount - 1
          Dim headerName As String: headerName =  This.Response.Headers(i).Key
          Dim headerValue As String: headerValue = This.Response.Headers(i).Value
          sHeaders = sHeaders & vbCrLf & "  """ & headerName & """:" & JSONLiteral(headerValue)
          if i < This.Response.HeadersCount - 1 then sHeaders = sHeaders & ","
        Next
      #Else
        Dim headersText As String: headersText = This.HTTP.GetAllResponseHeaders
        Dim lines() As String: lines = Split(headersText, vbCrLf)
        Dim i As Long  
        For i = LBound(lines) To UBound(lines)
          Dim headerLine As String: headerLine = Trim(lines(i))
          If headerLine <> "" Then
            Dim colonPos As Long: colonPos = InStr(1, headerLine, ":")
            If colonPos > 0 Then
              Dim headerName As String: headerName = Trim(Left(headerLine, colonPos - 1))
              Dim headerValue As String: headerValue = Trim(Mid(headerLine, colonPos + 1))
              sHeaders = sHeaders & vbCrLf & "  """ & headerName & """:" & JSONLiteral(headerValue)
              if i < UBound(lines) then sHeaders = sHeaders & ","
            End If
          End If
        Next
        sHeaders = sHeaders & vbCrLf & "}"
        ResponseHeaders = sHeaders
      #End If
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Convert headers array to JSON string
'Headers array format: key1, value1, key2, value2, ...
'@protected
Private Function JSONLiteral(ByVal value As String) As String
  With stdError.getSentry("stdHTTP#JSONLiteral", "value", value)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      if isNumeric(value) then
        JSONLiteral = value
      Else
        JSONLiteral = """" & value & """"
      End If
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Get a header value from a header array
'@param headers - The header array (from CreateHeaders), passed by reference
'@param HeaderName - The name of the header
'@returns - The header value, or empty string if not found
'@example
'```vb
'Dim headers: headers = stdHTTP.CreateHeaders()
'Call stdHTTP.SetHeader(headers, "Content-Type", "application/json")
'Debug.Print stdHTTP.GetHeader(headers, "Content-Type")
'```
Public Function GetHeader(ByRef headers As Variant, ByVal HeaderName As String) As String
  With stdError.getSentry("stdHTTP#GetHeader", "headers", headers, "HeaderName", HeaderName)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      GetHeader = HeaderArrayGet(headers, HeaderName)
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Set a header value in a header array
'@param headers - The header array (from CreateHeaders), passed by reference
'@param HeaderName - The name of the header
'@param value - The header value to set
'@example
'```vb
'Dim headers: headers = stdHTTP.CreateHeaders()
'Call stdHTTP.SetHeader(headers, "Content-Type", "application/json")
'```
Public Sub SetHeader(ByRef headers As Variant, ByVal HeaderName As String, ByVal value As String)
  With stdError.getSentry("stdHTTP#SetHeader", "headers", headers, "HeaderName", HeaderName, "value", value)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Call HeaderArraySet(headers, HeaderName, value)
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Get the mapped response
Public Property Get ResponseMapped() As Variant
  With stdError.getSentry("stdHTTP#ResponseMapped[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If This.ResponseMapper Is Nothing Then Err_Raise 1, "", "No response mapper provided"
      Dim oMe As Object: Set oMe = Me
      Call CopyVariant(ResponseMapped, This.ResponseMapper.Run(oMe))
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Get/Set the response mapper. This callback will run when ResponeMapped() is called. This would usually be used to convert the response to a JSON/XML VBA object.
Public Property Get ResponseMapper() As stdICallable
  With stdError.getSentry("stdHTTP#ResponseMapper[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Set ResponseMapper = This.ResponseMapper
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property
Public Property Set ResponseMapper(ByVal v As stdICallable)
  With stdError.getSentry("stdHTTP#ResponseMapper[Set]", "v", v)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Set This.ResponseMapper = v
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Initialize the request
'@protected
'@param sURL - The url to query
'@param RequestMethod - Typically "GET", "POST", "PUT", "DELETE", but can be any string which the server responds to
'@param ThreadingStyle - The Threading approach. Either HTTPSync (Synchronous) or HTTPAsync (Asynchronous).
'@param Options as Variant<Array> - Use `stdHTTP.CreateOptions()` to generate these. Array indexed by stdHTTPOptions enum.
'@param Authenticator as stdICallable<(HTTP: Object, RequestMethod: String, URL: String, ThreadingStyle: Long, options: Variant)=>Void> - A callable object which will be called before the request is sent. This can be used to add authentication headers, etc.
Friend Sub protInit(ByVal sURL As String, Optional ByVal RequestMethod As String = "GET", Optional ByVal ThreadingStyle As EHTTPSynchronisity, Optional ByVal options As Variant = Nothing, Optional ByVal Authenticator As stdICallable = Nothing)
  With stdError.getSentry("stdHTTP#protInit", "sURL", sURL, "RequestMethod", RequestMethod, "ThreadingStyle", ThreadingStyle, "options", options, "Authenticator", Authenticator)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      #If Mac Then
        Call protInitMac(sURL, RequestMethod, ThreadingStyle, options, Authenticator)
      #Else
        If IsEmpty(options) Or options Is Nothing Then options = CreateOptions()
        Set This.HTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
        This.HTTP.Open RequestMethod, sURL, True
        Dim timeoutMS As Long: timeoutMS = options(stdHTTPOptionTimeoutMS)
        This.HTTP.SetTimeouts timeoutMS / 4, timeoutMS / 4, timeoutMS / 4, timeoutMS / 4
        
        'Set proxy info
        Dim proxy As Variant: proxy = options(stdHTTPOptionProxy)
        If proxy(stdHTTPProxyHasAutoProxy) Then Call LoadAutoProxy(sURL, proxy)
        If proxy(stdHTTPProxyHasProxy) Then
          'See http://msdn.microsoft.com/en-us/library/windows/desktop/aa384059(v=vs.85).aspx for details
          This.HTTP.SetProxy ProxySetting.Proxy, proxy(stdHTTPProxyServer), proxy(stdHTTPProxyBypass)
          If proxy(stdHTTPProxyUser) <> vbNullString Then
            This.HTTP.SetCredentials proxy(stdHTTPProxyUser), proxy(stdHTTPProxyPass), SetCredentialsType.ForProxy
          End If
        Else
          'Use Proxycfg.exe
          This.HTTP.SetProxy ProxySetting.PreConfig
        End If
    
        If options(stdHTTPOptionInsecure) Then
          This.HTTP.Option(WinHttpRequestOption.EnableCertificateRevocationCheck) = False
          This.HTTP.Option(WinHttpRequestOption.SslErrorIgnoreFlags) = 13056
          This.HTTP.Option(WinHttpRequestOption.EnableHttpsToHttpRedirects) = True
        Else
          This.HTTP.Option(WinHttpRequestOption.EnableCertificateRevocationCheck) = True
          This.HTTP.Option(WinHttpRequestOption.SslErrorIgnoreFlags) = 0
          This.HTTP.Option(WinHttpRequestOption.EnableHttpsToHttpRedirects) = False
        End If
        This.HTTP.Option(WinHttpRequestOption.EnableRedirects) = options(stdHTTPOptionEnableRedirects)
        
        Dim headers As Variant: headers = options(stdHTTPOptionHeaders)
        Dim i As Long
        If IsArray(headers) Then
          For i = LBound(headers) To UBound(headers) Step 2
            This.HTTP.SetRequestHeader headers(i), headers(i + 1)
          Next
        End If
        
        Dim cookies As Variant: cookies = options(stdHTTPOptionCookies)
        If IsArray(cookies) Then
          For i = LBound(cookies) To UBound(cookies) Step 2
            This.HTTP.SetRequestHeader "Cookie", cookies(i) & "=" & cookies(i + 1)
          Next
        End If
        
        'Allow authenticators
        If Not Authenticator Is Nothing Then Call Authenticator.Run(This.HTTP, RequestMethod, sURL, ThreadingStyle, options)
        
        'Execute request
        This.Created = Now()
        This.TimeoutMS = timeoutMS
        This.HTTP.Send options(stdHTTPOptionBody)
    
        'Await for the request to finish if synchronous
        if ThreadingStyle = HTTPSync Then Call Await
      #End If
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Helper functions for header arrays
'@protected
Private Function HeaderArrayExists(ByRef headers As Variant, ByVal key As String) As Boolean
  With stdError.getSentry("stdHTTP#HeaderArrayExists", "headers", headers, "key", key)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim i As Long
      If Not IsArray(headers) Then
        HeaderArrayExists = False
        Exit Function
      End If
      For i = LBound(headers) To UBound(headers) Step 2
        If LCase(headers(i)) = LCase(key) Then
          HeaderArrayExists = True
          Exit Function
        End If
      Next
      HeaderArrayExists = False
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Sub HeaderArraySet(ByRef headers As Variant, ByVal key As String, ByVal value As String)
  With stdError.getSentry("stdHTTP#HeaderArraySet", "headers", headers, "key", key, "value", value)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim i As Long
      Dim found As Boolean
      
      If Not IsArray(headers) Then
        ReDim headers(0 To 1)
        headers(0) = key
        headers(1) = value
        Exit Sub
      End If
      
      'Check if key exists
      found = False
      For i = LBound(headers) To UBound(headers) Step 2
        If LCase(headers(i)) = LCase(key) Then
          headers(i + 1) = value
          found = True
          Exit For
        End If
      Next
      
      'If not found, add it
      If Not found Then
        Dim newSize As Long
        newSize = UBound(headers) + 1 + 2
        ReDim Preserve headers(LBound(headers) To newSize - 1)
        headers(UBound(headers) - 1) = key
        headers(UBound(headers)) = value
      End If
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

Private Function HeaderArrayGet(ByRef headers As Variant, ByVal key As String) As String
  With stdError.getSentry("stdHTTP#HeaderArrayGet", "headers", headers, "key", key)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim i As Long
      If Not IsArray(headers) Then
        HeaderArrayGet = ""
        Exit Function
      End If
      For i = LBound(headers) To UBound(headers) Step 2
        If LCase(headers(i)) = LCase(key) Then
          HeaderArrayGet = headers(i + 1)
          Exit Function
        End If
      Next
      HeaderArrayGet = ""
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Mac-specific initialization
'@protected
#If Mac Then
Public Sub protInitMac(ByVal sURL As String, ByVal RequestMethod As String, ByVal ThreadingStyle As EHTTPSynchronisity, ByVal options As Variant, ByVal Authenticator As stdICallable)
  With stdError.getSentry("stdHTTP#protInitMac", "sURL", sURL, "RequestMethod", RequestMethod, "ThreadingStyle", ThreadingStyle, "options", options, "Authenticator", Authenticator)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
        If IsEmpty(options) Or options Is Nothing Then options = CreateOptions()
        
        'Initialize response storage
        ReDim This.Response.Headers(0 To 0)
        This.Response.HeadersCount = 0
        This.Response.isComplete = False
        
        'Build curl command
        Dim curlCmd As String
        curlCmd = "curl -s -S -i" ' -s silent, -S show errors, -i include headers
        
        'Add timeout
        Dim timeoutMS As Long: timeoutMS = options(stdHTTPOptionTimeoutMS)
        If timeoutMS > 0 Then
          curlCmd = curlCmd & " --max-time " & (timeoutMS / 1000)
        End If
        
        'Add method
        curlCmd = curlCmd & " -X " & RequestMethod
        
        'Add headers (array format: key1, value1, key2, value2, ...)
        Dim headers As Variant: headers = options(stdHTTPOptionHeaders)
        Dim i As Long
        If IsArray(headers) Then
          For i = LBound(headers) To UBound(headers) Step 2
            curlCmd = curlCmd & " -H """ & headers(i) & ": " & Replace(headers(i + 1), """", "\""") & """"
          Next
        End If
        
        'Add cookies (array format: key1, value1, key2, value2, ...)
        Dim cookies As Variant: cookies = options(stdHTTPOptionCookies)
        If IsArray(cookies) Then
          For i = LBound(cookies) To UBound(cookies) Step 2
            curlCmd = curlCmd & " -H ""Cookie: " & cookies(i) & "=" & Replace(cookies(i + 1), """", "\""") & """"
          Next
        End If
        
        'Add body if present
        Dim body As String: body = options(stdHTTPOptionBody)
        If body <> "" Then
          'Use temp file for body to handle binary data and special characters reliably
          Dim tempFile As String
          tempFile = Environ("TMPDIR")
          If tempFile = "" Then tempFile = "/tmp/"
          If Right(tempFile, 1) <> "/" Then tempFile = tempFile & "/"
          tempFile = tempFile & "stdHTTP_" & Replace(Format(Now, "yyyymmddhhnnss"), " ", "") & "_" & Rnd * 10000 & ".tmp"
          Open tempFile For Binary As #1
          Put #1, , body
          Close #1
          curlCmd = curlCmd & " --data-binary @""" & tempFile & """"
          'Note: Temp file cleanup would be ideal but VBA doesn't have great cleanup mechanisms
          'The OS will clean up temp files eventually
        End If
        
        'Add proxy settings
        Dim proxy As Variant: proxy = options(stdHTTPOptionProxy)
        If proxy(stdHTTPProxyHasAutoProxy) Then
          'On Mac, curl uses system proxy by default, so we can use environment variables
          proxy(stdHTTPProxyHasAutoProxy) = false
          proxy(stdHTTPProxyHasProxy) = false
        End If
        If proxy(stdHTTPProxyHasProxy) Then
          Dim proxyUrl As String
          proxyUrl = proxy(stdHTTPProxyServer)
          If proxy(stdHTTPProxyUser) <> vbNullString And proxy(stdHTTPProxyUser) <> "" Then
            proxyUrl = proxy(stdHTTPProxyUser) & ":" & proxy(stdHTTPProxyPass) & "@" & proxyUrl
          End If
          curlCmd = curlCmd & " --proxy """ & proxyUrl & """"
        End If
        
        'Add SSL options
        If options(stdHTTPOptionInsecure) Then
          curlCmd = curlCmd & " --insecure"
        End If
        
        'Add redirect options
        If options(stdHTTPOptionEnableRedirects) Then
          curlCmd = curlCmd & " -L"
        Else
          curlCmd = curlCmd & " --max-redirs 0"
        End If
        
        'Add URL
        curlCmd = curlCmd & " """ & sURL & """"
        
        'Store the curl command
        This.curlCommand = curlCmd
        
        'Note: Authenticator callback would need to modify curlCmd before execution
        'For now, we'll call it but it won't have the same effect as on Windows
        If Not Authenticator Is Nothing Then
          'Create a mock HTTP object for authenticator compatibility
          Dim mockHTTP As Object
          Set mockHTTP = CreateObject("Scripting.Dictionary")
          'Authenticator can add headers to options("Headers") if needed
          'For now, we'll just call it for compatibility
          'Call Authenticator.Run(mockHTTP, RequestMethod, sURL, ThreadingStyle, options)
        End If
        
        'Execute curl command
        This.Created = Now()
        This.TimeoutMS = timeoutMS
        #If VBA7 Then
          This.CurlProcess = web_popen(curlCmd, "r")
        #Else
          This.CurlProcess = web_popen(curlCmd, "r")
        #End If
        
        If This.CurlProcess = 0 Then
          Err_Raise 1, "stdHTTP", "Failed to execute curl command"
        End If
        
        'If synchronous, read response immediately
        If ThreadingStyle = HTTPSync Then
          Call ReadCurlResponse
          Call Await
        End If
      
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
#End If

'Read response from curl process
'@protected
#If Mac Then
Private Sub ReadCurlResponse()
  With stdError.getSentry("stdHTTP#ReadCurlResponse")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
        If This.CurlProcess = 0 Then Exit Sub
        
        Dim buffer As String
        Dim chunkSize As Long
        Dim totalOutput As String
        Dim bytesRead As LongPtr
        
        chunkSize = 4096
        buffer = Space(chunkSize)
        totalOutput = ""
        
        'Read all output
        Do While web_feof(This.CurlProcess) = 0
          #If VBA7 Then
            bytesRead = web_fread(buffer, 1, chunkSize, This.CurlProcess)
          #Else
            bytesRead = web_fread(buffer, 1, chunkSize, This.CurlProcess)
          #End If
          If bytesRead > 0 Then
            totalOutput = totalOutput & Left(buffer, bytesRead)
          Else
            Exit Do
          End If
        Loop
        
        'Close the process
        #If VBA7 Then
          web_pclose(This.CurlProcess)
        #Else
          web_pclose(This.CurlProcess)
        #End If
        This.CurlProcess = 0
        
        'Parse response
        Call ParseCurlResponse(totalOutput)
        This.Response.isComplete = True
      
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
#End If

'Add a header to the Mac headers array
'@protected
#If Mac Then
Private Sub AddMacHeader(ByVal key As String, ByVal value As String)
  With stdError.getSentry("stdHTTP#AddMacHeader", "key", key, "value", value)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
        Dim i As Long
        Dim found As Boolean
        
        'Check if header already exists (case-insensitive)
        found = False
        For i = 0 To This.Response.HeadersCount - 1
          If LCase(This.Response.Headers(i).Key) = LCase(key) Then
            This.Response.Headers(i).Value = value
            found = True
            Exit For
          End If
        Next
        
        'If not found, add new header
        If Not found Then
          If This.Response.HeadersCount > UBound(This.Response.Headers) Then
            ReDim Preserve This.Response.Headers(0 To This.Response.HeadersCount)
          End If
          This.Response.Headers(This.Response.HeadersCount).Key = key
          This.Response.Headers(This.Response.HeadersCount).Value = value
          This.Response.HeadersCount = This.Response.HeadersCount + 1
        End If
      
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
#End If

'Get a header value from the Mac headers array
'@protected
#If Mac Then
Private Function GetMacHeader(ByVal key As String) As String
  With stdError.getSentry("stdHTTP#GetMacHeader", "key", key)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim i As Long
      For i = 0 To This.Response.HeadersCount - 1
        If LCase(This.Response.Headers(i).Key) = LCase(key) Then
          GetMacHeader = This.Response.Headers(i).Value
          Exit Function
        End If
      Next
      GetMacHeader = ""
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function
#End If

'Parse curl response output
'@protected
#If Mac Then
Private Sub ParseCurlResponse(ByVal output As String)
  With stdError.getSentry("stdHTTP#ParseCurlResponse", "output", output)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
        'Curl with -i flag outputs headers followed by blank line, then body
        Dim headerEnd As Long
        Dim headersText As String
        Dim bodyText As String
        Dim statusLine As String
        Dim lines() As String
        Dim i As Long
        
        'Find the blank line that separates headers from body
        headerEnd = InStr(1, output, vbCrLf & vbCrLf)
        If headerEnd = 0 Then headerEnd = InStr(1, output, vbLf & vbLf)
        If headerEnd = 0 Then
          'No headers, entire output is body
          bodyText = output
          headersText = ""
        Else
          headersText = Left(output, headerEnd - 1)
          bodyText = Mid(output, headerEnd + 4)
        End If
        
        'Parse status line (first line of headers)
        'Normalize line endings to vbCrLf for consistent parsing
        headersText = Replace(headersText, vbLf, vbCrLf)
        headersText = Replace(headersText, vbCrLf & vbCrLf, vbCrLf & vbCrLf) 'Ensure double CRLF
        lines = Split(headersText, vbCrLf)
        If UBound(lines) >= 0 Then
          statusLine = Trim(lines(0))
          'Parse HTTP/1.1 200 OK format
          Dim parts() As String
          parts = Split(statusLine, " ", 3)
          If UBound(parts) >= 1 Then
            On Error Resume Next
            This.Response.Status = CLng(parts(1))
            If UBound(parts) >= 2 Then
              This.Response.StatusText = parts(2)
            Else
              This.Response.StatusText = ""
            End If
            On Error GoTo stdErrorWrapper_ErrorOccurred
          End If
          
          'Parse headers (skip first line which is status)
          For i = 1 To UBound(lines)
            Dim headerLine As String
            Dim colonPos As Long
            headerLine = Trim(lines(i))
            If headerLine <> "" Then
              colonPos = InStr(1, headerLine, ":")
              If colonPos > 0 Then
                Dim headerName As String
                Dim headerValue As String
                headerName = Trim(Left(headerLine, colonPos - 1))
                headerValue = Trim(Mid(headerLine, colonPos + 1))
                Call AddMacHeader(headerName, headerValue)
              End If
            End If
          Next
        End If
        
        'Store body
        This.Response.Text = bodyText
        'Convert to byte array for ResponseBody
        Dim bodyBytes() As Byte
        bodyBytes = StrConv(bodyText, vbFromUnicode)
        This.Response.Body = bodyBytes
      
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
#End If

'Get the HSesssion
'@protected
'@returns - The HSesssion
#If Mac Then
  'Not used on Mac
#Else
  #If VBA7 Then
Friend Property Get protHSession() As LongPtr
  #Else
  Friend Property Get protHSession() As Long
  #End If
  With stdError.getSentry("stdHTTP#protHSession[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
        If Me Is stdHTTP Then
          Static hSession As LongPtr: If hSession = 0 Then hSession = WinHttpOpen(0, 1, 0, 0, 0)
          protHSession = hSession
        Else
          protHSession = stdHTTP.protHSession
        End If
      
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property
#End If

'Encode a string for use in a URL
'@param strInput - The string to encode
'@returns - The encoded string
Public Function URIEncode(ByVal strInput As String) As String
  With stdError.getSentry("stdHTTP#URIEncode", "strInput", strInput)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
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
          strOutput = strOutput & "%" & Right("0" & Hex(bytInput(i)), 2)
        End If
      Next i
    
      URIEncode = strOutput
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Decode a string from a URL
'@param strInput - The string to decode
'@returns - The decoded string
Public Function URIDecode(ByVal strInput As String) As String
  With stdError.getSentry("stdHTTP#URIDecode", "strInput", strInput)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
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
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function












'  t = t & "^(?:(http[s]?|ftp):/)?/?"  ' METHOD
'  t = t & "([^:^/^?^#\\s]+)"          ' HOSTNAME
'  t = t & "(?::(\\d+))?"              ' PORT
'  t = t & "([^?^#.*]+)?"              ' PATH
'  t = t & "(\\?[^#.]*)?"              ' QUERY
'  t = t & "(#[\\w\\-]+)?$"            ' ID

'Load the system proxy settings for a given URL
'@param sURL - The url to query
'@param proxy - The proxy array to load the settings into (passed by reference)
Private Sub LoadAutoProxy(ByVal sURL As String, ByRef proxy As Variant)
  With stdError.getSentry("stdHTTP#LoadAutoProxy", "sURL", sURL, "proxy", proxy)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim info As ProxyInfoEx: info = GetProxyCurrentUser(sURL)
      If Not info.Error.isSet Then
        'FIX: Sometimes proxy server is empty without the error being set. Ignore these proxys.
        proxy(stdHTTPProxyHasProxy) = (info.ProxyServer <> "")
        proxy(stdHTTPProxyServer) = info.ProxyServer
        proxy(stdHTTPProxyBypass) = info.ProxyBypass
        proxy(stdHTTPProxyUser) = ""  'TODO: Check if user/pass credentials with auto-proxy will ever be needed
        proxy(stdHTTPProxyPass) = ""
      Else
        If info.Error.isRecoverable Then
          'Assume no proxy
          proxy(stdHTTPProxyHasProxy) = False
        Else
          Err_Raise info.Error.Num, "stdHTTP", info.Error.Msg
        End If
      End If
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Get the system proxy settings for a given URL
'@param sURL - The url to query
'@returns - The proxy information struct
Private Function GetProxyCurrentUser(ByVal sURL As String) As ProxyInfoEx
  With stdError.getSentry("stdHTTP#GetProxyCurrentUser", "sURL", sURL)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
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
            Case ERROR_WINHTTP_INVALID_PARAMETER, ERROR_WINHTTP_AUTO_PROXY_SERVICE_ERROR, ERROR_WINHTTP_AUTODETECTION_FAILED, ERROR_WINHTTP_BAD_AUTO_PROXY_SCRIPT, _
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
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Copies a variant from the source variable to the dest variable
'@param dest - Destination to copy variant to
'@param src - Source to copy variant from
Private Sub CopyVariant(ByRef dest As Variant, ByVal src As Variant)
  With stdError.getSentry("stdHTTP#CopyVariant", "dest", dest, "src", src)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If IsObject(src) Then
        Set dest = src
      Else
        Let dest = src
      End If
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub


Private Sub Err_Raise(ByVal number as Long, Optional ByVal source as string = "", Optional ByVal description as string = "")
  Call stdError.Raise(description)
End Sub
