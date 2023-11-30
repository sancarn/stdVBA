Attribute VB_Name = "stdHTTPAuthenticators"
'@lang VBA

Type DigestAuthInfo
  Realm As String
  nonce As String
  opaque As String
End Type

'Authenticator will logon with Windows login credentials if requested
'@param pHTTP - The HTTP object from the stdHTTP.Create() call
'@param RequestMethod - The HTTP method from the stdHTTP.Create() call
'@param sURL - The URL from the stdHTTP.Create() call
'@param ThreadingStyle - The threading style from the stdHTTP.Create() call
'@param options - The options object from the stdHTTP.Create() call
'@example `stdHTTP.Create("someURL", Authenticator:=stdCallback.CreateFromModule("stdHTTPAuthenticators", "WindowsAuthenticator"))`
Public Sub WindowsAuthenticator(ByVal pHTTP As Object, ByVal RequestMethod As String, ByVal sURL As String, ByVal ThreadingStyle As Long, ByVal options As Object)
  Const AutoLogonPolicy_Always = 0
  Const AutoLogonPolicy_OnlyIfBypassProxy = 1
  Const AutoLogonPolicy_Never = 2
  Call pHTTP.SetAutoLogonPolicy(AutoLogonPolicy_Always)
End Sub

'Basic Authenticator. 
'@param Username - The username supplied by the user during Bind()
'@param Password - The password supplied by the user during Bind()
'@param pHTTP - The HTTP object from the stdHTTP.Create() call
'@param RequestMethod - The HTTP method from the stdHTTP.Create() call
'@param sURL - The URL from the stdHTTP.Create() call
'@param ThreadingStyle - The threading style from the stdHTTP.Create() call
'@param options - The options object from the stdHTTP.Create() call
'@example `stdHTTP.Create("https://postman-echo.com/basic-auth", Authenticator:=stdCallback.CreateFromModule("stdHTTPAuthenticators", "HttpBasicAuthenticator").Bind("postman", "password"))`
'@remark This authenticator will send the username and password in the clear. It is recommended to use this only over HTTPS.
Public Sub HttpBasicAuthenticator(ByVal Username As String, ByVal Password As String, ByVal pHTTP As Object, ByVal RequestMethod As String, ByVal sURL As String, ByVal ThreadingStyle As Long, ByVal options As Object)
  Const SetCredentialsType_ForServer = 0
  pHTTP.SetCredentials Username, Password, SetCredentialsType_ForServer
End Sub

'Token Authenticator
'@param HeaderName - The name of the header to set supplied by the user during Bind()
'@param Token - The token to set the header to supplied by the user during Bind()
'@param pHTTP - The HTTP object from the stdHTTP.Create() call
'@param RequestMethod - The HTTP method from the stdHTTP.Create() call
'@param sURL - The URL from the stdHTTP.Create() call
'@param ThreadingStyle - The threading style from the stdHTTP.Create() call
'@param options - The options object from the stdHTTP.Create() call
'@example `stdHTTP.Create("https://postman-echo.com/basic-auth", Authenticator:=stdCallback.CreateFromModule("stdHTTPAuthenticators", "TokenAuthenticator").Bind("PRIVATE-TOKEN", "{{your-token}}"))`
Public Sub TokenAuthenticator(ByVal HeaderName As String, ByVal Token As String, ByVal pHTTP As Object, ByVal RequestMethod As String, ByVal sURL As String, ByVal ThreadingStyle As Long, ByVal options As Object)
  Call pHTTP.SetHeader(HeaderName, Token)
End Sub


'Digest Authenticator
'@param Username - The username supplied by the user during Bind()
'@param Password - The password supplied by the user during Bind()
'@param pHTTP - The HTTP object from the stdHTTP.Create() call
'@param RequestMethod - The HTTP method from the stdHTTP.Create() call
'@param sURL - The URL from the stdHTTP.Create() call
'@param ThreadingStyle - The threading style from the stdHTTP.Create() call
'@param options - The options object from the stdHTTP.Create() call
'@example `stdHTTP.Create("https://postman-echo.com/digest-auth", Authenticator:=stdCallback.CreateFromModule("stdHTTPAuthenticators", "DigestAuthenticator").Bind("postman", "password", "postman-echo.com"))`
'@TODO: Complete this
Public Sub DigestAuthenticator(ByVal Username As String, ByVal Password As String, ByVal pHTTP As Object, ByVal sDomain As String, ByVal RequestMethod As String, ByVal sURL As String, ByVal ThreadingStyle As Long, ByVal options As Object)
  Err.Raise 1, "", "Work in progress - This does not work yet"
  Static cache As Object: If cache Is Nothing Then Set cache = CreateObject("Scripting.Dictionary")
  If Not cache.exists(sDomain) Then
    'Clone request
    Dim rInitial As stdHTTP: Set rInitial = stdHTTP.Create(sURL, RequestMethod, ThreadingStyle, options)
    If rInitial.ResponseStatus >= 400 Then
      'cache(sDomain) = getDigestHeader(...)
    Else
      'cache(sDomain) = ...
    End If
  End If
  
  pHTTP.SetHeader "Authorization", cache(sDomain)
End Sub
