'WinHttp.WinHttpRequest.5.1
'Microsoft.XMLHTTP
'WinHTTP

Implements stdICallable

Private Declare Function WinHttpCrackUrl Lib "winhttp" (ByVal pwszUrl As Long, ByVal dwUrlLength As Long, ByVal dwFlags As Long, lpUrlComponents As Any) As Long
Private Declare Function WinHttpCreateUrl Lib "winhttp" (lpUrlComponents As Any, ByVal dwFlags As Long, ByVal pwszUrl As Long, pdwUrlLength As Long) As Long
 
Private Type URL_COMPONENTS
    dwStructSize        As Long
    lpszScheme          As Long
    dwSchemeLength      As Long
    nScheme             As Long
    lpszHostName        As Long
    dwHostNameLength    As Long
    nPort               As Long
    lpszUserName        As Long
    dwUserNameLength    As Long
    lpszPassword        As Long
    dwPasswordLength    As Long
    lpszUrlPath         As Long
    dwUrlPathLength     As Long
    lpszExtraInfo       As Long
    dwExtraInfoLength   As Long
End Type

Private tURL as URL_COMPONENTS
Private sURL as string

'Basic Authentication
Private pUser as string
Private pPass as string

 
'Public Function SetCredentials(sUrl As String, ByVal sUser As String, ByVal sPass As String) As String
'    Dim uUrlCom         As URL_COMPONENTS
'    Dim sRet            As String
'    Dim lSize           As Long
' 
'    With uUrlCom
'        .dwStructSize = Len(uUrlCom)
'        .dwSchemeLength = -1
'        .dwHostNameLength = -1
'        .dwUrlPathLength = -1
'        .dwExtraInfoLength = -1
'    End With
'    Call WinHttpCrackUrl(StrPtr(sUrl), Len(sUrl), 0, uUrlCom)
'    sUser = pvEscapeCredential(sUser)
'    sPass = pvEscapeCredential(sPass)
'    With uUrlCom
'        .lpszUserName = StrPtr(sUser)
'        .dwUserNameLength = Len(sUser)
'        .lpszPassword = StrPtr(sPass)
'        .dwPasswordLength = Len(sPass)
'    End With
'    sRet = String(4096, 0)
'    lSize = Len(sRet)
'    Call WinHttpCreateUrl(uUrlCom, 0, StrPtr(sRet), lSize)
'    SetCredentials = Left$(sRet, InStr(sRet, vbNullChar))
'End Function
 



Public Function Create(ByVal sMethod as string, ByVal sURL as string, ) as stdHTTP

End Function
Public Function Run(ByVal data as string)

End Function
Public Sub RunAsync(ByVal data as string, cb as stdICallable)

End Sub

Public Property Get isAutoProxyEnabled() as boolean

End Property
Public Property Let isAutoProxyEnabled(v as boolean)

End Property

Public Function BindAuth(options as object) as boolean
    Select case options("Type") 
        case "BASIC"
            pUser = options("User")
            pPass = options("Pass")
            With tURL
                .lpszUserName = StrPtr(pUser)
                .dwUserNameLength = Len(pUser)
                .lpszPassword = StrPtr(pPass)
                .dwPasswordLength = Len(pPass)
            End With
    End Select
End Function

'Gets a wrapper which applies basic authentication to the request
Public Function GetBasicAuth(ByVal sUser as string, ByVal sPass as string) as object
    set GetBasicAuth = CreateObject("Scripting.Dictionary")
    GetBasicAuth.item("Type") = "BASIC"
    GetBasicAuth.item("User") = sUser
    GetBasicAuth.item("Pass") = sPass
End Function

'Gets a wrapper which applies OAuth authentication to the request
'might be able to use special redirect: urn:ietf:wg:oauth:2.0:oob   as shown here: https://stackoverflow.com/questions/26428043/exchanging-authorization-code-for-access-token-for-google-calendar-api-with-vba
'https://help.salesforce.com/articleView?id=remoteaccess_oauth_user_agent_flow.htm&type=5
Public Function GetOAuth2Binding(ByVal sClientID as string,ByVal sClientSecret as string,ByVal sTokenHost as string) as object
    set GetBasicAuth = CreateObject("Scripting.Dictionary")
    GetBasicAuth.item("Type") = "OAUTH2"
    GetBasicAuth.item("ClientID") = sClientID
    GetBasicAuth.item("ClientSecret") = sClientSecret
    '? GetBasicAuth.item("User") = sUser
    '? GetBasicAuth.item("Pass") = sPass

    'step 1: host a server
End Function



















Private Function pvEscape(sText As String) As String
    pvEscape = Replace(Replace(Replace(Replace(sText, _
        "%", "%" & Hex(Asc("%"))), _
        "/", "%" & Hex(Asc("/"))), _
        "@", "%" & Hex(Asc("@"))), _
        "?", "%" & Hex(Asc("?")))
End Function