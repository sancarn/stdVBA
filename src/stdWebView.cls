VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdWebView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@class stdWebView
'@description Standalone WebView2 host for Excel UserForms (primary). Requires WebView2Loader.dll,
' Microsoft Edge WebView2 Runtime. Handler vtables use executable thunks (see AddressOfThunkForComProc/VBA7/CMonitors.cls).
'@attribution https://github.com/tarboh/WebView2-For-Excel-VBA - Heavily inspired by this project.
'@example:
'   Dim wv As stdWebView
'   Sub UserForm_Initialize()
'     Set wv = stdWebView.CreateFromFrame(UserForm1.Frame1, stdLambda.Create("$1.Navigate ""https://example.com"""))
'     wv.Navigate "https://example.com"
'     Debug.Print wv.JavaScriptRunSync("document.title")
'   End Sub
'
'Spec:
' CONSTRUCTORS
'   [X] CreateFromHwnd(hwnd, OnReady?) as stdWebView
'   [X] CreateFromFrame(frm as MSForms.Frame, OnReady?) as stdWebView
'   [X] CreateFromUserform(uf as Object, OnReady?) as stdWebView
' INSTANCE METHODS / PROPERTIES
'   [X] IsReady() as Boolean
'   [X] Quit()
'   [X] Navigate(url)
'   [X] Back()
'   [X] Forward()
'   [X] Get/Let Html() as String
'   [X] JavaScriptRunSync(script) as String
'   [X] JavaScriptRun(script, callback?)
'   [X] SetVirtualHostNameToFolderMapping(hostName, folderPath, accessKind?)
'   [X] ClearVirtualHostNameToFolderMapping(hostName)

Option Explicit

#If Win64 Then
    Private Const PTR_SZ As Long = 8
#Else
    Private Const PTR_SZ As Long = 4
#End If

#If VBA7 = 0 Then
    Private Enum LongPtr
        [_]
    End Enum
#End If

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type THandlerThunks
    pQI As LongPtr
    pAddRef As LongPtr
    pRelease As LongPtr
    pInvoke As LongPtr
End Type

Private Type TScriptItem
    requestId As Long
    isSyncWait As Boolean
    isCompleted As Boolean
    callback As stdICallable
    pHandlerObj As LongPtr
    pVTable As LongPtr
    thunks As THandlerThunks
End Type

Private Type TInitState
    userDataFolder As String
    waiting As Boolean
    failed As Boolean
    onReadyCallback As stdICallable
End Type

Private Type TScriptState
    result As String
    syncDone As Boolean
    syncWaitRequestId As Long
    syncBridgeRequestId As String
    syncBridgeDone As Boolean
    syncBridgeOk As Boolean
    syncBridgePayload As String
    nextRequestId As Long
End Type

Private Type TScripts
    mScripts() As TScriptItem
    scriptCount As Long
End Type

Private Type THostObjectState
    refs As Collection
End Type

Private Type TWebResourceFilterItem
    filterString As String
    context As Long
    callback As stdICallable
End Type

Private Type TWebResourceFilters
    items() As TWebResourceFilterItem
    count As Long
End Type

Private Type TDeferredWebResourceItem
    requestId As Long
    pArgs As LongPtr
    pDeferral As LongPtr
    createdAt As Double
End Type

Private Type TDeferredWebResources
    items() As TDeferredWebResourceItem
    count As Long
    nextRequestId As Long
End Type

Private Type TComState
    pEnvironment As LongPtr
    pController As LongPtr
    ppWebView2 As LongPtr
    pEnvHandlerObj As LongPtr
    pEnvVTable As LongPtr
    envThunk As THandlerThunks
    pCtrlHandlerObj As LongPtr
    pCtrlVTable As LongPtr
    ctrlThunk As THandlerThunks
    pWebReqHandlerObj As LongPtr
    pWebReqVTable As LongPtr
    webReqThunk As THandlerThunks
    webReqTokenLow As Long
    webReqTokenHigh As Long
    pWebMsgHandlerObj As LongPtr
    pWebMsgVTable As LongPtr
    webMsgThunk As THandlerThunks
    webMsgTokenLow As Long
    webMsgTokenHigh As Long
End Type

Private Type TThis
    parentHwnd As LongPtr
    lastHtml As String
    init As TInitState
    script As TScriptState
    scripts As TScripts
    host As THostObjectState
    webResource As TWebResourceFilters
    deferredWebResource As TDeferredWebResources
    com As TComState
End Type

Private This As TThis

Private Const CC_STDCALL As Long = 4
Private Const GPTR As Long = &H40
Private Const GMEM_MOVEABLE As Long = &H2&
Private Const GMEM_ZEROINIT As Long = &H40&

Private Const MEM_COMMIT As Long = &H1000&
Private Const MEM_RESERVE As Long = &H2000&
Private Const MEM_RELEASE As Long = &H8000&
Private Const PAGE_EXECUTE_READWRITE As Long = &H40&
Private Const GW_CHILD As Long = 5
Private Const GW_HWNDNEXT As Long = 2
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90
' Method slots are from ICoreWebView2_19Vtbl in WebView2.h (IUnknown + interface members).
Private Const COREWEBVIEW2_VTBL_GET_SETTINGS As Long = 3
Private Const COREWEBVIEW2_VTBL_GET_SOURCE As Long = 4
Private Const COREWEBVIEW2_VTBL_NAVIGATE As Long = 5
Private Const COREWEBVIEW2_VTBL_NAVIGATETOSTRING As Long = 6
Private Const COREWEBVIEW2_VTBL_ADD_NAVIGATIONSTARTING As Long = 7
Private Const COREWEBVIEW2_VTBL_REMOVE_NAVIGATIONSTARTING As Long = 8
Private Const COREWEBVIEW2_VTBL_ADD_CONTENTLOADING As Long = 9
Private Const COREWEBVIEW2_VTBL_REMOVE_CONTENTLOADING As Long = 10
Private Const COREWEBVIEW2_VTBL_ADD_SOURCECHANGED As Long = 11
Private Const COREWEBVIEW2_VTBL_REMOVE_SOURCECHANGED As Long = 12
Private Const COREWEBVIEW2_VTBL_ADD_HISTORYCHANGED As Long = 13
Private Const COREWEBVIEW2_VTBL_REMOVE_HISTORYCHANGED As Long = 14
Private Const COREWEBVIEW2_VTBL_ADD_NAVIGATIONCOMPLETED As Long = 15
Private Const COREWEBVIEW2_VTBL_REMOVE_NAVIGATIONCOMPLETED As Long = 16
Private Const COREWEBVIEW2_VTBL_ADD_FRAMENAVIGATIONSTARTING As Long = 17
Private Const COREWEBVIEW2_VTBL_REMOVE_FRAMENAVIGATIONSTARTING As Long = 18
Private Const COREWEBVIEW2_VTBL_ADD_FRAMENAVIGATIONCOMPLETED As Long = 19
Private Const COREWEBVIEW2_VTBL_REMOVE_FRAMENAVIGATIONCOMPLETED As Long = 20
Private Const COREWEBVIEW2_VTBL_ADD_SCRIPTDIALOGOPENING As Long = 21
Private Const COREWEBVIEW2_VTBL_REMOVE_SCRIPTDIALOGOPENING As Long = 22
Private Const COREWEBVIEW2_VTBL_ADD_PERMISSIONREQUESTED As Long = 23
Private Const COREWEBVIEW2_VTBL_REMOVE_PERMISSIONREQUESTED As Long = 24
Private Const COREWEBVIEW2_VTBL_ADD_PROCESSFAILED As Long = 25
Private Const COREWEBVIEW2_VTBL_REMOVE_PROCESSFAILED As Long = 26
Private Const COREWEBVIEW2_VTBL_ADDSCRIPTTOEXECUTEONDOCUMENTCREATED As Long = 27
Private Const COREWEBVIEW2_VTBL_REMOVESCRIPTTOEXECUTEONDOCUMENTCREATED As Long = 28
Private Const COREWEBVIEW2_VTBL_EXECUTESCRIPT As Long = 29
Private Const COREWEBVIEW2_VTBL_CAPTUREPREVIEW As Long = 30
Private Const COREWEBVIEW2_VTBL_RELOAD As Long = 31
Private Const COREWEBVIEW2_VTBL_POSTWEBMESSAGEASJSON As Long = 32
Private Const COREWEBVIEW2_VTBL_POSTWEBMESSAGEASSTRING As Long = 33
Private Const COREWEBVIEW2_VTBL_ADD_WEBMESSAGERECEIVED As Long = 34
Private Const COREWEBVIEW2_VTBL_REMOVE_WEBMESSAGERECEIVED As Long = 35
Private Const COREWEBVIEW2_VTBL_CALLDEVTOOLSPROTOCOLMETHOD As Long = 36
Private Const COREWEBVIEW2_VTBL_GET_BROWSERPROCESSID As Long = 37
Private Const COREWEBVIEW2_VTBL_GET_CANGOBACK As Long = 38
Private Const COREWEBVIEW2_VTBL_GET_CANGOFORWARD As Long = 39
Private Const COREWEBVIEW2_VTBL_GOBACK As Long = 40
Private Const COREWEBVIEW2_VTBL_GOFORWARD As Long = 41
Private Const COREWEBVIEW2_VTBL_GETDEVTOOLSPROTOCOLEVENTRECEIVER As Long = 42
Private Const COREWEBVIEW2_VTBL_STOP As Long = 43
Private Const COREWEBVIEW2_VTBL_ADD_NEWWINDOWREQUESTED As Long = 44
Private Const COREWEBVIEW2_VTBL_REMOVE_NEWWINDOWREQUESTED As Long = 45
Private Const COREWEBVIEW2_VTBL_ADD_DOCUMENTTITLECHANGED As Long = 46
Private Const COREWEBVIEW2_VTBL_REMOVE_DOCUMENTTITLECHANGED As Long = 47
Private Const COREWEBVIEW2_VTBL_GET_DOCUMENTTITLE As Long = 48
Private Const COREWEBVIEW2_VTBL_ADDHOSTOBJECTTOSCRIPT As Long = 49
Private Const COREWEBVIEW2_VTBL_REMOVEHOSTOBJECTFROMSCRIPT As Long = 50
Private Const COREWEBVIEW2_VTBL_OPENDEVTOOLSWINDOW As Long = 51
Private Const COREWEBVIEW2_VTBL_ADD_CONTAINSFULLSCREENELEMENTCHANGED As Long = 52
Private Const COREWEBVIEW2_VTBL_REMOVE_CONTAINSFULLSCREENELEMENTCHANGED As Long = 53
Private Const COREWEBVIEW2_VTBL_GET_CONTAINSFULLSCREENELEMENT As Long = 54
Private Const COREWEBVIEW2_VTBL_ADD_WEBRESOURCEREQUESTED As Long = 55
Private Const COREWEBVIEW2_VTBL_REMOVE_WEBRESOURCEREQUESTED As Long = 56
Private Const COREWEBVIEW2_VTBL_ADDWEBRESOURCEREQUESTEDFILTER As Long = 57
Private Const COREWEBVIEW2_VTBL_REMOVEWEBRESOURCEREQUESTEDFILTER As Long = 58
Private Const COREWEBVIEW2_VTBL_ADD_WINDOWCLOSEREQUESTED As Long = 59
Private Const COREWEBVIEW2_VTBL_REMOVE_WINDOWCLOSEREQUESTED As Long = 60
Private Const COREWEBVIEW2_VTBL_ADD_WEBRESOURCERESPONSERECEIVED As Long = 61
Private Const COREWEBVIEW2_VTBL_REMOVE_WEBRESOURCERESPONSERECEIVED As Long = 62
Private Const COREWEBVIEW2_VTBL_NAVIGATEWITHWEBRESOURCEREQUEST As Long = 63
Private Const COREWEBVIEW2_VTBL_ADD_DOMCONTENTLOADED As Long = 64
Private Const COREWEBVIEW2_VTBL_REMOVE_DOMCONTENTLOADED As Long = 65
Private Const COREWEBVIEW2_VTBL_GET_COOKIEMANAGER As Long = 66
Private Const COREWEBVIEW2_VTBL_GET_ENVIRONMENT As Long = 67
Private Const COREWEBVIEW2_VTBL_TRYSUSPEND As Long = 68
Private Const COREWEBVIEW2_VTBL_RESUME As Long = 69
Private Const COREWEBVIEW2_VTBL_GET_ISSUSPENDED As Long = 70
Private Const COREWEBVIEW2_VTBL_SETVIRTUALHOSTNAMETOFOLDERMAPPING As Long = 71
Private Const COREWEBVIEW2_VTBL_CLEARVIRTUALHOSTNAMETOFOLDERMAPPING As Long = 72
Private Const COREWEBVIEW2_VTBL_ADD_FRAMECREATED As Long = 73
Private Const COREWEBVIEW2_VTBL_REMOVE_FRAMECREATED As Long = 74
Private Const COREWEBVIEW2_VTBL_ADD_DOWNLOADSTARTING As Long = 75
Private Const COREWEBVIEW2_VTBL_REMOVE_DOWNLOADSTARTING As Long = 76
Private Const COREWEBVIEW2_VTBL_ADD_CLIENTCERTIFICATEREQUESTED As Long = 77
Private Const COREWEBVIEW2_VTBL_REMOVE_CLIENTCERTIFICATEREQUESTED As Long = 78
Private Const COREWEBVIEW2_VTBL_OPENTASKMANAGERWINDOW As Long = 79
Private Const COREWEBVIEW2_VTBL_PRINTTOPDF As Long = 80
Private Const COREWEBVIEW2_VTBL_ADD_ISMUTEDCHANGED As Long = 81
Private Const COREWEBVIEW2_VTBL_REMOVE_ISMUTEDCHANGED As Long = 82
Private Const COREWEBVIEW2_VTBL_GET_ISMUTED As Long = 83
Private Const COREWEBVIEW2_VTBL_PUT_ISMUTED As Long = 84
Private Const COREWEBVIEW2_VTBL_ADD_ISDOCUMENTPLAYINGAUDIOCHANGED As Long = 85
Private Const COREWEBVIEW2_VTBL_REMOVE_ISDOCUMENTPLAYINGAUDIOCHANGED As Long = 86
Private Const COREWEBVIEW2_VTBL_GET_ISDOCUMENTPLAYINGAUDIO As Long = 87
Private Const COREWEBVIEW2_VTBL_ADD_ISDEFAULTDOWNLOADDIALOGOPENCHANGED As Long = 88
Private Const COREWEBVIEW2_VTBL_REMOVE_ISDEFAULTDOWNLOADDIALOGOPENCHANGED As Long = 89
Private Const COREWEBVIEW2_VTBL_GET_ISDEFAULTDOWNLOADDIALOGOPEN As Long = 90
Private Const COREWEBVIEW2_VTBL_OPENDEFAULTDOWNLOADDIALOG As Long = 91
Private Const COREWEBVIEW2_VTBL_CLOSEDEFAULTDOWNLOADDIALOG As Long = 92
Private Const COREWEBVIEW2_VTBL_GET_DEFAULTDOWNLOADDIALOGCORNERALIGNMENT As Long = 93
Private Const COREWEBVIEW2_VTBL_PUT_DEFAULTDOWNLOADDIALOGCORNERALIGNMENT As Long = 94
Private Const COREWEBVIEW2_VTBL_GET_DEFAULTDOWNLOADDIALOGMARGIN As Long = 95
Private Const COREWEBVIEW2_VTBL_PUT_DEFAULTDOWNLOADDIALOGMARGIN As Long = 96
Private Const COREWEBVIEW2_VTBL_ADD_BASICAUTHENTICATIONREQUESTED As Long = 97
Private Const COREWEBVIEW2_VTBL_REMOVE_BASICAUTHENTICATIONREQUESTED As Long = 98
Private Const COREWEBVIEW2_VTBL_CALLDEVTOOLSPROTOCOLMETHODFORSESSION As Long = 99
Private Const COREWEBVIEW2_VTBL_ADD_CONTEXTMENUREQUESTED As Long = 100
Private Const COREWEBVIEW2_VTBL_REMOVE_CONTEXTMENUREQUESTED As Long = 101
Private Const COREWEBVIEW2_VTBL_ADD_STATUSBARTEXTCHANGED As Long = 102
Private Const COREWEBVIEW2_VTBL_REMOVE_STATUSBARTEXTCHANGED As Long = 103
Private Const COREWEBVIEW2_VTBL_GET_STATUSBARTEXT As Long = 104
Private Const COREWEBVIEW2_VTBL_GET_PROFILE As Long = 105
Private Const COREWEBVIEW2_VTBL_ADD_SERVERCERTIFICATEERRORDETECTED As Long = 106
Private Const COREWEBVIEW2_VTBL_REMOVE_SERVERCERTIFICATEERRORDETECTED As Long = 107
Private Const COREWEBVIEW2_VTBL_CLEARSERVERCERTIFICATEERRORACTIONS As Long = 108
Private Const COREWEBVIEW2_VTBL_ADD_FAVICONCHANGED As Long = 109
Private Const COREWEBVIEW2_VTBL_REMOVE_FAVICONCHANGED As Long = 110
Private Const COREWEBVIEW2_VTBL_GET_FAVICONURI As Long = 111
Private Const COREWEBVIEW2_VTBL_GETFAVICON As Long = 112
Private Const COREWEBVIEW2_VTBL_PRINT As Long = 113
Private Const COREWEBVIEW2_VTBL_SHOWPRINTUI As Long = 114
Private Const COREWEBVIEW2_VTBL_PRINTTOPDFSTREAM As Long = 115
Private Const COREWEBVIEW2_VTBL_POSTSHAREDBUFFERTOSCRIPT As Long = 116
Private Const COREWEBVIEW2_VTBL_ADD_LAUNCHINGEXTERNALURISCHEME As Long = 117
Private Const COREWEBVIEW2_VTBL_REMOVE_LAUNCHINGEXTERNALURISCHEME As Long = 118
Private Const COREWEBVIEW2_VTBL_GET_MEMORYUSAGETARGETLEVEL As Long = 119
Private Const COREWEBVIEW2_VTBL_PUT_MEMORYUSAGETARGETLEVEL As Long = 120

Private Const COREWEBVIEW2ENVIRONMENT_VTBL_CREATEWEBRESOURCERESPONSE As Long = 4

Private Const COREWEBVIEW2WEBRESOURCEREQUEST_VTBL_GET_URI As Long = 3
Private Const COREWEBVIEW2WEBRESOURCEREQUEST_VTBL_GET_METHOD As Long = 5
Private Const COREWEBVIEW2WEBRESOURCEREQUEST_VTBL_GET_HEADERS As Long = 9

Private Const COREWEBVIEW2HTTPREQUESTHEADERS_VTBL_SETHEADER As Long = 6
Private Const COREWEBVIEW2HTTPREQUESTHEADERS_VTBL_REMOVEHEADER As Long = 7
Private Const COREWEBVIEW2HTTPREQUESTHEADERS_VTBL_GETITERATOR As Long = 8

Private Const COREWEBVIEW2HTTPHEADERSITERATOR_VTBL_GETCURRENTHEADER As Long = 3
Private Const COREWEBVIEW2HTTPHEADERSITERATOR_VTBL_GETHASCURRENTHEADER As Long = 4
Private Const COREWEBVIEW2HTTPHEADERSITERATOR_VTBL_MOVENEXT As Long = 5

Private Const COREWEBVIEW2WEBRESOURCEREQUESTEDEVENTARGS_VTBL_GET_REQUEST As Long = 3
Private Const COREWEBVIEW2WEBRESOURCEREQUESTEDEVENTARGS_VTBL_PUT_RESPONSE As Long = 5
Private Const COREWEBVIEW2WEBRESOURCEREQUESTEDEVENTARGS_VTBL_GETDEFERRAL As Long = 6
Private Const COREWEBVIEW2WEBRESOURCEREQUESTEDEVENTARGS_VTBL_GET_RESOURCECONTEXT As Long = 7
Private Const COREWEBVIEW2DEFERRAL_VTBL_COMPLETE As Long = 3
Private Const COREWEBVIEW2WEBMESSAGERECEIVEDEVENTARGS_VTBL_GET_WEBMESSAGEASJSON As Long = 4

Private Const SCRIPT_SYNC_BRIDGE_PREFIX As String = "__stdwv_sync__"
Private Const SCRIPT_SYNC_TIMEOUT_MS As Long = 120000

Public Enum COREWEBVIEW2_HOST_RESOURCE_ACCESS_KIND
    COREWEBVIEW2_HOST_RESOURCE_ACCESS_KIND_DENY = 0
    COREWEBVIEW2_HOST_RESOURCE_ACCESS_KIND_ALLOW = 1
    COREWEBVIEW2_HOST_RESOURCE_ACCESS_KIND_DENY_CORS = 2
End Enum

Public Enum ECoreWebView2WebResourceContext
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_ALL = 0
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_DOCUMENT = 1
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_STYLESHEET = 2
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_IMAGE = 3
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_MEDIA = 4
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_FONT = 5
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_SCRIPT = 6
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_XML_HTTP_REQUEST = 7
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_FETCH = 8
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_TEXT_TRACK = 9
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_EVENT_SOURCE = 10
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_WEBSOCKET = 11
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_MANIFEST = 12
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_SIGNED_EXCHANGE = 13
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_PING = 14
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_CSP_VIOLATION_REPORT = 15
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_OTHER = 16
End Enum

#If VBA7 Then

    Private Declare PtrSafe Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader.dll" ( _
        ByVal browserExecutableFolder As LongPtr, _
        ByVal userDataFolder As LongPtr, _
        ByVal additionalBrowserArguments As LongPtr, _
        ByVal environmentCreatedHandler As LongPtr) As Long

    Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" ( _
        ByVal lpAddress As LongPtr, _
        ByVal dwSize As LongPtr, _
        ByVal flAllocationType As Long, _
        ByVal flProtect As Long) As LongPtr

    Private Declare PtrSafe Function VirtualFree Lib "kernel32" ( _
        ByVal lpAddress As LongPtr, _
        ByVal dwSize As LongPtr, _
        ByVal dwFreeType As Long) As Long

    Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" ( _
        ByVal lpModuleName As String) As LongPtr

    Private Declare PtrSafe Function GetProcAddress Lib "kernel32" ( _
        ByVal hModule As LongPtr, _
        ByVal lpProcName As String) As LongPtr

    Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long

    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

    Private Declare PtrSafe Sub CopyMemoryByPtr Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByVal Destination As LongPtr, _
        ByRef Source As Any, _
        ByVal Length As LongPtr)

    Private Declare PtrSafe Sub CopyMemoryByAddr Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByRef Destination As Any, _
        ByVal Source As LongPtr, _
        ByVal Length As LongPtr)
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As LongPtr)
    Private Declare PtrSafe Function CreateStreamOnHGlobal Lib "ole32.dll" (ByVal hGlobal As LongPtr, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any) As Long

    Private Declare PtrSafe Function DispCallFunc Lib "oleaut32.dll" ( _
        ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal cc As Long, _
        ByVal vtReturn As Integer, ByVal cArgs As Long, ByVal rgvt As LongPtr, _
        ByVal rgpvarg As LongPtr, ByRef pvargResult As Variant) As Long

    Private Declare PtrSafe Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pUnk As IUnknown, ByVal pHwnd As LongPtr) As Long

    Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hwnd As LongPtr, ByRef lpRect As RECT) As Long
    Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
        ByVal hWndParent As LongPtr, ByVal hWndChildAfter As LongPtr, ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal wCmd As Long) As LongPtr
    Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, ByRef lpRect As RECT) As Long
    Private Declare PtrSafe Function ClientToScreen Lib "user32" (ByVal hwnd As LongPtr, ByRef lpPoint As POINTAPI) As Long
#Else
    Private Declare Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader.dll" ( _
        ByVal browserExecutableFolder As LongPtr, _
        ByVal userDataFolder As LongPtr, _
        ByVal additionalBrowserArguments As LongPtr, _
        ByVal environmentCreatedHandler As LongPtr) As Long

    Private Declare Function VirtualAlloc Lib "kernel32" ( _
        ByVal lpAddress As LongPtr, _
        ByVal dwSize As LongPtr, _
        ByVal flAllocationType As Long, _
        ByVal flProtect As Long) As LongPtr

    Private Declare Function VirtualFree Lib "kernel32" ( _
        ByVal lpAddress As LongPtr, _
        ByVal dwSize As LongPtr, _
        ByVal dwFreeType As Long) As Long

    Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" ( _
        ByVal lpModuleName As String) As LongPtr

    Private Declare Function GetProcAddress Lib "kernel32" ( _
        ByVal hModule As LongPtr, _
        ByVal lpProcName As String) As LongPtr

    Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long

    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

    Private Declare Sub CopyMemoryByPtr Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByVal Destination As LongPtr, _
        ByRef Source As Any, _
        ByVal Length As LongPtr)

    Private Declare Sub CopyMemoryByAddr Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByRef Destination As Any, _
        ByVal Source As LongPtr, _
        ByVal Length As LongPtr)
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As LongPtr)
    Private Declare Function CreateStreamOnHGlobal Lib "ole32.dll" (ByVal hGlobal As LongPtr, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any) As Long

    Private Declare Function DispCallFunc Lib "oleaut32.dll" ( _
        ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal cc As Long, _
        ByVal vtReturn As Integer, ByVal cArgs As Long, ByVal rgvt As LongPtr, _
        ByVal rgpvarg As LongPtr, ByRef pvargResult As Variant) As Long

    Private Declare Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pUnk As IUnknown, ByVal phwnd As LongPtr) As Long

    Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As LongPtr, ByRef lpRect As RECT) As Long
    Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
        ByVal hWndParent As LongPtr, ByVal hWndChildAfter As LongPtr, ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
    Private Declare Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal wCmd As Long) As LongPtr
    Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, ByRef lpRect As RECT) As Long
    Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As LongPtr, ByRef lpPoint As POINTAPI) As Long
#End If

'Call a vtable method (same idea as stdCOM.CallVT). VBA forbids Optional before ParamArray, so use CallVTQuiet when DispCallFunc errors must not raise.
Private Function CallVTImpl(ByVal pInterface As LongPtr, ByVal VTableOffset As Long, ByVal ReturnType As VbVarType, ByVal raiseOnDispFailure As Boolean, ByVal vArgs As Variant) As Variant
    Dim iParamCount As Long
    Dim lb As Long
    Dim lIdx As Long
    iParamCount = 0
    lb = 0
    On Error Resume Next
    lb = LBound(vArgs)
    iParamCount = UBound(vArgs) - lb + 1
    If Err.Number <> 0 Or iParamCount < 0 Then
        iParamCount = 0
        lb = 0
    End If
    Err.Clear
    On Error GoTo 0

    Dim vPtr() As LongPtr
    Dim vType() As Integer

    If iParamCount = 0 Then
        ReDim vType(0 To 0)
        ReDim vPtr(0 To 0)
    Else
        ReDim vType(0 To iParamCount - 1)
        ReDim vPtr(0 To iParamCount - 1)
        For lIdx = 0 To iParamCount - 1
            vType(lIdx) = VarType(vArgs(lb + lIdx))
            vPtr(lIdx) = VarPtr(vArgs(lb + lIdx))
        Next lIdx
    End If

    Dim hResult As Long
    Dim vv As Variant
    hResult = DispCallFunc(pInterface, VTableOffset * PTR_SZ, CC_STDCALL, ReturnType, iParamCount, VarPtr(vType(0)), VarPtr(vPtr(0)), vv)
    If raiseOnDispFailure And hResult < 0 Then
        Err.Raise hResult, "stdWebView::CallVT", "DispCallFunc - Unknown error occurred."
    End If
    CallVTImpl = vv
End Function

Private Function CallVT(ByVal pInterface As LongPtr, ByVal VTableOffset As Long, ByVal ReturnType As VbVarType, ParamArray FunctionParameters() As Variant) As Variant
    CallVT = CallVTImpl(pInterface, VTableOffset, ReturnType, True, FunctionParameters)
End Function

Private Function CallVTQuiet(ByVal pInterface As LongPtr, ByVal VTableOffset As Long, ByVal ReturnType As VbVarType, ParamArray FunctionParameters() As Variant) As Variant
    CallVTQuiet = CallVTImpl(pInterface, VTableOffset, ReturnType, False, FunctionParameters)
End Function

Private Function PtrToStrW(ByVal pWStr As LongPtr) As String
    Dim n As Long
    Dim buf As String
    If pWStr = 0 Then Exit Function
    n = lstrlenW(pWStr)
    If n <= 0 Then Exit Function
    buf = string$(n, vbNullChar)
    CopyMemory ByVal StrPtr(buf), ByVal pWStr, n * 2&
    PtrToStrW = buf
End Function

' ============================================================================
' IMPORTANT - DO NOT MOVE THESE METHODS,
' AND DO NOT ADD ANY NEW PUBLIC METHODS ABOVE
' OR INBETWEEN THEM!!
' This is because the WebView2 handler vtables (IUnknown + Invoke)
' depend on the order of the Public methods.
'
' `GetComProcAddress(Me, 0..6)` depends on that. Do not move other Public methods
' above them without adjusting the indexes.
' ============================================================================

'WebView2 QueryInterface Thunk Implementation. DO NOT CALL THIS METHOD.
'@protected
Public Function zzProtWebView_QI( _
    ByVal callbackThis As LongPtr, _
    ByVal riid As LongPtr, _
    ByRef ppvObject As LongPtr) As Long

    ppvObject = callbackThis
    zzProtWebView_QI = 0&
End Function

'WebView2 AddRef Thunk Implementation. DO NOT CALL THIS METHOD.
'@protected
Public Function zzProtWebView_AddRef(ByVal callbackThis As LongPtr) As Long
    zzProtWebView_AddRef = 1&
End Function

'WebView2 Release Thunk Implementation. DO NOT CALL THIS METHOD.
'@protected
Public Function zzProtWebView_Release(ByVal callbackThis As LongPtr) As Long
    zzProtWebView_Release = 1&
End Function

'WebView2 Environment Invoke Thunk Implementation. DO NOT CALL THIS METHOD.
'@protected
Public Function zzProtWebView_EnvInvoke( _
    ByVal callbackThis As LongPtr, _
    ByVal errorCode As Long, _
    ByVal pEnvironment As LongPtr) As Long

    On Error Resume Next
    Call protEnvCompleted(errorCode, pEnvironment)
    zzProtWebView_EnvInvoke = 0&
End Function

'WebView2 Controller Invoke Thunk Implementation. DO NOT CALL THIS METHOD.
'@protected
Public Function zzProtWebView_CtrlInvoke( _
    ByVal callbackThis As LongPtr, _
    ByVal errorCode As Long, _
    ByVal pController As LongPtr) As Long

    On Error Resume Next
    Call protCtrlCompleted(errorCode, pController)
    zzProtWebView_CtrlInvoke = 0&
End Function

'WebView2 Script Invoke Thunk Implementation. DO NOT CALL THIS METHOD.
'@protected
Public Function zzProtWebView_ScriptInvoke( _
    ByVal callbackThis As LongPtr, _
    ByVal errorCode As Long, _
    ByVal resultJsonPtr As LongPtr) As Long

    On Error Resume Next
    Dim reqIdLP As LongPtr
    Dim reqId As Long
    Call CopyMemory(reqIdLP, ByVal callbackThis + PTR_SZ, PTR_SZ)
    reqId = CLng(reqIdLP)
    Call protScriptCompleted(errorCode, PtrToStrW(resultJsonPtr), reqId)
    zzProtWebView_ScriptInvoke = 0&
End Function

'WebView2 WebResourceRequested Invoke Thunk Implementation. DO NOT CALL THIS METHOD.
'@protected
Public Function zzProtWebView_WebResourceRequestedInvoke( _
    ByVal callbackThis As LongPtr, _
    ByVal pSender As LongPtr, _
    ByVal pArgs As LongPtr) As Long

    On Error Resume Next
    Call protWebResourceRequested(pSender, pArgs)
    zzProtWebView_WebResourceRequestedInvoke = 0&
End Function

'WebView2 WebMessageReceived Invoke Thunk Implementation. DO NOT CALL THIS METHOD.
'@protected
Public Function zzProtWebView_WebMessageReceivedInvoke( _
    ByVal callbackThis As LongPtr, _
    ByVal pSender As LongPtr, _
    ByVal pArgs As LongPtr) As Long

    On Error Resume Next
    Call protWebMessageReceived(pSender, pArgs)
    zzProtWebView_WebMessageReceivedInvoke = 0&
End Function

' ============================================================================

'Create a stdWebView from a window handle.
'@constructor
'@param hwnd - Parent window handle that will host the WebView2 client area.
'@param OnReady as stdICallable<(stdWebView)=>void> - Optional callback run when WebView2 controller/CoreWebView2 is ready.
'@returns - Created stdWebView instance.
'@remark - Initialization blocks and pumps DoEvents until WebView2 is ready or times out.
Public Function CreateFromHwnd(ByVal hwnd As LongPtr, Optional ByVal OnReady As stdICallable = Nothing) As stdWebView
    Dim w As New stdWebView
    Call w.protCreate(hwnd, OnReady)
    Set CreateFromHwnd = w
End Function

'Create a stdWebView from an MSForms.Frame.
'@constructor
'@param frm - MSForms.Frame used as the host container.
'@param OnReady as stdICallable<(wv as stdWebView)=>void> - Optional callback run when WebView2 controller/CoreWebView2 is ready.
'@returns - Created stdWebView instance.
'@remark - Resolves the frame HWND and delegates to CreateFromHwnd.
Public Function CreateFromFrame(ByVal frm As Object, Optional ByVal OnReady As stdICallable = Nothing) As stdWebView
    If frm Is Nothing Then Err.Raise 5, "stdWebView::CreateFromFrame", "frm is Nothing"
    If TypeName(frm) <> "Frame" Then Err.Raise 5, "stdWebView::CreateFromFrame", "Expected MSForms.Frame"
    Set CreateFromFrame = CreateFromHwnd(FrameToHwnd(frm), OnReady)
End Function

'Create a stdWebView from a UserForm.
'@constructor
'@param uf - UserForm used as the host container.
'@param OnReady as stdICallable<(wv as stdWebView)=>void> - Optional callback run when WebView2 controller/CoreWebView2 is ready.
'@returns - Created stdWebView instance.
'@remark - Resolves the UserForm root HWND via IOleWindow, then hosts in its first child HWND.
Public Function CreateFromUserform(ByVal uf As Object, Optional ByVal OnReady As stdICallable = Nothing) As stdWebView
    If uf Is Nothing Then Err.Raise 5, "stdWebView::CreateFromUserform", "uf is Nothing"
    Set CreateFromUserform = CreateFromHwnd(UserFormToHwnd(uf), OnReady)
End Function

'Check whether the WebView2 is ready to be used.
'@returns - `True` when CoreWebView2 is available and navigation/script APIs can be used.
Public Function IsReady() As Boolean
    IsReady = (This.com.ppWebView2 <> 0)
End Function

'@remark - Releases controller/environment state and frees COM handler allocations. Safe to call more than once.
Public Sub Quit()
    On Error Resume Next
    Call WebMessage_RemoveEventHandler
    Call WebResource_RemoveEventHandler
    If This.com.pController <> 0 Then
        Call CallVTQuiet(This.com.pController, 24, vbLong)
    End If
    Set This.init.onReadyCallback = Nothing
    This.com.pController = 0
    This.com.ppWebView2 = 0
    This.com.pEnvironment = 0
    Call FreeHandlerAllocs
End Sub

'Navigate to a URL.
'@param url - URL to navigate to.
'@remark - Raises if the WebView is not ready.
'@TODO: Add a callback parameter.
Public Sub Navigate(ByVal url As String)
    If Not IsReady Then Err.Raise 5, "stdWebView::Navigate", "WebView not ready"
    Call CallVT(This.com.ppWebView2, 5, vbLong, StrPtr(url))
End Sub

'Get current page URL.
'@returns - Current `CoreWebView2.Source`.
'@remark - Raises if the WebView is not ready.
Public Property Get Source() As String
    Dim pSource As LongPtr
    If Not IsReady Then Err.Raise 5, "stdWebView::Source [Get]", "WebView not ready"
    Call CallVT(This.com.ppWebView2, COREWEBVIEW2_VTBL_GET_SOURCE, vbLong, VarPtr(pSource))
    Source = PtrToStrW(pSource)
    If pSource <> 0 Then Call CoTaskMemFree(pSource)
End Property

'Get/Let Address
Public Property Get Address() As String
    Address = Source
End Property
Public Property Let Address(v As String)
    Navigate v
End Property

'Wait until current page is fully loaded and URL prefix is stable.
'@param urlPrefix - Required URL prefix (case-insensitive). Pass "" to skip URL matching.
'@param timeoutMs - Max wait time in milliseconds.
'@param stableMs - Consecutive milliseconds condition must remain true before success.
'@param pollMs - Poll interval in milliseconds.
'@returns - `True` when ready condition is satisfied within timeout, else `False`.
Public Function WaitForDocumentReady( _
    ByVal urlPrefix As String, _
    Optional ByVal timeoutMs As Long = 120000, _
    Optional ByVal stableMs As Long = 1000, _
    Optional ByVal pollMs As Long = 200) As Boolean

    Dim requiredPrefix As String
    Dim startedAt As Double
    Dim stableStart As Double
    Dim readyState As String
    Dim currentUrl As String
    Dim isReadyNow As Boolean
    Dim hasPrefix As Boolean

    If Not IsReady Then Err.Raise 5, "stdWebView::WaitForDocumentReady", "WebView not ready"
    requiredPrefix = LCase$(Trim$(urlPrefix))
    If timeoutMs < 0 Then timeoutMs = 0
    If stableMs < 0 Then stableMs = 0
    If pollMs < 1 Then pollMs = 1

    startedAt = Timer
    stableStart = -1#

    Do
        currentUrl = LCase$(Trim$(Source))
        hasPrefix = (LenB(requiredPrefix) = 0) Or (Left$(currentUrl, Len(requiredPrefix)) = requiredPrefix)

        If hasPrefix Then
            readyState = LCase$(Trim$(JsonUnquoteString(JavaScriptRunSync("(function(){try{return document.readyState||'';}catch(e){return '';}})()"))))
            isReadyNow = (readyState = "complete")
        Else
            isReadyNow = False
        End If

        If isReadyNow Then
            If stableStart < 0 Then stableStart = Timer
            If TimerElapsedMs(stableStart, Timer) >= CDbl(stableMs) Then
                WaitForDocumentReady = True
                Exit Function
            End If
        Else
            stableStart = -1#
        End If

        If TimerElapsedMs(startedAt, Timer) >= CDbl(timeoutMs) Then Exit Do
        Call SpinWaitMs(pollMs)
    Loop
End Function

'Nagivate backward
'@remark - Navigates backward using injected `history.back()`.
Public Sub Back()
    Call JavaScriptRunSync("history.back(); void 0")
End Sub

'Navigate forward
'@remark - Navigates forward using injected `history.forward()`.
Public Sub Forward()
    Call JavaScriptRunSync("history.forward(); void 0")
End Sub

'Get the current document outer HTML.
'@returns - Current document outer HTML (`document.documentElement.outerHTML`).
'@remark - Raises if the WebView is not ready.
Public Property Get Html() As String
    If Not IsReady Then Err.Raise 5, "stdWebView::Html [Get]", "WebView not ready"
    Html = JsonUnquoteString(JavaScriptRunSync("(function(){try{return document.documentElement?document.documentElement.outerHTML:'';}catch(e){return '';}})()"))
End Property

'Set the current document outer HTML.
'@param rhs - HTML content loaded with WebView2 `NavigateToString`.
'@remark - Raises if the WebView is not ready.
Public Property Let Html(ByVal rhs As String)
    If Not IsReady Then Err.Raise 5, "stdWebView::Html [Let]", "WebView not ready"
    This.lastHtml = rhs
    Call CallVT(This.com.ppWebView2, 6, vbLong, StrPtr(rhs))
End Property

'Execute a JavaScript script synchronously.
'@param script - JavaScript executed in the page context.
'@returns - JSON-encoded result string returned by WebView2 script execution.
'@remark - Blocks with DoEvents until completion. Only one synchronous run can be active at a time.
Public Function JavaScriptRunSync(ByVal script As String) As String
    Dim bridgeRequestId As String
    Dim wrappedScript As String
    Dim startedAt As Double
    Dim elapsedMs As Double

    If Not IsReady Then Err.Raise 5, "stdWebView::JavaScriptRunSync", "WebView not ready"
    If This.script.syncWaitRequestId <> 0 Then
        Err.Raise 5, "stdWebView::JavaScriptRunSync", "A synchronous script call is already in progress"
    End If
    Call Scripts_ClearCompleted
    Call WebMessage_EnsureEventHandler
    This.script.syncDone = False
    This.script.result = vbNullString
    bridgeRequestId = CStr(NextScriptRequestId())
    wrappedScript = BuildSyncBridgeScript(script, bridgeRequestId)
    Call ScriptSyncBridge_Reset
    This.script.syncBridgeRequestId = bridgeRequestId
    This.script.syncWaitRequestId = JavaScriptRunCore(wrappedScript, Nothing, False)
    If This.script.syncWaitRequestId = 0 Then
        Call ScriptSyncBridge_Reset
        This.script.syncWaitRequestId = 0
        Err.Raise 5, "stdWebView::JavaScriptRunSync", "Failed to queue script execution"
    End If
    startedAt = Timer
    Do
        If This.script.syncBridgeDone Then Exit Do
        elapsedMs = TimerElapsedMs(startedAt, Timer)
        If elapsedMs >= CDbl(SCRIPT_SYNC_TIMEOUT_MS) Then
            This.script.syncWaitRequestId = 0
            Call ScriptSyncBridge_Reset
            Call Scripts_ClearCompleted
            Err.Raise 5, "stdWebView::JavaScriptRunSync", "Timed out waiting for script result"
        End If
        DoEvents
    Loop
    This.script.syncWaitRequestId = 0
    Call Scripts_ClearCompleted
    If This.script.syncBridgeOk Then
        JavaScriptRunSync = This.script.syncBridgePayload
    Else
        Err.Raise 5, "stdWebView::JavaScriptRunSync", JsonUnquoteString(This.script.syncBridgePayload)
    End If
    Call ScriptSyncBridge_Reset
End Function

'Execute a JavaScript script asynchronously.
'@param script - JavaScript executed in the page context.
'@param callback as stdICallable<(errorCode as long, resultJson as string)=>void> - Optional callback invoked when script execution completes.
'@remark - Runs asynchronously. Callback receives `Array(errorCode, resultJson)`.
Public Sub JavaScriptRun(ByVal script As String, Optional ByVal callback As stdICallable = Nothing)
    If Not IsReady Then Err.Raise 5, "stdWebView::JavaScriptRun", "WebView not ready"
    Call Scripts_ClearCompleted
    If JavaScriptRunCore(script, callback, False) = 0 Then
        Err.Raise 5, "stdWebView::JavaScriptRun", "Failed to queue script execution"
    End If
End Sub

'Inject a VBA IDispatch object into JavaScript (`chrome.webview.hostObjects.<name>`).
'@param name - JavaScript host object name.
'@param hostObject - VBA object implementing IDispatch.
'@remark - Exposes async proxy access in JavaScript. Use `await chrome.webview.hostObjects.<name>.Method()`.
Public Sub AddHostObject(ByVal name As String, ByVal hostObject As Object)
    Dim hostName As String
    Dim hostVariant As Variant
    hostName = Trim$(name)
    If Not IsReady Then Err.Raise 5, "stdWebView::AddHostObject", "WebView not ready"
    If LenB(hostName) = 0 Then Err.Raise 5, "stdWebView::AddHostObject", "name cannot be blank"
    If hostObject Is Nothing Then Err.Raise 5, "stdWebView::AddHostObject", "hostObject is Nothing"

    Set hostVariant = hostObject
    Call CallVT(This.com.ppWebView2, COREWEBVIEW2_VTBL_ADDHOSTOBJECTTOSCRIPT, vbLong, StrPtr(hostName), VarPtr(hostVariant))
    Call HostObjects_SetRef(hostName, hostObject)
End Sub

'Remove a previously injected JavaScript host object.
'@param name - JavaScript host object name.
'@remark - Removes `chrome.webview.hostObjects.<name>`.
Public Sub RemoveHostObject(ByVal name As String)
    Dim hostName As String
    hostName = Trim$(name)
    If Not IsReady Then Err.Raise 5, "stdWebView::RemoveHostObject", "WebView not ready"
    If LenB(hostName) = 0 Then Err.Raise 5, "stdWebView::RemoveHostObject", "name cannot be blank"

    Call CallVT(This.com.ppWebView2, COREWEBVIEW2_VTBL_REMOVEHOSTOBJECTFROMSCRIPT, vbLong, StrPtr(hostName))
    Call HostObjects_RemoveRef(hostName)
End Sub

'Map a virtual host name to a local folder path for loading local content.
'@param hostName - Virtual host name (e.g. "app.local").
'@param folderPath - Local folder path to map (e.g. "C:\MyApp\wwwroot").
'@param accessKind - Access level for resources (ALLOW, DENY, or DENY_CORS).
'@remark - Navigate to `https://<hostName>/path` to load local content from `<folderPath>\path`.
Public Sub SetVirtualHostNameToFolderMapping(ByVal hostName As String, ByVal folderPath As String, Optional ByVal accessKind As COREWEBVIEW2_HOST_RESOURCE_ACCESS_KIND = COREWEBVIEW2_HOST_RESOURCE_ACCESS_KIND_ALLOW)
    Dim host As String
    Dim folder As String
    host = Trim$(hostName)
    folder = Trim$(folderPath)
    If Not IsReady Then Err.Raise 5, "stdWebView::SetVirtualHostNameToFolderMapping", "WebView not ready"
    If LenB(host) = 0 Then Err.Raise 5, "stdWebView::SetVirtualHostNameToFolderMapping", "hostName cannot be blank"
    If LenB(folder) = 0 Then Err.Raise 5, "stdWebView::SetVirtualHostNameToFolderMapping", "folderPath cannot be blank"

    Call CallVT(This.com.ppWebView2, COREWEBVIEW2_VTBL_SETVIRTUALHOSTNAMETOFOLDERMAPPING, vbLong, StrPtr(host), StrPtr(folder), CLng(accessKind))
End Sub

'Clear a previously set virtual host name to folder mapping.
'@param hostName - Virtual host name to clear (e.g. "app.local").
'@remark - Removes the mapping set by SetVirtualHostNameToFolderMapping.
Public Sub ClearVirtualHostNameToFolderMapping(ByVal hostName As String)
    Dim host As String
    host = Trim$(hostName)
    If Not IsReady Then Err.Raise 5, "stdWebView::ClearVirtualHostNameToFolderMapping", "WebView not ready"
    If LenB(host) = 0 Then Err.Raise 5, "stdWebView::ClearVirtualHostNameToFolderMapping", "hostName cannot be blank"

    Call CallVT(This.com.ppWebView2, COREWEBVIEW2_VTBL_CLEARVIRTUALHOSTNAMETOFOLDERMAPPING, vbLong, StrPtr(host))
End Sub

'Intercept matching web resource requests via a callback.
'@param filterString - URI filter pattern (for example `https://example.com/*`).
'@param callback as stdICallable<(request as Dictionary)=>void> - Callback that mutates request envelope keys.
'@param context - Optional resource context filter (`COREWEBVIEW2_WEB_RESOURCE_CONTEXT_*`).
'@remark - Callback may either set `respond=True` + response fields to short-circuit, or mutate request headers and continue.
Public Sub AddWebResourceRequestedFilter(ByVal filterString As String, ByVal callback As stdICallable, Optional ByVal context As ECoreWebView2WebResourceContext = COREWEBVIEW2_WEB_RESOURCE_CONTEXT_ALL)
    Dim pattern As String
    Dim idx As Long
    Dim hr As Long

    If Not IsReady Then Err.Raise 5, "stdWebView::AddWebResourceRequestedFilter", "WebView not ready"
    pattern = Trim$(filterString)
    If LenB(pattern) = 0 Then Err.Raise 5, "stdWebView::AddWebResourceRequestedFilter", "filterString cannot be blank"
    If callback Is Nothing Then Err.Raise 5, "stdWebView::AddWebResourceRequestedFilter", "callback is Nothing"

    Call WebResource_EnsureEventHandler
    hr = CLng(CallVT(This.com.ppWebView2, COREWEBVIEW2_VTBL_ADDWEBRESOURCEREQUESTEDFILTER, vbLong, StrPtr(pattern), CLng(context)))
    If hr < 0 Then Err.Raise hr, "stdWebView::AddWebResourceRequestedFilter", "AddWebResourceRequestedFilter failed (" & CStr(hr) & ")"

    idx = WebResource_AddFilter(pattern, CLng(context), callback)
    If idx < 0 Then Err.Raise 5, "stdWebView::AddWebResourceRequestedFilter", "Failed to register filter callback"
End Sub

'Complete a deferred web resource request with a custom response.
'@param requestId - Request identifier provided to the callback envelope.
'@param responseEnvelope - Dictionary containing response fields (`responseBody`, `responseStatusCode`, etc.).
Public Sub WebResourceRequest_Complete(ByVal requestId As Long, ByVal responseEnvelope As Object)
    Dim idx As Long
    Dim pArgs As LongPtr
    Dim pResponse As LongPtr
    Dim hr As Long

    idx = DeferredWebResource_FindIndex(requestId)
    If idx < 0 Then Err.Raise 5, "stdWebView::WebResourceRequest_Complete", "Unknown deferred requestId: " & CStr(requestId)

    pArgs = This.deferredWebResource.items(idx).pArgs
    If Not responseEnvelope Is Nothing Then
        pResponse = WebResource_CreateResponse(responseEnvelope)
        If pResponse <> 0 Then
            hr = CLng(CallVT(pArgs, COREWEBVIEW2WEBRESOURCEREQUESTEDEVENTARGS_VTBL_PUT_RESPONSE, vbLong, pResponse))
            Call ComRelease(pResponse)
            If hr < 0 Then
                Call DeferredWebResource_CompleteAndRemoveAt(idx, True)
                Err.Raise hr, "stdWebView::WebResourceRequest_Complete", "put_Response failed (" & CStr(hr) & ")"
            End If
        End If
    End If

    Call DeferredWebResource_CompleteAndRemoveAt(idx, True)
End Sub

'Fail/continue a deferred web resource request without custom response.
'@param requestId - Request identifier provided to the callback envelope.
Public Sub WebResourceRequest_Fail(ByVal requestId As Long)
    Dim idx As Long
    idx = DeferredWebResource_FindIndex(requestId)
    If idx < 0 Then Exit Sub
    Call DeferredWebResource_CompleteAndRemoveAt(idx, True)
End Sub

'Resize WebView2 surface to fill the current host client area.
'@remark - Call from host resize events (for example UserForm_Resize).
Public Sub Resize()
    If This.com.pController = 0 Then Exit Sub
    Dim r As RECT
    Call GetClientRect(This.parentHwnd, r)
    Call ControllerPutBounds(0&, 0&, r.Right - r.Left, r.Bottom - r.Top)
    Call ResizeChromeWidgets
End Sub

'Protected method to handle the environment creation callback.
'@protected
'@param errorCode - HRESULT from environment creation callback.
'@param pEnvironment - IWebView2Environment pointer returned by WebView2.
Friend Sub protEnvCompleted(ByVal errorCode As Long, ByVal pEnvironment As LongPtr)
    If errorCode <> 0 Then
        This.init.waiting = False
        This.init.failed = True
        Exit Sub
    End If
    This.com.pEnvironment = pEnvironment
    This.com.pCtrlVTable = BuildHandlerVTable(4, This.com.ctrlThunk)
    This.com.pCtrlHandlerObj = GlobalAlloc(GPTR, PTR_SZ)
    Call CopyMemory(ByVal This.com.pCtrlHandlerObj, This.com.pCtrlVTable, PTR_SZ)
    Call CallVT(pEnvironment, 3, vbLong, This.parentHwnd, This.com.pCtrlHandlerObj)
End Sub

'Protected method to handle the controller creation callback.
'@protected
'@param errorCode - HRESULT from controller creation callback.
'@param pController - IWebView2Controller pointer returned by WebView2.
Friend Sub protCtrlCompleted(ByVal errorCode As Long, ByVal pController As LongPtr)
    If errorCode <> 0 Then
        This.init.waiting = False
        This.init.failed = True
        Exit Sub
    End If
    Call ComAddRef(pController)
    This.com.pController = pController
    Dim r As RECT
    Call GetClientRect(This.parentHwnd, r)
    Call ControllerPutBounds(0&, 0&, r.Right - r.Left, r.Bottom - r.Top)
    Call ControllerPutVisible(True)
    Call ControllerGetCoreWebView2
    Call ResizeChromeWidgets
    This.script.syncDone = True
    This.init.waiting = False
    If Not This.init.onReadyCallback Is Nothing Then
        Dim cbReady As stdICallable
        Set cbReady = This.init.onReadyCallback
        Set This.init.onReadyCallback = Nothing
        On Error Resume Next
        Call cbReady.RunEx(Array(Me))
    End If
End Sub

'Protected method to handle the script completion callback.
'@protected
'@param errorCode - HRESULT from script completion callback.
'@param resultJson - JSON result string returned by WebView2.
'@param requestId - Internal request identifier for queued script work.
Friend Sub protScriptCompleted(ByVal errorCode As Long, ByVal resultJson As String, ByVal requestId As Long)
    Dim idx As Long
    Dim cb As stdICallable
    Dim isSyncWait As Boolean

    idx = Scripts_FindScriptIndex(requestId)
    If idx < 0 Then Exit Sub

    If This.scripts.mScripts(idx).isCompleted Then Exit Sub
    This.scripts.mScripts(idx).isCompleted = True
    If This.scripts.mScripts(idx).callback Is Nothing Then
        Set cb = Nothing
    Else
        Set cb = This.scripts.mScripts(idx).callback
    End If
    Set This.scripts.mScripts(idx).callback = Nothing
    isSyncWait = This.scripts.mScripts(idx).isSyncWait

    If isSyncWait And This.script.syncWaitRequestId = requestId Then
        This.script.result = resultJson
        This.script.syncDone = True
    End If

    If Not cb Is Nothing Then
        On Error Resume Next
        Call cb.RunEx(Array(errorCode, resultJson))
    End If
End Sub

Friend Sub protWebMessageReceived(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    Dim pMessageJson As LongPtr
    Dim messageJson As String
    Dim messageText As String

    If pArgs = 0 Then Exit Sub
    If This.script.syncWaitRequestId = 0 Then Exit Sub

    Call CallVT(pArgs, COREWEBVIEW2WEBMESSAGERECEIVEDEVENTARGS_VTBL_GET_WEBMESSAGEASJSON, vbLong, VarPtr(pMessageJson))
    messageJson = PtrToStrW(pMessageJson)
    If pMessageJson <> 0 Then Call CoTaskMemFree(pMessageJson)
    messageText = JsonUnquoteString(messageJson)
    Call ScriptSyncBridge_TryConsumeMessage(messageText)
End Sub

Friend Sub protWebResourceRequested(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    Dim pRequest As LongPtr
    Dim pHeaders As LongPtr
    Dim pResponse As LongPtr
    Dim pUri As LongPtr
    Dim pMethod As LongPtr
    Dim uri As String
    Dim method As String
    Dim contextValue As Long
    Dim i As Long
    Dim envelope As Object
    Dim cb As stdICallable
    Dim requestId As Long
    Dim shouldDefer As Boolean
    Dim pDeferral As LongPtr
    Dim hr As Long

    If This.webResource.count <= 0 Then Exit Sub
    If pArgs = 0 Then Exit Sub

    Call CallVT(pArgs, COREWEBVIEW2WEBRESOURCEREQUESTEDEVENTARGS_VTBL_GET_REQUEST, vbLong, VarPtr(pRequest))
    If pRequest = 0 Then Exit Sub

    On Error Resume Next
    Call CallVT(pArgs, COREWEBVIEW2WEBRESOURCEREQUESTEDEVENTARGS_VTBL_GET_RESOURCECONTEXT, vbLong, VarPtr(contextValue))
    On Error GoTo fail

    Call CallVT(pRequest, COREWEBVIEW2WEBRESOURCEREQUEST_VTBL_GET_URI, vbLong, VarPtr(pUri))
    Call CallVT(pRequest, COREWEBVIEW2WEBRESOURCEREQUEST_VTBL_GET_METHOD, vbLong, VarPtr(pMethod))
    uri = PtrToStrW(pUri)
    method = PtrToStrW(pMethod)
    If pUri <> 0 Then Call CoTaskMemFree(pUri)
    If pMethod <> 0 Then Call CoTaskMemFree(pMethod)
    pUri = 0
    pMethod = 0

    Call CallVT(pRequest, COREWEBVIEW2WEBRESOURCEREQUEST_VTBL_GET_HEADERS, vbLong, VarPtr(pHeaders))
    For i = 0 To This.webResource.count - 1
        If WebResource_FilterMatches(This.webResource.items(i), uri, contextValue) Then
            Set cb = This.webResource.items(i).callback
            If cb Is Nothing Then GoTo nextFilter

            requestId = DeferredWebResource_NextRequestId()
            Set envelope = WebResource_BuildEnvelope(uri, method, contextValue, pHeaders, requestId)
            On Error Resume Next
            Call cb.RunEx(Array(envelope))
            On Error GoTo fail

            Call WebResource_ApplyHeaderChanges(envelope, pHeaders)
            shouldDefer = WebResource_ShouldDefer(envelope)
            If shouldDefer Then
                hr = CLng(CallVT(pArgs, COREWEBVIEW2WEBRESOURCEREQUESTEDEVENTARGS_VTBL_GETDEFERRAL, vbLong, VarPtr(pDeferral)))
                If hr < 0 Or pDeferral = 0 Then
                    Err.Raise IIf(hr < 0, hr, 5), "stdWebView::protWebResourceRequested", "GetDeferral failed"
                End If
                Call DeferredWebResource_Add(requestId, pArgs, pDeferral)
                Call ComRelease(pDeferral)
                pDeferral = 0
                Exit For
            End If

            If WebResource_ShouldRespond(envelope) Then
                pResponse = WebResource_CreateResponse(envelope)
                If pResponse <> 0 Then
                    Call CallVT(pArgs, COREWEBVIEW2WEBRESOURCEREQUESTEDEVENTARGS_VTBL_PUT_RESPONSE, vbLong, pResponse)
                    Call ComRelease(pResponse)
                    pResponse = 0
                End If
                Exit For
            End If
        End If
nextFilter:
    Next i

cleanup:
    If pDeferral <> 0 Then Call ComRelease(pDeferral)
    If pHeaders <> 0 Then Call ComRelease(pHeaders)
    If pRequest <> 0 Then Call ComRelease(pRequest)
    Exit Sub

fail:
    Resume cleanup
End Sub

'Protected method to create the stdWebView instance.
'@constructor
'@protected
'@param hwnd - Parent window handle used to host WebView2.
'@param OnReady as stdICallable<(stdWebView)=>void> - Optional callback run when initialization completes.
'@remark - Internal creation entry point used by public constructors.
Friend Sub protCreate(ByVal hwnd As LongPtr, Optional ByVal OnReady As stdICallable = Nothing)
    If hwnd = 0 Then Err.Raise 5, "stdWebView::CreateFromHwnd", "hwnd is 0"
    This.parentHwnd = hwnd
    This.init.waiting = True
    This.init.failed = False
    This.script.syncDone = False
    This.script.syncWaitRequestId = 0
    This.script.syncBridgeRequestId = vbNullString
    This.script.syncBridgeDone = False
    This.script.syncBridgeOk = False
    This.script.syncBridgePayload = vbNullString
    This.script.nextRequestId = 0
    This.scripts.scriptCount = 0
    This.deferredWebResource.count = 0
    This.deferredWebResource.nextRequestId = 0
    This.com.webReqTokenLow = 0
    This.com.webReqTokenHigh = 0
    This.com.webMsgTokenLow = 0
    This.com.webMsgTokenHigh = 0
    Erase This.scripts.mScripts
    Erase This.deferredWebResource.items
    Set This.init.onReadyCallback = OnReady
    Randomize
    This.init.userDataFolder = Environ$("TEMP") & "\stdWebView_" & CStr(CLng(Timer * 10000) Xor CLng(Rnd * 1000000))
    This.com.pEnvVTable = BuildHandlerVTable(3, This.com.envThunk)
    This.com.pEnvHandlerObj = GlobalAlloc(GPTR, PTR_SZ)
    Call CopyMemory(ByVal This.com.pEnvHandlerObj, This.com.pEnvVTable, PTR_SZ)
    Dim hr As Long
    hr = CreateCoreWebView2EnvironmentWithOptions(0, StrPtr(This.init.userDataFolder), 0, This.com.pEnvHandlerObj)
    If hr <> 0 Then
        Call CleanupFailedInit
        Err.Raise hr, "stdWebView", "CreateCoreWebView2EnvironmentWithOptions failed (" & CStr(hr) & ")"
    End If
    Dim j As Long
    For j = 1 To 60000
        If Not This.init.waiting Then Exit For
        DoEvents
    Next j
    If This.init.failed Or This.com.ppWebView2 = 0 Then
        Call CleanupFailedInit
        Err.Raise &H80040201, "stdWebView", "WebView2 initialization failed or timed out"
    End If
End Sub

Private Sub ComAddRef(ByVal pUnk As LongPtr)
    Call CallVT(pUnk, 1, vbLong)
End Sub

Private Sub ComRelease(ByVal pUnk As LongPtr)
    If pUnk <> 0 Then Call CallVTQuiet(pUnk, 2, vbLong)
End Sub

Private Sub ControllerPutBounds(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    Dim rc As RECT
    rc.Left = x
    rc.Top = y
    rc.Right = x + w
    rc.Bottom = y + h
    Call CallVT(This.com.pController, 6, vbLong, VarPtr(rc))
End Sub

Private Sub ControllerPutVisible(ByVal visible As Boolean)
    Call CallVT(This.com.pController, 4, vbLong, CLng(IIf(visible, 1, 0)))
End Sub

Private Sub ControllerGetCoreWebView2()
    Dim pp As LongPtr
    Call CallVT(This.com.pController, 25, vbLong, VarPtr(pp))
    This.com.ppWebView2 = pp
End Sub

Private Sub ResizeChromeWidgets()
    Dim r As RECT
    Call GetClientRect(This.parentHwnd, r)
    Dim w As Long
    Dim h As Long
    w = r.Right - r.Left
    h = r.Bottom - r.Top
    Dim ch0 As LongPtr
    ch0 = FindWindowEx(This.parentHwnd, 0, "Chrome_WidgetWin_0", vbNullString)
    If ch0 <> 0 Then
        Call MoveWindow(ch0, 0, 0, w, h, 1)
        Dim ch1 As LongPtr
        ch1 = FindWindowEx(ch0, 0, "Chrome_WidgetWin_1", vbNullString)
        If ch1 <> 0 Then
            Call MoveWindow(ch1, 0, 0, w, h, 1)
            Dim chD3D As LongPtr
            chD3D = FindWindowEx(ch1, 0, "Intermediate D3D Window", vbNullString)
            If chD3D <> 0 Then Call MoveWindow(chD3D, 0, 0, w, h, 1)
        End If
    End If
End Sub

Private Function BuildHandlerVTable( _
    ByVal invokePublicIndex As Long, _
    ByRef thunks As THandlerThunks, _
    Optional ByVal invokeThunkArgCount As Long = 3) As LongPtr

    Dim vTable(3) As LongPtr
    Dim sz As LongPtr
    Dim pProc As LongPtr

    EraseThunkStruct thunks

    pProc = GetComProcAddress(Me, 0)
    If pProc = 0 Then Err.Raise 5, , "stdWebView: failed to resolve QI."
    thunks.pQI = AddressOfThunkForComProc(Me, pProc, 3, sz)
    If thunks.pQI = 0 Then GoTo fail

    pProc = GetComProcAddress(Me, 1)
    If pProc = 0 Then GoTo fail
    thunks.pAddRef = AddressOfThunkForComProc(Me, pProc, 1, sz)
    If thunks.pAddRef = 0 Then GoTo fail

    pProc = GetComProcAddress(Me, 2)
    If pProc = 0 Then GoTo fail
    thunks.pRelease = AddressOfThunkForComProc(Me, pProc, 1, sz)
    If thunks.pRelease = 0 Then GoTo fail

    pProc = GetComProcAddress(Me, invokePublicIndex)
    If pProc = 0 Then GoTo fail
    thunks.pInvoke = AddressOfThunkForComProc(Me, pProc, invokeThunkArgCount, sz)
    If thunks.pInvoke = 0 Then GoTo fail

    vTable(0) = thunks.pQI
    vTable(1) = thunks.pAddRef
    vTable(2) = thunks.pRelease
    vTable(3) = thunks.pInvoke

    BuildHandlerVTable = GlobalAlloc(GPTR, 4& * PTR_SZ)
    If BuildHandlerVTable = 0 Then GoTo fail
    Call CopyMemory(ByVal BuildHandlerVTable, vTable(0), 4& * PTR_SZ)
    Exit Function

fail:
    Call FreeHandlerThunks(thunks)
    Err.Raise 5, , "stdWebView: failed to build handler vtable/thunks."
End Function

Private Sub EraseThunkStruct(ByRef thunks As THandlerThunks)
    thunks.pQI = 0
    thunks.pAddRef = 0
    thunks.pRelease = 0
    thunks.pInvoke = 0
End Sub

Private Sub FreeHandlerThunks(ByRef thunks As THandlerThunks)
    If thunks.pQI <> 0 Then
        Call VirtualFree(thunks.pQI, 0, MEM_RELEASE)
        thunks.pQI = 0
    End If
    If thunks.pAddRef <> 0 Then
        Call VirtualFree(thunks.pAddRef, 0, MEM_RELEASE)
        thunks.pAddRef = 0
    End If
    If thunks.pRelease <> 0 Then
        Call VirtualFree(thunks.pRelease, 0, MEM_RELEASE)
        thunks.pRelease = 0
    End If
    If thunks.pInvoke <> 0 Then
        Call VirtualFree(thunks.pInvoke, 0, MEM_RELEASE)
        thunks.pInvoke = 0
    End If
End Sub

Private Sub CleanupFailedInit()
    On Error Resume Next
    Set This.init.onReadyCallback = Nothing
    Call FreeHandlerAllocs
End Sub

Private Sub FreeHandlerAllocs()
    On Error Resume Next
    Set This.init.onReadyCallback = Nothing
    Call HostObjects_ClearAll
    Call FreeHandlerThunks(This.com.envThunk)
    Call FreeHandlerThunks(This.com.ctrlThunk)
    Call FreeHandlerThunks(This.com.webReqThunk)
    Call FreeHandlerThunks(This.com.webMsgThunk)
    If This.com.pEnvHandlerObj <> 0 Then
        Call GlobalFree(This.com.pEnvHandlerObj)
        This.com.pEnvHandlerObj = 0
    End If
    If This.com.pEnvVTable <> 0 Then
        Call GlobalFree(This.com.pEnvVTable)
        This.com.pEnvVTable = 0
    End If
    If This.com.pCtrlHandlerObj <> 0 Then
        Call GlobalFree(This.com.pCtrlHandlerObj)
        This.com.pCtrlHandlerObj = 0
    End If
    If This.com.pCtrlVTable <> 0 Then
        Call GlobalFree(This.com.pCtrlVTable)
        This.com.pCtrlVTable = 0
    End If
    If This.com.pWebReqHandlerObj <> 0 Then
        Call GlobalFree(This.com.pWebReqHandlerObj)
        This.com.pWebReqHandlerObj = 0
    End If
    If This.com.pWebReqVTable <> 0 Then
        Call GlobalFree(This.com.pWebReqVTable)
        This.com.pWebReqVTable = 0
    End If
    If This.com.pWebMsgHandlerObj <> 0 Then
        Call GlobalFree(This.com.pWebMsgHandlerObj)
        This.com.pWebMsgHandlerObj = 0
    End If
    If This.com.pWebMsgVTable <> 0 Then
        Call GlobalFree(This.com.pWebMsgVTable)
        This.com.pWebMsgVTable = 0
    End If
    This.com.webReqTokenLow = 0
    This.com.webReqTokenHigh = 0
    This.com.webMsgTokenLow = 0
    This.com.webMsgTokenHigh = 0
    Call ScriptSyncBridge_Reset
    Call DeferredWebResource_ClearAll(True)
    Call WebResource_ClearFilters
    Call Scripts_ClearAll
End Sub

Private Function JavaScriptRunCore(ByVal script As String, ByVal callback As stdICallable, ByVal isSyncWait As Boolean) As Long
    Dim thunks As THandlerThunks
    Dim pVTable As LongPtr
    Dim pObj As LongPtr
    Dim reqId As Long
    Dim reqIdLP As LongPtr
    Dim idx As Long
    idx = -1
    On Error GoTo fail

    reqId = NextScriptRequestId()
    pVTable = BuildHandlerVTable(5, thunks, 3)
    pObj = GlobalAlloc(GPTR, 2& * PTR_SZ)
    If pObj = 0 Then GoTo fail
    Call CopyMemory(ByVal pObj, pVTable, PTR_SZ)
    reqIdLP = CLngPtr(reqId)
    Call CopyMemory(ByVal pObj + PTR_SZ, reqIdLP, PTR_SZ)

    idx = Scripts_AddScript(reqId, callback, isSyncWait, pObj, pVTable, thunks)
    Call CallVT(This.com.ppWebView2, COREWEBVIEW2_VTBL_EXECUTESCRIPT, vbLong, StrPtr(script), pObj)
    On Error GoTo 0
    JavaScriptRunCore = reqId
    Exit Function

fail:
    If idx >= 0 Then Call Scripts_RemoveScriptAt(idx)
    If pObj <> 0 Then Call GlobalFree(pObj)
    If pVTable <> 0 Then Call GlobalFree(pVTable)
    Call FreeHandlerThunks(thunks)
End Function

Private Sub HostObjects_EnsureMap()
    If This.host.refs Is Nothing Then Set This.host.refs = New Collection
End Sub

Private Sub HostObjects_SetRef(ByVal name As String, ByVal hostObject As Object)
    Call HostObjects_EnsureMap
    On Error Resume Next
    This.host.refs.Remove name
    Err.Clear
    On Error GoTo 0
    This.host.refs.Add hostObject, name
End Sub

Private Sub HostObjects_RemoveRef(ByVal name As String)
    If This.host.refs Is Nothing Then Exit Sub
    On Error Resume Next
    This.host.refs.Remove name
    On Error GoTo 0
End Sub

Private Sub HostObjects_ClearAll()
    Set This.host.refs = Nothing
End Sub

Private Sub WebResource_EnsureCapacity()
    Dim cap As Long
    If This.webResource.count < 0 Then This.webResource.count = 0
    cap = WebResource_Capacity()
    If This.webResource.count < cap Then Exit Sub
    cap = cap + 4
    ReDim Preserve This.webResource.items(0 To cap - 1)
End Sub

Private Function WebResource_Capacity() As Long
    On Error GoTo noArray
    WebResource_Capacity = UBound(This.webResource.items) - LBound(This.webResource.items) + 1
    Exit Function
noArray:
    WebResource_Capacity = 0
End Function

Private Function WebResource_AddFilter(ByVal filterString As String, ByVal context As Long, ByVal callback As stdICallable) As Long
    Dim idx As Long
    Call WebResource_EnsureCapacity
    idx = This.webResource.count
    This.webResource.items(idx).filterString = filterString
    This.webResource.items(idx).context = context
    Set This.webResource.items(idx).callback = callback
    This.webResource.count = This.webResource.count + 1
    WebResource_AddFilter = idx
End Function

Private Sub WebResource_ClearFilters()
    Dim i As Long
    For i = 0 To This.webResource.count - 1
        Set This.webResource.items(i).callback = Nothing
        This.webResource.items(i).filterString = vbNullString
        This.webResource.items(i).context = 0
    Next i
    This.webResource.count = 0
    Erase This.webResource.items
End Sub

Private Sub DeferredWebResource_EnsureCapacity()
    Dim cap As Long
    If This.deferredWebResource.count < 0 Then This.deferredWebResource.count = 0
    cap = DeferredWebResource_Capacity()
    If This.deferredWebResource.count < cap Then Exit Sub
    cap = cap + 8
    ReDim Preserve This.deferredWebResource.items(0 To cap - 1)
End Sub

Private Function DeferredWebResource_Capacity() As Long
    On Error GoTo noArray
    DeferredWebResource_Capacity = UBound(This.deferredWebResource.items) - LBound(This.deferredWebResource.items) + 1
    Exit Function
noArray:
    DeferredWebResource_Capacity = 0
End Function

Private Function DeferredWebResource_NextRequestId() As Long
    This.deferredWebResource.nextRequestId = This.deferredWebResource.nextRequestId + 1
    If This.deferredWebResource.nextRequestId <= 0 Then This.deferredWebResource.nextRequestId = 1
    DeferredWebResource_NextRequestId = This.deferredWebResource.nextRequestId
End Function

Private Sub DeferredWebResource_Add(ByVal requestId As Long, ByVal pArgs As LongPtr, ByVal pDeferral As LongPtr)
    Dim idx As Long
    Call DeferredWebResource_EnsureCapacity
    idx = This.deferredWebResource.count
    This.deferredWebResource.items(idx).requestId = requestId
    This.deferredWebResource.items(idx).pArgs = pArgs
    This.deferredWebResource.items(idx).pDeferral = pDeferral
    This.deferredWebResource.items(idx).createdAt = Timer
    Call ComAddRef(pArgs)
    Call ComAddRef(pDeferral)
    This.deferredWebResource.count = This.deferredWebResource.count + 1
End Sub

Private Function DeferredWebResource_FindIndex(ByVal requestId As Long) As Long
    Dim i As Long
    For i = 0 To This.deferredWebResource.count - 1
        If This.deferredWebResource.items(i).requestId = requestId Then
            DeferredWebResource_FindIndex = i
            Exit Function
        End If
    Next i
    DeferredWebResource_FindIndex = -1
End Function

Private Sub DeferredWebResource_CompleteAndRemoveAt(ByVal idx As Long, ByVal completeDeferral As Boolean)
    Dim lastIdx As Long
    Dim pArgs As LongPtr
    Dim pDeferral As LongPtr

    lastIdx = This.deferredWebResource.count - 1
    If idx < 0 Or idx > lastIdx Then Exit Sub

    pArgs = This.deferredWebResource.items(idx).pArgs
    pDeferral = This.deferredWebResource.items(idx).pDeferral

    If completeDeferral And pDeferral <> 0 Then
        Call CallVTQuiet(pDeferral, COREWEBVIEW2DEFERRAL_VTBL_COMPLETE, vbLong)
    End If
    If pDeferral <> 0 Then Call ComRelease(pDeferral)
    If pArgs <> 0 Then Call ComRelease(pArgs)

    If idx <> lastIdx Then
        This.deferredWebResource.items(idx) = This.deferredWebResource.items(lastIdx)
    End If
    This.deferredWebResource.items(lastIdx).requestId = 0
    This.deferredWebResource.items(lastIdx).pArgs = 0
    This.deferredWebResource.items(lastIdx).pDeferral = 0
    This.deferredWebResource.items(lastIdx).createdAt = 0#
    This.deferredWebResource.count = lastIdx
End Sub

Private Sub DeferredWebResource_ClearAll(ByVal completeDeferrals As Boolean)
    Dim i As Long
    For i = This.deferredWebResource.count - 1 To 0 Step -1
        Call DeferredWebResource_CompleteAndRemoveAt(i, completeDeferrals)
    Next i
    This.deferredWebResource.count = 0
    Erase This.deferredWebResource.items
End Sub

Private Sub WebResource_EnsureEventHandler()
    Dim hr As Long
    If This.com.pWebReqHandlerObj <> 0 Then Exit Sub
    If This.com.ppWebView2 = 0 Then Err.Raise 5, "stdWebView::WebResource_EnsureEventHandler", "WebView not ready"

    This.com.pWebReqVTable = BuildHandlerVTable(6, This.com.webReqThunk, 3)
    This.com.pWebReqHandlerObj = GlobalAlloc(GPTR, PTR_SZ)
    If This.com.pWebReqHandlerObj = 0 Then Err.Raise 5, "stdWebView::WebResource_EnsureEventHandler", "Failed allocating WebResourceRequested handler object"
    Call CopyMemory(ByVal This.com.pWebReqHandlerObj, This.com.pWebReqVTable, PTR_SZ)
    This.com.webReqTokenLow = 0
    This.com.webReqTokenHigh = 0
    hr = CLng(CallVT(This.com.ppWebView2, COREWEBVIEW2_VTBL_ADD_WEBRESOURCEREQUESTED, vbLong, This.com.pWebReqHandlerObj, VarPtr(This.com.webReqTokenLow)))
    If hr < 0 Then
        Err.Raise hr, "stdWebView::WebResource_EnsureEventHandler", "add_WebResourceRequested failed (" & CStr(hr) & ")"
    End If
End Sub

Private Sub WebResource_RemoveEventHandler()
    If This.com.ppWebView2 = 0 Then Exit Sub
    If This.com.webReqTokenLow = 0 And This.com.webReqTokenHigh = 0 Then Exit Sub
    Call CallVTQuiet(This.com.ppWebView2, COREWEBVIEW2_VTBL_REMOVE_WEBRESOURCEREQUESTED, vbLong, VarPtr(This.com.webReqTokenLow))
    This.com.webReqTokenLow = 0
    This.com.webReqTokenHigh = 0
End Sub

Private Sub WebMessage_EnsureEventHandler()
    Dim hr As Long
    If This.com.pWebMsgHandlerObj <> 0 Then Exit Sub
    If This.com.ppWebView2 = 0 Then Err.Raise 5, "stdWebView::WebMessage_EnsureEventHandler", "WebView not ready"

    This.com.pWebMsgVTable = BuildHandlerVTable(7, This.com.webMsgThunk, 3)
    This.com.pWebMsgHandlerObj = GlobalAlloc(GPTR, PTR_SZ)
    If This.com.pWebMsgHandlerObj = 0 Then Err.Raise 5, "stdWebView::WebMessage_EnsureEventHandler", "Failed allocating WebMessageReceived handler object"
    Call CopyMemory(ByVal This.com.pWebMsgHandlerObj, This.com.pWebMsgVTable, PTR_SZ)
    This.com.webMsgTokenLow = 0
    This.com.webMsgTokenHigh = 0
    hr = CLng(CallVT(This.com.ppWebView2, COREWEBVIEW2_VTBL_ADD_WEBMESSAGERECEIVED, vbLong, This.com.pWebMsgHandlerObj, VarPtr(This.com.webMsgTokenLow)))
    If hr < 0 Then
        Err.Raise hr, "stdWebView::WebMessage_EnsureEventHandler", "add_WebMessageReceived failed (" & CStr(hr) & ")"
    End If
End Sub

Private Sub WebMessage_RemoveEventHandler()
    If This.com.ppWebView2 = 0 Then Exit Sub
    If This.com.webMsgTokenLow = 0 And This.com.webMsgTokenHigh = 0 Then Exit Sub
    Call CallVTQuiet(This.com.ppWebView2, COREWEBVIEW2_VTBL_REMOVE_WEBMESSAGERECEIVED, vbLong, VarPtr(This.com.webMsgTokenLow))
    This.com.webMsgTokenLow = 0
    This.com.webMsgTokenHigh = 0
End Sub

Private Function WebResource_FilterMatches(ByRef filterItem As TWebResourceFilterItem, ByVal uri As String, ByVal context As Long) As Boolean
    Dim u As String
    Dim p As String

    If filterItem.context <> COREWEBVIEW2_WEB_RESOURCE_CONTEXT_ALL Then
        If filterItem.context <> context Then Exit Function
    End If

    p = LCase$(Trim$(filterItem.filterString))
    u = LCase$(uri)
    If LenB(p) = 0 Then Exit Function
    If p = "*" Then
        WebResource_FilterMatches = True
        Exit Function
    End If

    On Error Resume Next
    WebResource_FilterMatches = (u Like p)
    If Err.Number <> 0 Then
        Err.Clear
        WebResource_FilterMatches = (u = p)
    End If
    On Error GoTo 0
End Function

Private Function WebResource_NewDictionary() As Object
    Set WebResource_NewDictionary = CreateObject("Scripting.Dictionary")
End Function

Private Function WebResource_BuildEnvelope(ByVal uri As String, ByVal method As String, ByVal resourceContext As Long, ByVal pHeaders As LongPtr, ByVal requestId As Long) As Object
    Dim requestHeaders As Object
    Dim setHeaders As Object
    Dim envelope As Object

    Set envelope = WebResource_NewDictionary()
    Set requestHeaders = WebResource_ReadHeadersAsDictionary(pHeaders)
    Set setHeaders = WebResource_NewDictionary()

    envelope("uri") = uri
    envelope("method") = method
    envelope("resourceContext") = resourceContext
    envelope("requestId") = requestId
    Set envelope("requestHeaders") = requestHeaders
    Set envelope("setRequestHeaders") = setHeaders
    envelope("removeRequestHeaders") = Empty
    envelope("defer") = False
    envelope("respond") = False
    envelope("responseStatusCode") = 200
    envelope("responseReasonPhrase") = "OK"
    envelope("responseHeaders") = vbNullString
    envelope("responseBody") = Empty
    Set WebResource_BuildEnvelope = envelope
End Function

Private Function WebResource_ReadHeadersAsDictionary(ByVal pHeaders As LongPtr) As Object
    Dim result As Object
    Dim pIterator As LongPtr
    Dim hasCurrent As Long
    Dim pName As LongPtr
    Dim pValue As LongPtr
    Dim name As String
    Dim value As String

    Set result = WebResource_NewDictionary()
    If pHeaders = 0 Then
        Set WebResource_ReadHeadersAsDictionary = result
        Exit Function
    End If

    Call CallVT(pHeaders, COREWEBVIEW2HTTPREQUESTHEADERS_VTBL_GETITERATOR, vbLong, VarPtr(pIterator))
    If pIterator = 0 Then
        Set WebResource_ReadHeadersAsDictionary = result
        Exit Function
    End If

    On Error GoTo cleanup
    Call CallVT(pIterator, COREWEBVIEW2HTTPHEADERSITERATOR_VTBL_GETHASCURRENTHEADER, vbLong, VarPtr(hasCurrent))
    Do While hasCurrent <> 0
        pName = 0
        pValue = 0
        Call CallVT(pIterator, COREWEBVIEW2HTTPHEADERSITERATOR_VTBL_GETCURRENTHEADER, vbLong, VarPtr(pName), VarPtr(pValue))
        name = PtrToStrW(pName)
        value = PtrToStrW(pValue)
        If pName <> 0 Then Call CoTaskMemFree(pName)
        If pValue <> 0 Then Call CoTaskMemFree(pValue)
        If LenB(name) > 0 Then
            If result.Exists(name) Then
                result(name) = CStr(result(name)) & ", " & value
            Else
                result(name) = value
            End If
        End If
        Call CallVT(pIterator, COREWEBVIEW2HTTPHEADERSITERATOR_VTBL_MOVENEXT, vbLong, VarPtr(hasCurrent))
    Loop

cleanup:
    If pIterator <> 0 Then Call ComRelease(pIterator)
    Set WebResource_ReadHeadersAsDictionary = result
End Function

Private Sub WebResource_ApplyHeaderChanges(ByVal envelope As Object, ByVal pHeaders As LongPtr)
    Dim setHeaders As Object
    Dim removeHeaders As Variant
    Dim key As Variant
    Dim headerName As String
    Dim headerValue As String
    Dim i As Long
    Dim hr As Long

    If pHeaders = 0 Then Exit Sub
    If envelope Is Nothing Then Exit Sub

    If envelope.Exists("setRequestHeaders") Then
        If IsObject(envelope("setRequestHeaders")) Then
            Set setHeaders = envelope("setRequestHeaders")
            If Not setHeaders Is Nothing Then
                For Each key In setHeaders.Keys
                    headerName = CStr(key)
                    headerValue = CStr(setHeaders(key))
                    If LenB(headerName) > 0 Then
                        hr = CLng(CallVT(pHeaders, COREWEBVIEW2HTTPREQUESTHEADERS_VTBL_SETHEADER, vbLong, StrPtr(headerName), StrPtr(headerValue)))
                        If hr < 0 Then
                            Err.Raise hr, "stdWebView::WebResource_ApplyHeaderChanges", "SetHeader failed for '" & headerName & "' (" & CStr(hr) & ")"
                        End If
                    End If
                Next key
            End If
        End If
    End If

    If envelope.Exists("removeRequestHeaders") Then
        removeHeaders = envelope("removeRequestHeaders")
        If IsArray(removeHeaders) Then
            For i = LBound(removeHeaders) To UBound(removeHeaders)
                headerName = CStr(removeHeaders(i))
                If LenB(headerName) > 0 Then
                    hr = CLng(CallVT(pHeaders, COREWEBVIEW2HTTPREQUESTHEADERS_VTBL_REMOVEHEADER, vbLong, StrPtr(headerName)))
                    If hr < 0 Then
                        Err.Raise hr, "stdWebView::WebResource_ApplyHeaderChanges", "RemoveHeader failed for '" & headerName & "' (" & CStr(hr) & ")"
                    End If
                End If
            Next i
        ElseIf IsObject(removeHeaders) Then
            Dim objHeaders As Object
            Set objHeaders = removeHeaders
            If TypeName(objHeaders) = "Collection" Then
                For i = 1 To objHeaders.count
                    headerName = CStr(objHeaders(i))
                    If LenB(headerName) > 0 Then
                        hr = CLng(CallVT(pHeaders, COREWEBVIEW2HTTPREQUESTHEADERS_VTBL_REMOVEHEADER, vbLong, StrPtr(headerName)))
                        If hr < 0 Then
                            Err.Raise hr, "stdWebView::WebResource_ApplyHeaderChanges", "RemoveHeader failed for '" & headerName & "' (" & CStr(hr) & ")"
                        End If
                    End If
                Next i
            Else
                For Each key In objHeaders.Keys
                    headerName = CStr(key)
                    If LenB(headerName) > 0 Then
                        hr = CLng(CallVT(pHeaders, COREWEBVIEW2HTTPREQUESTHEADERS_VTBL_REMOVEHEADER, vbLong, StrPtr(headerName)))
                        If hr < 0 Then
                            Err.Raise hr, "stdWebView::WebResource_ApplyHeaderChanges", "RemoveHeader failed for '" & headerName & "' (" & CStr(hr) & ")"
                        End If
                    End If
                Next key
            End If
        ElseIf VarType(removeHeaders) = vbString Then
            headerName = CStr(removeHeaders)
            If LenB(headerName) > 0 Then
                hr = CLng(CallVT(pHeaders, COREWEBVIEW2HTTPREQUESTHEADERS_VTBL_REMOVEHEADER, vbLong, StrPtr(headerName)))
                If hr < 0 Then
                    Err.Raise hr, "stdWebView::WebResource_ApplyHeaderChanges", "RemoveHeader failed for '" & headerName & "' (" & CStr(hr) & ")"
                End If
            End If
        End If
    End If
End Sub

Private Function WebResource_ShouldDefer(ByVal envelope As Object) As Boolean
    If envelope Is Nothing Then Exit Function
    If Not envelope.Exists("defer") Then Exit Function
    On Error Resume Next
    WebResource_ShouldDefer = CBool(envelope("defer"))
    If Err.Number <> 0 Then
        Err.Clear
        WebResource_ShouldDefer = False
    End If
    On Error GoTo 0
End Function

Private Function WebResource_ShouldRespond(ByVal envelope As Object) As Boolean
    If envelope Is Nothing Then Exit Function
    If Not envelope.Exists("respond") Then Exit Function
    On Error Resume Next
    WebResource_ShouldRespond = CBool(envelope("respond"))
    If Err.Number <> 0 Then
        Err.Clear
        WebResource_ShouldRespond = False
    End If
    On Error GoTo 0
End Function

Private Function WebResource_CreateResponse(ByVal envelope As Object) As LongPtr
    Dim statusCode As Long
    Dim reasonPhrase As String
    Dim headersRaw As String
    Dim body As Variant
    Dim pStream As LongPtr
    Dim pResponse As LongPtr
    Dim hr As Long
    Dim pHeadersRaw As LongPtr

    statusCode = 200
    reasonPhrase = "OK"
    headersRaw = vbNullString

    If envelope.Exists("responseStatusCode") Then statusCode = CLng(envelope("responseStatusCode"))
    If envelope.Exists("responseReasonPhrase") Then reasonPhrase = CStr(envelope("responseReasonPhrase"))
    If envelope.Exists("responseHeaders") Then headersRaw = CStr(envelope("responseHeaders"))
    If envelope.Exists("responseBody") Then body = envelope("responseBody") Else body = Empty
    If LenB(headersRaw) > 0 Then pHeadersRaw = StrPtr(headersRaw) Else pHeadersRaw = 0

    pStream = WebResource_CreateBodyStream(body)
    hr = CallVT(This.com.pEnvironment, COREWEBVIEW2ENVIRONMENT_VTBL_CREATEWEBRESOURCERESPONSE, vbLong, pStream, statusCode, StrPtr(reasonPhrase), pHeadersRaw, VarPtr(pResponse))
    If pStream <> 0 Then Call ComRelease(pStream)
    If hr < 0 Then
        WebResource_CreateResponse = 0
        Exit Function
    End If
    WebResource_CreateResponse = pResponse
End Function

Private Function WebResource_CreateBodyStream(ByVal body As Variant) As LongPtr
    Dim bytes() As Byte
    Dim cb As Long
    Dim hMem As LongPtr
    Dim pMem As LongPtr
    Dim pStream As LongPtr
    Dim hr As Long

    If IsEmpty(body) Then Exit Function
    If IsNull(body) Then Exit Function

    On Error Resume Next
    If IsArray(body) Then
        bytes = body
    Else
        bytes = StrConv(CStr(body), vbFromUnicode)
    End If
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo fail

    cb = 0
    On Error Resume Next
    cb = UBound(bytes) - LBound(bytes) + 1
    On Error GoTo fail
    If cb <= 0 Then Exit Function

    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, cb)
    If hMem = 0 Then Exit Function
    pMem = GlobalLock(hMem)
    If pMem = 0 Then GoTo fail

    Call CopyMemory(ByVal pMem, bytes(LBound(bytes)), cb)
    Call GlobalUnlock(hMem)

    hr = CreateStreamOnHGlobal(hMem, 1, pStream)
    If hr <> 0 Then GoTo fail
    WebResource_CreateBodyStream = pStream
    Exit Function

fail:
    On Error Resume Next
    If pMem <> 0 Then Call GlobalUnlock(hMem)
    If hMem <> 0 Then Call GlobalFree(hMem)
    WebResource_CreateBodyStream = 0
End Function

Private Function NextScriptRequestId() As Long
    This.script.nextRequestId = This.script.nextRequestId + 1
    If This.script.nextRequestId <= 0 Then This.script.nextRequestId = 1
    NextScriptRequestId = This.script.nextRequestId
End Function

Private Sub Scripts_EnsureCapacity()
    Dim cap As Long
    If This.scripts.scriptCount < 0 Then This.scripts.scriptCount = 0
    cap = Scripts_Capacity()
    If This.scripts.scriptCount < cap Then Exit Sub
    cap = cap + 4
    ReDim Preserve This.scripts.mScripts(0 To cap - 1)
End Sub

Private Function Scripts_Capacity() As Long
    On Error GoTo noArray
    Scripts_Capacity = UBound(This.scripts.mScripts) - LBound(This.scripts.mScripts) + 1
    Exit Function
noArray:
    Scripts_Capacity = 0
End Function

Private Function Scripts_AddScript( _
    ByVal requestId As Long, _
    ByVal callback As stdICallable, _
    ByVal isSyncWait As Boolean, _
    ByVal pHandlerObj As LongPtr, _
    ByVal pVTable As LongPtr, _
    ByRef thunks As THandlerThunks) As Long

    Dim idx As Long
    Call Scripts_EnsureCapacity
    idx = This.scripts.scriptCount
    This.scripts.mScripts(idx).requestId = requestId
    This.scripts.mScripts(idx).isSyncWait = isSyncWait
    This.scripts.mScripts(idx).isCompleted = False
    Set This.scripts.mScripts(idx).callback = callback
    This.scripts.mScripts(idx).pHandlerObj = pHandlerObj
    This.scripts.mScripts(idx).pVTable = pVTable
    This.scripts.mScripts(idx).thunks = thunks
    This.scripts.scriptCount = This.scripts.scriptCount + 1
    Scripts_AddScript = idx
End Function

Private Function Scripts_FindScriptIndex(ByVal requestId As Long) As Long
    Dim i As Long
    For i = 0 To This.scripts.scriptCount - 1
        If This.scripts.mScripts(i).requestId = requestId Then
            Scripts_FindScriptIndex = i
            Exit Function
        End If
    Next i
    Scripts_FindScriptIndex = -1
End Function

Private Sub Scripts_RemoveScriptAt(ByVal idx As Long)
    Dim lastIdx As Long
    lastIdx = This.scripts.scriptCount - 1
    If idx < 0 Or idx > lastIdx Then Exit Sub
    If idx <> lastIdx Then
        This.scripts.mScripts(idx).requestId = This.scripts.mScripts(lastIdx).requestId
        This.scripts.mScripts(idx).isSyncWait = This.scripts.mScripts(lastIdx).isSyncWait
        This.scripts.mScripts(idx).isCompleted = This.scripts.mScripts(lastIdx).isCompleted
        Set This.scripts.mScripts(idx).callback = This.scripts.mScripts(lastIdx).callback
        This.scripts.mScripts(idx).pHandlerObj = This.scripts.mScripts(lastIdx).pHandlerObj
        This.scripts.mScripts(idx).pVTable = This.scripts.mScripts(lastIdx).pVTable
        This.scripts.mScripts(idx).thunks = This.scripts.mScripts(lastIdx).thunks
    End If
    This.scripts.mScripts(lastIdx).requestId = 0
    This.scripts.mScripts(lastIdx).isSyncWait = False
    This.scripts.mScripts(lastIdx).isCompleted = False
    Set This.scripts.mScripts(lastIdx).callback = Nothing
    This.scripts.mScripts(lastIdx).pHandlerObj = 0
    This.scripts.mScripts(lastIdx).pVTable = 0
    Call EraseThunkStruct(This.scripts.mScripts(lastIdx).thunks)
    This.scripts.scriptCount = lastIdx
End Sub

Private Sub Scripts_FreeScriptResources(ByRef scriptItem As TScriptItem)
    If scriptItem.pHandlerObj <> 0 Then Call GlobalFree(scriptItem.pHandlerObj)
    If scriptItem.pVTable <> 0 Then Call GlobalFree(scriptItem.pVTable)
    Call FreeHandlerThunks(scriptItem.thunks)
End Sub

Private Sub Scripts_ClearAll()
    Dim i As Long
    For i = 0 To This.scripts.scriptCount - 1
        Call Scripts_FreeScriptResources(This.scripts.mScripts(i))
        Set This.scripts.mScripts(i).callback = Nothing
    Next i
    This.scripts.scriptCount = 0
    Erase This.scripts.mScripts
End Sub

Private Sub Scripts_ClearCompleted()
    Dim i As Long
    Dim pHandlerObj As LongPtr
    Dim pVTable As LongPtr
    Dim thunks As THandlerThunks
    For i = This.scripts.scriptCount - 1 To 0 Step -1
        If This.scripts.mScripts(i).isCompleted Then
            pHandlerObj = This.scripts.mScripts(i).pHandlerObj
            pVTable = This.scripts.mScripts(i).pVTable
            thunks = This.scripts.mScripts(i).thunks
            Call Scripts_RemoveScriptAt(i)
            If pHandlerObj <> 0 Then Call GlobalFree(pHandlerObj)
            If pVTable <> 0 Then Call GlobalFree(pVTable)
            Call FreeHandlerThunks(thunks)
        End If
    Next i
End Sub

Private Function UserFormToHwnd(ByVal uf As Object) As LongPtr
    Dim hr As Long
    Dim formHwnd As LongPtr
    Dim childHwnd As LongPtr

    hr = IUnknown_GetWindow(uf, VarPtr(formHwnd))
    If hr <> 0 Or formHwnd = 0 Then
        Err.Raise 5, "stdWebView::UserFormToHwnd", "Could not get UserForm window handle (IOleWindow)."
    End If

    childHwnd = GetWindow(formHwnd, GW_CHILD)
    If childHwnd = 0 Then
        Err.Raise 5, "stdWebView::UserFormToHwnd", "Could not find a child HWND for the UserForm."
    End If

    UserFormToHwnd = childHwnd
End Function

Private Function FrameToHwnd(ByVal frm As Object) As LongPtr
    Dim hr As Long
    Dim hwnd As LongPtr
    hwnd = 0
    hr = IUnknown_GetWindow(frm, VarPtr(hwnd))
    If hr = 0 And hwnd <> 0 Then
        FrameToHwnd = hwnd
        Exit Function
    End If
    Dim pHwnd As LongPtr
    hr = IUnknown_GetWindow(frm.Parent, VarPtr(pHwnd))
    If hr <> 0 Or pHwnd = 0 Then
        Err.Raise 5, "stdWebView::FrameToHwnd", "Could not get parent window (IOleWindow)."
    End If
    FrameToHwnd = FindChildMatchingFrame(pHwnd, frm)
    If FrameToHwnd = 0 Then
        Err.Raise 5, "stdWebView::FrameToHwnd", "Could not match frame to a child HWND. Use CreateFromHwnd with an explicit host window."
    End If
End Function

Private Function FindChildMatchingFrame(ByVal parentHwnd As LongPtr, ByVal frm As Object) As LongPtr
    Dim pt As POINTAPI
    pt.x = PtsToPx(CLng(frm.Left), False)
    pt.y = PtsToPx(CLng(frm.Top), True)
    Call ClientToScreen(parentHwnd, pt)
    Dim ex As Long
    Dim ey As Long
    Dim ew As Long
    Dim eh As Long
    ex = pt.x
    ey = pt.y
    ew = PtsToPx(CLng(frm.Width), False)
    eh = PtsToPx(CLng(frm.Height), True)
    Dim child As LongPtr
    Dim r As RECT
    Dim best As LongPtr
    Dim dMin As Double
    best = 0
    dMin = 1E+30
    child = GetWindow(parentHwnd, GW_CHILD)
    Do While child <> 0
        Call GetWindowRect(child, r)
        Dim w As Long
        Dim h As Long
        w = r.Right - r.Left
        h = r.Bottom - r.Top
        If Abs(w - ew) < 100 And Abs(h - eh) < 100 Then
            Dim cx As Long
            Dim cy As Long
            cx = (r.Left + r.Right) \ 2
            cy = (r.Top + r.Bottom) \ 2
            Dim d As Double
            d = Abs(cx - (ex + ew \ 2)) + Abs(cy - (ey + eh \ 2))
            If d < dMin Then
                dMin = d
                best = child
            End If
        End If
        child = GetWindow(child, GW_HWNDNEXT)
    Loop
    FindChildMatchingFrame = best
End Function

Private Function PtsToPx(ByVal pts As Long, ByVal vertical As Boolean) As Long
    Dim hdc As LongPtr
    Dim dpi As Long
    hdc = GetDC(0)
    If vertical Then
        dpi = GetDeviceCaps(hdc, LOGPIXELSY)
    Else
        dpi = GetDeviceCaps(hdc, LOGPIXELSX)
    End If
    Call ReleaseDC(0, hdc)
    PtsToPx = CLng(CDbl(pts) * (CDbl(dpi) / 72#))
End Function

Private Sub SpinWaitMs(ByVal milliseconds As Long)
    Dim waitStart As Double
    waitStart = Timer
    Do
        DoEvents
        If TimerElapsedMs(waitStart, Timer) >= CDbl(milliseconds) Then Exit Do
    Loop
End Sub

Private Function TimerElapsedMs(ByVal startedAt As Double, ByVal nowAt As Double) As Double
    If nowAt >= startedAt Then
        TimerElapsedMs = (nowAt - startedAt) * 1000#
    Else
        TimerElapsedMs = ((86400# - startedAt) + nowAt) * 1000#
    End If
End Function

Private Function JsonUnquoteString(ByVal s As String) As String
    Dim t As String
    Dim i As Long
    Dim c As String
    Dim out As String
    t = Trim$(s)
    If Len(t) < 2 Then
        JsonUnquoteString = t
        Exit Function
    End If
    If Left$(t, 1) <> """" Or Right$(t, 1) <> """" Then
        JsonUnquoteString = t
        Exit Function
    End If
    t = Mid$(t, 2, Len(t) - 2)
    i = 1
    Do While i <= Len(t)
        c = Mid$(t, i, 1)
        If c = "\" And i < Len(t) Then
            i = i + 1
            Select Case Mid$(t, i, 1)
                Case """": out = out & """"
                Case "\": out = out & "\"
                Case "n": out = out & vbLf
                Case "r": out = out & vbCr
                Case "t": out = out & vbTab
                Case Else: out = out & Mid$(t, i, 1)
            End Select
        Else
            out = out & c
        End If
        i = i + 1
    Loop
    JsonUnquoteString = out
End Function

Private Function JsonQuoteString(ByVal value As String) As String
    Dim s As String
    s = value
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    JsonQuoteString = """" & s & """"
End Function

Private Sub ScriptSyncBridge_Reset()
    This.script.syncBridgeRequestId = vbNullString
    This.script.syncBridgeDone = False
    This.script.syncBridgeOk = False
    This.script.syncBridgePayload = vbNullString
End Sub

Private Function ScriptSyncBridge_TryConsumeMessage(ByVal messageText As String) As Boolean
    Dim prefixLen As Long
    Dim p1 As Long
    Dim p2 As Long
    Dim token As String
    Dim requestId As String
    Dim okFlag As String
    Dim payload As String

    prefixLen = Len(SCRIPT_SYNC_BRIDGE_PREFIX)
    p1 = InStr(1, messageText, "|", vbBinaryCompare)
    If p1 <= prefixLen Then Exit Function

    token = Left$(messageText, p1 - 1)
    If Left$(token, prefixLen) <> SCRIPT_SYNC_BRIDGE_PREFIX Then Exit Function
    requestId = Mid$(token, prefixLen + 1)
    If requestId <> This.script.syncBridgeRequestId Then Exit Function

    p2 = InStr(p1 + 1, messageText, "|", vbBinaryCompare)
    If p2 <= p1 Then Exit Function

    okFlag = Mid$(messageText, p1 + 1, p2 - p1 - 1)
    payload = Mid$(messageText, p2 + 1)
    If LenB(payload) = 0 Then payload = "null"

    This.script.syncBridgeOk = (okFlag = "1")
    This.script.syncBridgePayload = payload
    This.script.syncBridgeDone = True
    ScriptSyncBridge_TryConsumeMessage = True
End Function

Private Function BuildSyncBridgeScript(ByVal userScript As String, ByVal bridgeRequestId As String) As String
    Dim scriptLiteral As String
    Dim tokenLiteral As String

    scriptLiteral = JsonQuoteString(userScript)
    tokenLiteral = JsonQuoteString(SCRIPT_SYNC_BRIDGE_PREFIX & bridgeRequestId)

    BuildSyncBridgeScript = "(function(){" _
        & "const __stdwvToken=" & tokenLiteral & ";" _
        & "const __stdwvSource=" & scriptLiteral & ";" _
        & "const __stdwvPost=function(ok,payload){try{if(window.chrome&&chrome.webview&&chrome.webview.postMessage){chrome.webview.postMessage(__stdwvToken+'|'+(ok?'1':'0')+'|'+payload);}}catch(_){}};" _
        & "const __stdwvToJson=function(v){try{const s=JSON.stringify(v);return (s===undefined)?'null':s;}catch(_){return 'null';}};" _
        & "try{" _
            & "const __stdwvValue=(0,eval)(__stdwvSource);" _
            & "if(__stdwvValue&&typeof __stdwvValue.then==='function'){" _
                & "__stdwvValue.then(function(v){__stdwvPost(true,__stdwvToJson(v));},function(e){const m=(e&&e.message)?String(e.message):String(e);__stdwvPost(false,JSON.stringify(m));});" _
            & "}else{" _
                & "__stdwvPost(true,__stdwvToJson(__stdwvValue));" _
            & "}" _
        & "}catch(e){" _
            & "const m=(e&&e.message)?String(e.message):String(e);__stdwvPost(false,JSON.stringify(m));" _
        & "}" _
        & "return null;" _
    & "})()"
End Function

' ============================================================================
' Resolve a COM method address from the class vtable (IUnknown + IDispatch + Public)
' ============================================================================

Private Function GetComProcAddress( _
    ByVal o As Object, _
    ByVal publicMemberIndex As Long) As LongPtr

    Dim pObj As LongPtr
    Dim pVTable As LongPtr
    Dim cbPtr As LongPtr
    Dim pEntry As LongPtr

    pObj = ObjPtr(o)
    If pObj = 0 Then Exit Function

    cbPtr = LenB(pObj)

    CopyMemoryByAddr pVTable, pObj, cbPtr

    ' IUnknown = 3 entries, IDispatch = 4 entries, then Public members
    pEntry = pVTable + ((7 + publicMemberIndex) * cbPtr)

    CopyMemoryByAddr GetComProcAddress, pEntry, cbPtr
End Function

' ============================================================================
' Thunk builder (from AddressOfThunkForComProc / CMonitors pattern)
' Converts a COM instance method:
'   HRESULT Method(this, arg1, arg2, ..., retval)
' into a stdcall callback:
'   Long Callback(arg1, arg2, ...)
' 1 to 4 args; pointer-sized args; return Long in hidden retval slot.
' ============================================================================

Private Function AddressOfThunkForComProc( _
    ByVal o As Object, _
    ByVal ComProcAddress As LongPtr, _
    ByVal ComProcArgCount As Long, _
    ByRef iThunkSize As LongPtr, _
    Optional ByVal AddIDEsafety As Boolean = True) As LongPtr

    Dim lEbMode As LongPtr
    If AddIDEsafety Then
        On Error Resume Next
        Debug.Print 1 / 0
        If Err Then
            Err.Clear: On Error GoTo 0
            Dim hVBE As LongPtr
            hVBE = GetModuleHandle("vbe7.dll")
            If hVBE <> 0 Then
                lEbMode = GetProcAddress(hVBE, "EbMode")
            End If
        End If
        On Error GoTo 0
    End If

#If Win64 Then
    Dim pThis As LongPtr: pThis = ObjPtr(o)
    If pThis = 0 Or ComProcAddress = 0 Then Exit Function
    If ComProcArgCount < 1 Or ComProcArgCount > 4 Then Exit Function

    Dim p As Long: p = 0
    Dim bb() As Byte: ReDim bb(0 To 159)

    EmitBytes bb, p, &H48, &H83, &HEC, &H58
    EmitBytes bb, p, &H48, &H89, &H4C, &H24, &H38
    EmitBytes bb, p, &H48, &H89, &H54, &H24, &H40
    EmitBytes bb, p, &H4C, &H89, &H44, &H24, &H48
    EmitBytes bb, p, &H4C, &H89, &H4C, &H24, &H50

    EmitBytes bb, p, &H48, &HB8
    EmitPtr bb, p, lEbMode
    EmitBytes bb, p, &H48, &H85, &HC0
    EmitBytes bb, p, &H74, &H7
    EmitBytes bb, p, &HFF, &HD0
    EmitBytes bb, p, &H83, &HF8, &H1
    Dim jneAt As Long: jneAt = p + 1
    EmitBytes bb, p, &H75, &H0

    Select Case ComProcArgCount
        Case 1
            EmitBytes bb, p, &H4C, &H8D, &H44, &H24, &H30
            EmitBytes bb, p, &H48, &H8B, &H54, &H24, &H38
        Case 2
            EmitBytes bb, p, &H4C, &H8D, &H4C, &H24, &H30
            EmitBytes bb, p, &H4C, &H8B, &H44, &H24, &H40
            EmitBytes bb, p, &H48, &H8B, &H54, &H24, &H38
        Case 3
            EmitBytes bb, p, &H4C, &H8D, &H54, &H24, &H30
            EmitBytes bb, p, &H4C, &H89, &H54, &H24, &H20
            EmitBytes bb, p, &H4C, &H8B, &H4C, &H24, &H48
            EmitBytes bb, p, &H4C, &H8B, &H44, &H24, &H40
            EmitBytes bb, p, &H48, &H8B, &H54, &H24, &H38
        Case 4
            EmitBytes bb, p, &H4C, &H8B, &H54, &H24, &H50
            EmitBytes bb, p, &H4C, &H89, &H54, &H24, &H20
            EmitBytes bb, p, &H4C, &H8D, &H54, &H24, &H30
            EmitBytes bb, p, &H4C, &H89, &H54, &H24, &H28
            EmitBytes bb, p, &H4C, &H8B, &H4C, &H24, &H48
            EmitBytes bb, p, &H4C, &H8B, &H44, &H24, &H40
            EmitBytes bb, p, &H48, &H8B, &H54, &H24, &H38
    End Select

    EmitBytes bb, p, &H48, &HB9
    EmitPtr bb, p, pThis

    EmitBytes bb, p, &H48, &HB8
    EmitPtr bb, p, ComProcAddress

    EmitBytes bb, p, &HFF, &HD0

    EmitBytes bb, p, &H8B, &H44, &H24, &H30
    EmitBytes bb, p, &H48, &H83, &HC4, &H58
    EmitBytes bb, p, &HC3

    Dim retZeroAt As Long: retZeroAt = p
    EmitBytes bb, p, &H31, &HC0
    EmitBytes bb, p, &H48, &H83, &HC4, &H58
    EmitBytes bb, p, &HC3

    bb(jneAt) = CByte(retZeroAt - (jneAt + 1))

    iThunkSize = p
    ReDim Preserve bb(0 To p - 1)

    AddressOfThunkForComProc = VirtualAlloc( _
        0, _
        iThunkSize, _
        MEM_COMMIT Or MEM_RESERVE, _
        PAGE_EXECUTE_READWRITE)

    If AddressOfThunkForComProc = 0 Then Exit Function

    CopyMemoryByPtr AddressOfThunkForComProc, bb(0), iThunkSize

#Else
    If ComProcAddress = 0 Or ObjPtr(o) = 0 Then Exit Function
    If ComProcArgCount < 1 Or ComProcArgCount > 4 Then Exit Function

    Dim p32 As Long
    Dim bb32() As Byte: ReDim bb32(0 To 127)

    EmitBytes bb32, p32, &H55
    EmitBytes bb32, p32, &H89, &HE5
    EmitBytes bb32, p32, &H83, &HEC, &H4

    EmitBytes bb32, p32, &HB8
    Dim ebModeAt32 As Long: ebModeAt32 = p32
    EmitBytes bb32, p32, &H0, &H0, &H0, &H0
    EmitBytes bb32, p32, &H85, &HC0
    EmitBytes bb32, p32, &H74, &H7
    EmitBytes bb32, p32, &HFF, &HD0
    EmitBytes bb32, p32, &H83, &HF8, &H1
    Dim jneAt32 As Long: jneAt32 = p32 + 1
    EmitBytes bb32, p32, &H75, &H0

    EmitBytes bb32, p32, &H89, &HE8
    EmitBytes bb32, p32, &H83, &HE8, &H4
    EmitBytes bb32, p32, &H50

    Dim arg4PushAt As Long: arg4PushAt = p32
    EmitBytes bb32, p32, &HFF, &H75, &H14
    Dim arg3PushAt As Long: arg3PushAt = p32
    EmitBytes bb32, p32, &HFF, &H75, &H10
    Dim arg2PushAt As Long: arg2PushAt = p32
    EmitBytes bb32, p32, &HFF, &H75, &HC
    EmitBytes bb32, p32, &HFF, &H75, &H8

    EmitBytes bb32, p32, &H68
    Dim objImmAt As Long: objImmAt = p32
    EmitBytes bb32, p32, &H55, &H55, &H55, &H55

    EmitBytes bb32, p32, &HB8
    Dim procImmAt As Long: procImmAt = p32
    EmitBytes bb32, p32, &H66, &H66, &H66, &H66
    EmitBytes bb32, p32, &HFF, &HD0

    EmitBytes bb32, p32, &H8B, &H45, &HFC
    EmitBytes bb32, p32, &H89, &HEC
    EmitBytes bb32, p32, &H5D

    EmitBytes bb32, p32, &HC2
    Dim retImmAt As Long: retImmAt = p32
    EmitBytes bb32, p32, &H8, &H0

    Dim retZeroAt32 As Long: retZeroAt32 = p32
    EmitBytes bb32, p32, &H31, &HC0
    EmitBytes bb32, p32, &H89, &HEC
    EmitBytes bb32, p32, &H5D
    EmitBytes bb32, p32, &HC2
    Dim retImmAt2 As Long: retImmAt2 = p32
    EmitBytes bb32, p32, &H8, &H0

    bb32(jneAt32) = CByte(retZeroAt32 - (jneAt32 + 1))

    iThunkSize = p32
    ReDim Preserve bb32(0 To p32 - 1)

    If ComProcArgCount < 4 Then
        bb32(arg4PushAt) = &H90
        bb32(arg4PushAt + 1) = &H90
        bb32(arg4PushAt + 2) = &H90
    End If

    If ComProcArgCount < 3 Then
        bb32(arg3PushAt) = &H90
        bb32(arg3PushAt + 1) = &H90
        bb32(arg3PushAt + 2) = &H90
    End If

    If ComProcArgCount < 2 Then
        bb32(arg2PushAt) = &H90
        bb32(arg2PushAt + 1) = &H90
        bb32(arg2PushAt + 2) = &H90
    End If

    CopyMemory bb32(ebModeAt32), lEbMode, 4&
    Dim pObj32 As Long: pObj32 = ObjPtr(o)
    CopyMemory bb32(objImmAt), pObj32, 4&
    CopyMemory bb32(procImmAt), ComProcAddress, 4&

    bb32(retImmAt) = ComProcArgCount * 4&
    bb32(retImmAt2) = ComProcArgCount * 4&

    Dim pThunk32 As LongPtr: pThunk32 = VirtualAlloc( _
        0, _
        iThunkSize, _
        MEM_COMMIT Or MEM_RESERVE, _
        PAGE_EXECUTE_READWRITE)

    If pThunk32 = 0 Then Exit Function

    CopyMemoryByPtr pThunk32, bb32(0), iThunkSize
    AddressOfThunkForComProc = pThunk32
#End If
End Function

Private Sub EmitBytes( _
    ByRef bb() As Byte, _
    ByRef p As Long, _
    ParamArray values() As Variant)

    Dim i As Long

    For i = LBound(values) To UBound(values)
        bb(p) = CByte(CLng(values(i)) And &HFF&)
        p = p + 1
    Next
End Sub

Private Sub EmitPtr( _
    ByRef bb() As Byte, _
    ByRef p As Long, _
    ByVal value As LongPtr)

    Dim cb As Long

    cb = LenB(value)
    CopyMemory bb(p), value, cb
    p = p + cb
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Call Quit
End Sub
