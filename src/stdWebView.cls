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
    nextRequestId As Long
End Type

Private Type TScripts
    mScripts() As TScriptItem
    scriptCount As Long
End Type

Private Type THostObjectState
    refs As Collection
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
End Type

Private Type TThis
    parentHwnd As LongPtr
    lastHtml As String
    init As TInitState
    script As TScriptState
    scripts As TScripts
    host As THostObjectState
    com As TComState
End Type

Private This As TThis

Private Const CC_STDCALL As Long = 4
Private Const GPTR As Long = &H40

Private Const MEM_COMMIT As Long = &H1000&
Private Const MEM_RESERVE As Long = &H2000&
Private Const MEM_RELEASE As Long = &H8000&
Private Const PAGE_EXECUTE_READWRITE As Long = &H40&
Private Const GW_CHILD As Long = 5
Private Const GW_HWNDNEXT As Long = 2
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90
' Method slots are from ICoreWebView2Vtbl in WebView2.h (IUnknown + interface members).
Private Const COREWEBVIEW2_VTBL_EXECUTESCRIPT As Long = 29
Private Const COREWEBVIEW2_VTBL_ADDHOSTOBJECTTOSCRIPT As Long = 49
Private Const COREWEBVIEW2_VTBL_REMOVEHOSTOBJECTFROMSCRIPT As Long = 50
Private Const COREWEBVIEW2_VTBL_SETVIRTUALHOSTNAMETOFOLDERMAPPING As Long = 71
Private Const COREWEBVIEW2_VTBL_CLEARVIRTUALHOSTNAMETOFOLDERMAPPING As Long = 72

Public Enum COREWEBVIEW2_HOST_RESOURCE_ACCESS_KIND
    COREWEBVIEW2_HOST_RESOURCE_ACCESS_KIND_DENY = 0
    COREWEBVIEW2_HOST_RESOURCE_ACCESS_KIND_ALLOW = 1
    COREWEBVIEW2_HOST_RESOURCE_ACCESS_KIND_DENY_CORS = 2
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
    Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As LongPtr)

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
    Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As LongPtr)

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
    buf = String$(n, vbNullChar)
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
' `GetComProcAddress(Me, 0..5)` depends on that. Do not move other Public methods
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
    If Not IsReady Then Err.Raise 5, "stdWebView::JavaScriptRunSync", "WebView not ready"
    If This.script.syncWaitRequestId <> 0 Then
        Err.Raise 5, "stdWebView::JavaScriptRunSync", "A synchronous script call is already in progress"
    End If
    Call Scripts_ClearCompleted
    This.script.syncDone = False
    This.script.result = vbNullString
    This.script.syncWaitRequestId = JavaScriptRunCore(script, Nothing, True)
    If This.script.syncWaitRequestId = 0 Then
        This.script.syncWaitRequestId = 0
        Err.Raise 5, "stdWebView::JavaScriptRunSync", "Failed to queue script execution"
    End If
    Do
        If This.script.syncDone Then Exit Do
        DoEvents
    Loop
    This.script.syncWaitRequestId = 0
    Call Scripts_ClearCompleted
    JavaScriptRunSync = This.script.result
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
    This.script.nextRequestId = 0
    This.scripts.scriptCount = 0
    Erase This.scripts.mScripts
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
