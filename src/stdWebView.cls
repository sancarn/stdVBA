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

Private Type TPendingScript
    requestId As Long
    isSyncWait As Boolean
    isCompleted As Boolean
    pHandlerObj As LongPtr
    pVTable As LongPtr
    thunks As THandlerThunks
End Type

Private Type TThis
    parentHwnd As LongPtr
    pEnvironment As LongPtr
    pController As LongPtr
    ppWebView2 As LongPtr
    userDataFolder As String
    pEnvHandlerObj As LongPtr
    pEnvVTable As LongPtr
    envThunk As THandlerThunks
    pCtrlHandlerObj As LongPtr
    pCtrlVTable As LongPtr
    ctrlThunk As THandlerThunks
    waitingInit As Boolean
    initFailed As Boolean
    lastHtml As String
    scriptResult As String
    syncScriptDone As Boolean
    syncWaitRequestId As Long
    nextRequestId As Long
End Type

Private This As TThis
Private mOnReadyCallback As stdICallable
Private mPendingScripts() As TPendingScript
Private mPendingCallbacks() As stdICallable
Private mPendingScriptCount As Long

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
' WebView2 handler vtables (IUnknown + Invoke)
' These six Public methods must remain the first Public members in this class:
' GetComProcAddress(Me, 0..5) depends on that. Do not move other Public methods
' above them without adjusting the indexes.
' ============================================================================

Public Function zzProtWebView_QI( _
    ByVal callbackThis As LongPtr, _
    ByVal riid As LongPtr, _
    ByRef ppvObject As LongPtr) As Long

    ppvObject = callbackThis
    zzProtWebView_QI = 0&
End Function

Public Function zzProtWebView_AddRef(ByVal callbackThis As LongPtr) As Long
    zzProtWebView_AddRef = 1&
End Function

Public Function zzProtWebView_Release(ByVal callbackThis As LongPtr) As Long
    zzProtWebView_Release = 1&
End Function

Public Function zzProtWebView_EnvInvoke( _
    ByVal callbackThis As LongPtr, _
    ByVal errorCode As Long, _
    ByVal pEnvironment As LongPtr) As Long

    On Error Resume Next
    Call protEnvCompleted(errorCode, pEnvironment)
    zzProtWebView_EnvInvoke = 0&
End Function

Public Function zzProtWebView_CtrlInvoke( _
    ByVal callbackThis As LongPtr, _
    ByVal errorCode As Long, _
    ByVal pController As LongPtr) As Long

    On Error Resume Next
    Call protCtrlCompleted(errorCode, pController)
    zzProtWebView_CtrlInvoke = 0&
End Function

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

'@constructor
'@param OnReady as stdICallable<(stdWebView)=>void> - Optional callback run when WebView2 controller/CoreWebView2 is ready.
Public Function CreateFromHwnd(ByVal hwnd As LongPtr, Optional ByVal OnReady As stdICallable = Nothing) As stdWebView
    Dim w As New stdWebView
    Call w.protCreate(hwnd, OnReady)
    Set CreateFromHwnd = w
End Function

'@constructor
'@param OnReady as stdICallable<(stdWebView)=>void> - Optional callback run when WebView2 controller/CoreWebView2 is ready.
Public Function CreateFromFrame(ByVal frm As Object, Optional ByVal OnReady As stdICallable = Nothing) As stdWebView
    If frm Is Nothing Then Err.Raise 5, "stdWebView::CreateFromFrame", "frm is Nothing"
    If TypeName(frm) <> "Frame" Then Err.Raise 5, "stdWebView::CreateFromFrame", "Expected MSForms.Frame"
    Set CreateFromFrame = CreateFromHwnd(FrameToHwnd(frm), OnReady)
End Function

Public Function IsReady() As Boolean
    IsReady = (This.ppWebView2 <> 0)
End Function

Public Sub Quit()
    On Error Resume Next
    If This.pController <> 0 Then
        Call CallVTQuiet(This.pController, 24, vbLong)
    End If
    Set mOnReadyCallback = Nothing
    This.pController = 0
    This.ppWebView2 = 0
    This.pEnvironment = 0
    Call FreeHandlerAllocs
End Sub

Public Sub Navigate(ByVal url As String)
    If Not IsReady Then Err.Raise 5, "stdWebView::Navigate", "WebView not ready"
    Call CallVT(This.ppWebView2, 5, vbLong, StrPtr(url))
End Sub

Public Sub Back()
    Call JavaScriptRunSync("history.back(); void 0")
End Sub

Public Sub Forward()
    Call JavaScriptRunSync("history.forward(); void 0")
End Sub

Public Property Get Html() As String
    If Not IsReady Then Err.Raise 5, "stdWebView::Html [Get]", "WebView not ready"
    Html = JsonUnquoteString(JavaScriptRunSync("(function(){try{return document.documentElement?document.documentElement.outerHTML:'';}catch(e){return '';}})()"))
End Property

Public Property Let Html(ByVal rhs As String)
    If Not IsReady Then Err.Raise 5, "stdWebView::Html [Let]", "WebView not ready"
    This.lastHtml = rhs
    Call CallVT(This.ppWebView2, 6, vbLong, StrPtr(rhs))
End Property

Public Function JavaScriptRunSync(ByVal script As String) As String
    If Not IsReady Then Err.Raise 5, "stdWebView::JavaScriptRunSync", "WebView not ready"
    If This.syncWaitRequestId <> 0 Then
        Err.Raise 5, "stdWebView::JavaScriptRunSync", "A synchronous script call is already in progress"
    End If
    Call SweepCompletedPendingScripts
    This.syncScriptDone = False
    This.scriptResult = vbNullString
    This.syncWaitRequestId = JavaScriptRunCore(script, Nothing, True)
    If This.syncWaitRequestId = 0 Then
        This.syncWaitRequestId = 0
        Err.Raise 5, "stdWebView::JavaScriptRunSync", "Failed to queue script execution"
    End If
    Do
        Debug.Print "Waiting..."
        If This.syncScriptDone Then Exit Do
        DoEvents
    Loop
    This.syncWaitRequestId = 0
    Call SweepCompletedPendingScripts
    JavaScriptRunSync = This.scriptResult
End Function

'@param callback as stdICallable<(errorCode as long, resultJson as string)=>void> - Optional callback invoked when script execution completes.
Public Sub JavaScriptRun(ByVal script As String, Optional ByVal callback As stdICallable = Nothing)
    If Not IsReady Then Err.Raise 5, "stdWebView::JavaScriptRun", "WebView not ready"
    Call SweepCompletedPendingScripts
    If JavaScriptRunCore(script, callback, False) = 0 Then
        Err.Raise 5, "stdWebView::JavaScriptRun", "Failed to queue script execution"
    End If
End Sub

Friend Sub protEnvCompleted(ByVal errorCode As Long, ByVal pEnvironment As LongPtr)
    If errorCode <> 0 Then
        This.waitingInit = False
        This.initFailed = True
        Exit Sub
    End If
    This.pEnvironment = pEnvironment
    This.pCtrlVTable = BuildHandlerVTable(4, This.ctrlThunk)
    This.pCtrlHandlerObj = GlobalAlloc(GPTR, PTR_SZ)
    Call CopyMemory(ByVal This.pCtrlHandlerObj, This.pCtrlVTable, PTR_SZ)
    Call CallVT(pEnvironment, 3, vbLong, This.parentHwnd, This.pCtrlHandlerObj)
End Sub

Friend Sub protCtrlCompleted(ByVal errorCode As Long, ByVal pController As LongPtr)
    If errorCode <> 0 Then
        This.waitingInit = False
        This.initFailed = True
        Exit Sub
    End If
    Call ComAddRef(pController)
    This.pController = pController
    Dim r As RECT
    Call GetClientRect(This.parentHwnd, r)
    Call ControllerPutBounds(0&, 0&, r.Right - r.Left, r.Bottom - r.Top)
    Call ControllerPutVisible(True)
    Call ControllerGetCoreWebView2
    Call ResizeChromeWidgets
    This.syncScriptDone = True
    This.waitingInit = False
    If Not mOnReadyCallback Is Nothing Then
        Dim cbReady As stdICallable
        Set cbReady = mOnReadyCallback
        Set mOnReadyCallback = Nothing
        On Error Resume Next
        Call cbReady.RunEx(Array(Me))
    End If
End Sub

Friend Sub protScriptCompleted(ByVal errorCode As Long, ByVal resultJson As String, ByVal requestId As Long)
    Dim idx As Long
    Dim cb As stdICallable
    Dim isSyncWait As Boolean

    idx = FindPendingScriptIndex(requestId)
    If idx < 0 Then Exit Sub

    If mPendingScripts(idx).isCompleted Then Exit Sub
    mPendingScripts(idx).isCompleted = True
    If mPendingCallbacks(idx) Is Nothing Then
        Set cb = Nothing
    Else
        Set cb = mPendingCallbacks(idx)
    End If
    Set mPendingCallbacks(idx) = Nothing
    isSyncWait = mPendingScripts(idx).isSyncWait

    If isSyncWait And This.syncWaitRequestId = requestId Then
        This.scriptResult = resultJson
        This.syncScriptDone = True
    End If

    If Not cb Is Nothing Then
        On Error Resume Next
        Call cb.RunEx(Array(errorCode, resultJson))
    End If
End Sub

Friend Sub protCreate(ByVal hwnd As LongPtr, Optional ByVal OnReady As stdICallable = Nothing)
    If hwnd = 0 Then Err.Raise 5, "stdWebView::CreateFromHwnd", "hwnd is 0"
    This.parentHwnd = hwnd
    This.waitingInit = True
    This.initFailed = False
    This.syncScriptDone = False
    This.syncWaitRequestId = 0
    This.nextRequestId = 0
    mPendingScriptCount = 0
    Erase mPendingScripts
    Erase mPendingCallbacks
    Set mOnReadyCallback = OnReady
    Randomize
    This.userDataFolder = Environ$("TEMP") & "\stdWebView_" & CStr(CLng(Timer * 10000) Xor CLng(Rnd * 1000000))
    This.pEnvVTable = BuildHandlerVTable(3, This.envThunk)
    This.pEnvHandlerObj = GlobalAlloc(GPTR, PTR_SZ)
    Call CopyMemory(ByVal This.pEnvHandlerObj, This.pEnvVTable, PTR_SZ)
    Dim hr As Long
    hr = CreateCoreWebView2EnvironmentWithOptions(0, StrPtr(This.userDataFolder), 0, This.pEnvHandlerObj)
    If hr <> 0 Then
        Call CleanupFailedInit
        Err.Raise hr, "stdWebView", "CreateCoreWebView2EnvironmentWithOptions failed (" & CStr(hr) & ")"
    End If
    Dim j As Long
    For j = 1 To 60000
        If Not This.waitingInit Then Exit For
        DoEvents
    Next j
    If This.initFailed Or This.ppWebView2 = 0 Then
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
    Call CallVT(This.pController, 6, vbLong, VarPtr(rc))
End Sub

Private Sub ControllerPutVisible(ByVal visible As Boolean)
    Call CallVT(This.pController, 4, vbLong, CLng(IIf(visible, 1, 0)))
End Sub

Private Sub ControllerGetCoreWebView2()
    Dim pp As LongPtr
    Call CallVT(This.pController, 25, vbLong, VarPtr(pp))
    This.ppWebView2 = pp
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
    Set mOnReadyCallback = Nothing
    Call FreeHandlerAllocs
End Sub

Private Sub FreeHandlerAllocs()
    On Error Resume Next
    Set mOnReadyCallback = Nothing
    Call FreeHandlerThunks(This.envThunk)
    Call FreeHandlerThunks(This.ctrlThunk)
    If This.pEnvHandlerObj <> 0 Then
        Call GlobalFree(This.pEnvHandlerObj)
        This.pEnvHandlerObj = 0
    End If
    If This.pEnvVTable <> 0 Then
        Call GlobalFree(This.pEnvVTable)
        This.pEnvVTable = 0
    End If
    If This.pCtrlHandlerObj <> 0 Then
        Call GlobalFree(This.pCtrlHandlerObj)
        This.pCtrlHandlerObj = 0
    End If
    If This.pCtrlVTable <> 0 Then
        Call GlobalFree(This.pCtrlVTable)
        This.pCtrlVTable = 0
    End If
    Call ClearAllPendingScripts
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

    idx = AddPendingScript(reqId, callback, isSyncWait, pObj, pVTable, thunks)
    Call CallVT(This.ppWebView2, 29, vbLong, StrPtr(script), pObj)
    On Error GoTo 0
    JavaScriptRunCore = reqId
    Exit Function

fail:
    If idx >= 0 Then Call RemovePendingScriptAt(idx)
    If pObj <> 0 Then Call GlobalFree(pObj)
    If pVTable <> 0 Then Call GlobalFree(pVTable)
    Call FreeHandlerThunks(thunks)
End Function

Private Function NextScriptRequestId() As Long
    This.nextRequestId = This.nextRequestId + 1
    If This.nextRequestId <= 0 Then This.nextRequestId = 1
    NextScriptRequestId = This.nextRequestId
End Function

Private Sub EnsurePendingCapacity()
    Dim cap As Long
    If mPendingScriptCount < 0 Then mPendingScriptCount = 0
    cap = PendingCapacity()
    If mPendingScriptCount < cap Then Exit Sub
    cap = cap + 4
    ReDim Preserve mPendingScripts(0 To cap - 1)
    ReDim Preserve mPendingCallbacks(0 To cap - 1)
End Sub

Private Function PendingCapacity() As Long
    On Error GoTo noArray
    PendingCapacity = UBound(mPendingScripts) - LBound(mPendingScripts) + 1
    Exit Function
noArray:
    PendingCapacity = 0
End Function

Private Function AddPendingScript( _
    ByVal requestId As Long, _
    ByVal callback As stdICallable, _
    ByVal isSyncWait As Boolean, _
    ByVal pHandlerObj As LongPtr, _
    ByVal pVTable As LongPtr, _
    ByRef thunks As THandlerThunks) As Long

    Dim idx As Long
    Call EnsurePendingCapacity
    idx = mPendingScriptCount
    mPendingScripts(idx).requestId = requestId
    mPendingScripts(idx).isSyncWait = isSyncWait
    mPendingScripts(idx).isCompleted = False
    mPendingScripts(idx).pHandlerObj = pHandlerObj
    mPendingScripts(idx).pVTable = pVTable
    mPendingScripts(idx).thunks = thunks
    Set mPendingCallbacks(idx) = callback
    mPendingScriptCount = mPendingScriptCount + 1
    AddPendingScript = idx
End Function

Private Function FindPendingScriptIndex(ByVal requestId As Long) As Long
    Dim i As Long
    For i = 0 To mPendingScriptCount - 1
        If mPendingScripts(i).requestId = requestId Then
            FindPendingScriptIndex = i
            Exit Function
        End If
    Next i
    FindPendingScriptIndex = -1
End Function

Private Sub RemovePendingScriptAt(ByVal idx As Long)
    Dim lastIdx As Long
    lastIdx = mPendingScriptCount - 1
    If idx < 0 Or idx > lastIdx Then Exit Sub
    If idx <> lastIdx Then
        mPendingScripts(idx) = mPendingScripts(lastIdx)
        Set mPendingCallbacks(idx) = mPendingCallbacks(lastIdx)
    End If
    mPendingScripts(lastIdx).requestId = 0
    mPendingScripts(lastIdx).isSyncWait = False
    mPendingScripts(lastIdx).isCompleted = False
    mPendingScripts(lastIdx).pHandlerObj = 0
    mPendingScripts(lastIdx).pVTable = 0
    Call EraseThunkStruct(mPendingScripts(lastIdx).thunks)
    Set mPendingCallbacks(lastIdx) = Nothing
    mPendingScriptCount = lastIdx
End Sub

Private Sub FreePendingScriptResources(ByRef pending As TPendingScript)
    If pending.pHandlerObj <> 0 Then Call GlobalFree(pending.pHandlerObj)
    If pending.pVTable <> 0 Then Call GlobalFree(pending.pVTable)
    Call FreeHandlerThunks(pending.thunks)
End Sub

Private Sub ClearAllPendingScripts()
    Dim i As Long
    For i = 0 To mPendingScriptCount - 1
        Call FreePendingScriptResources(mPendingScripts(i))
        Set mPendingCallbacks(i) = Nothing
    Next i
    mPendingScriptCount = 0
    Erase mPendingScripts
    Erase mPendingCallbacks
End Sub

Private Sub SweepCompletedPendingScripts()
    Dim i As Long
    Dim pending As TPendingScript
    For i = mPendingScriptCount - 1 To 0 Step -1
        If mPendingScripts(i).isCompleted Then
            pending = mPendingScripts(i)
            Call RemovePendingScriptAt(i)
            Call FreePendingScriptResources(pending)
        End If
    Next i
End Sub

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
