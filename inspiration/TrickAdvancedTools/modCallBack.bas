Attribute VB_Name = "modCallBack"
' // modCallBack.bas - intercepted functions handlers
' // © Krivous Anatoly Anatolevich (The trick), 2016

Option Explicit

Public pfnTipCompileProject     As Long ' // Original address of TipCompileProject function
Public pfnTipCompileProjectFull As Long ' //
Public pfnTipMakeExe2           As Long ' //
Public pfnTipFinishExe2         As Long ' //
Public pVBInstance              As VBE  ' // IDE instance

' // Compile project handler (F5)
Public Function TipCompileProject_user( _
                ByVal pVBProjectNative As Long) As Long
    
    On Error GoTo error_handler
    
    PauseHook pfnTipCompileProject
    
    MakeEvent pVBProjectNative, "CompBefore"
    
    TipCompileProject_user = CallByPointer(pfnTipCompileProject, vbLong, pVBProjectNative)
    
    MakeEvent pVBProjectNative, "CompAfter"
    
    ResumeHook pfnTipCompileProject
    
    Exit Function
    
error_handler:
    
    ErrorLog "modCallBack::TipCompileProject_user"
    
End Function

' // Full compile project handler (Ctrl+F5)
Public Function TipCompileProjectFull_user( _
                ByVal pVBProjectNative As Long) As Long

    On Error GoTo error_handler
    
    PauseHook pfnTipCompileProjectFull
    
    MakeEvent pVBProjectNative, "CompBefore"
    
    TipCompileProjectFull_user = CallByPointer(pfnTipCompileProjectFull, vbLong, pVBProjectNative)
    
    MakeEvent pVBProjectNative, "CompAfter"
    
    ResumeHook pfnTipCompileProjectFull
    
    Exit Function
    
error_handler:
    
    ErrorLog "modCallBack::TipCompileProjectFull_user"
    
End Function

' // Compilation to OBJ handler
Public Function TipMakeExe2_user( _
                 ByVal pVBProjectNative As Long, _
                 ByVal lUnused1 As Long, _
                 ByVal lUnused2 As Long, _
                 ByVal lUnused3 As Long, _
                 ByVal lUnused4 As Long) As Long
    Dim sOrigin As String
    Dim sEXE    As String
    Dim sNew    As String
    
    On Error GoTo error_handler
    
    PauseHook pfnTipMakeExe2
    
    ' // Get conditional arguments from project
    sOrigin = GetConditionalArguments(pVBProjectNative)
    ' // Get TAT conditional arguments
    sEXE = GetConditionalArguments(pVBProjectNative, "CondCOMP")
    
    ' // Concatenation arguments
    If Len(sOrigin) > 0 And Len(sEXE) Then
        sNew = sOrigin & ": " & sEXE
    ElseIf Len(sOrigin) > 0 Then
        sNew = sOrigin
    ElseIf Len(sEXE) > 0 Then
        sNew = sEXE
    End If
    
    ' // Set new arguments
    SetConditionalArguments pVBProjectNative, sNew
    
    MakeEvent pVBProjectNative, "BuildBefore"
    
    TipMakeExe2_user = CallByPointer(pfnTipMakeExe2, vbLong, pVBProjectNative, lUnused1, lUnused2, lUnused3, lUnused4)
    
    MakeEvent pVBProjectNative, "BuildAfter"
    
    ' // Restore arguments after compile
    SetConditionalArguments pVBProjectNative, sOrigin
    
    ResumeHook pfnTipMakeExe2

    Exit Function
    
error_handler:
    
    ErrorLog "modCallBack::TipMakeExe2_user"
           
End Function

' // Linking
Public Function TipFinishExe2_user( _
                ByVal pVBProjectNative As Long, _
                ByVal lUnused As Long) As Long

    On Error GoTo error_handler
    
    PauseHook pfnTipFinishExe2
    
    MakeEvent pVBProjectNative, "LinkBefore"
    
    TipFinishExe2_user = CallByPointer(pfnTipFinishExe2, vbLong, pVBProjectNative, lUnused)
    
    MakeEvent pVBProjectNative, "LinkAfter"
    
    ResumeHook pfnTipFinishExe2
    
    Exit Function
    
error_handler:
    
    ErrorLog "modCallBack::TipFinishExe2_user"
           
End Function

' // Get conditional argument. If the second param is mising get original conditional args
Private Function GetConditionalArguments( _
                 ByVal pVBProjectNative As Long, _
                 Optional ByRef sSection As String) As String
    Dim sConstant   As String
    Dim hr          As Long
    Dim pVBProject  As VBProject
    
    On Error GoTo error_handler
    
    Set pVBProject = GetProjectFromNativeProject(pVBProjectNative)
    If pVBProject Is Nothing Then Exit Function
    
    If Len(sSection) Then
        
        On Error Resume Next
        
        sConstant = pVBProject.ReadProperty("TAT", sSection)
        If Len(sConstant) = 0 Then Exit Function
        
        On Error GoTo -1
        On Error GoTo error_handler
        
    Else
        
        hr = TipGetConstantValues(pVBProjectNative, VarPtr(sConstant))
        If hr < 0 Then Exit Function
        
    End If
    
    GetConditionalArguments = sConstant
    
    Exit Function
    
error_handler:
    
    ErrorLog "modCallBack::GetConditionalArguments"
    
End Function

' // Set conditional arg
Private Function SetConditionalArguments( _
                 ByVal pVBProjectNative As Long, _
                 ByRef sConstant As String) As Boolean
    Dim hr  As Long
    
    hr = TipSetConstantValues(pVBProjectNative, sConstant)

    SetConditionalArguments = hr >= 0
                     
End Function

' // Run specified event
Private Function MakeEvent( _
                 ByVal pVBProjectNative As Long, _
                 ByRef sEventName As String) As Boolean
    Dim pVBProject  As VBProject
    Dim sEvent      As String
    Dim shInfo      As SHELLEXECUTEINFO
    Dim lStatus     As Long
    
    On Error GoTo error_handler
    
    Set pVBProject = GetProjectFromNativeProject(pVBProjectNative)
    If pVBProject Is Nothing Then Exit Function
    
    On Error Resume Next
    
    sEvent = pVBProject.ReadProperty("TAT", sEventName)
    If Len(sEvent) = 0 Then Exit Function
    
    On Error GoTo -1
    On Error GoTo error_handler
    
    shInfo.cbSize = Len(shInfo)
    shInfo.lpFile = sEvent
    shInfo.fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_FLAG_NO_UI
    shInfo.nShow = SW_SHOWDEFAULT
    
    lStatus = ShellExecuteEx(shInfo)
    If lStatus = 0 Then Exit Function
    
    ' // Wait execution
    lStatus = WaitForSingleObject(shInfo.hProcess, INFINITE)
    If lStatus <> WAIT_OBJECT_0 Then
        CloseHandle shInfo.hProcess
        Exit Function
    End If
    
    CloseHandle shInfo.hProcess
    
    MakeEvent = True
    
    Exit Function
    
error_handler:
    
    ErrorLog "modCallBack::MakeEvent"
    
End Function

' // Get VBProject from native project
Private Function GetProjectFromNativeProject( _
                 ByVal pVBProjectNative As Long) As VBProject
    Dim pProj   As VBProject
    Dim sName   As String
    
    sName = GetProjectName(pVBProjectNative)
    
    For Each pProj In pVBInstance.VBProjects
        
        If pProj.Name = sName Then
            
            Set GetProjectFromNativeProject = pProj
            Exit For
            
        End If
        
    Next
    
End Function

' // Get ProjectName from native project
Private Function GetProjectName( _
                 ByVal pVBProjectNative As Long) As String
    Dim hr  As Long
    
    hr = TipGetProjName(pVBProjectNative, GetProjectName)
    
    If hr < 0 Then
        Err.Raise hr
    End If
    
End Function

' // Call function by pointer
Private Function CallByPointer( _
                 ByVal pFunc As Long, _
                 ByVal lRetType As VbVarType, _
                 ParamArray vParams() As Variant) As Variant
    Dim iTypes()    As Integer: Dim lList()     As Long
    Dim vParam()    As Variant: Dim lIndex      As Long
    Dim pList       As Long:    Dim pTypes      As Long
    Dim resultCall  As Long

    Const CC_STDCALL    As Long = 4

    If LBound(vParams) <= UBound(vParams) Then
        
        ReDim lList(LBound(vParams) To UBound(vParams))
        ReDim iTypes(LBound(vParams) To UBound(vParams))
        ReDim vParam(LBound(vParams) To UBound(vParams))
        
        For lIndex = LBound(vParams) To UBound(vParams)
        
            vParam(lIndex) = vParams(lIndex)
            lList(lIndex) = VarPtr(vParam(lIndex))
            iTypes(lIndex) = VarType(vParam(lIndex))
            
        Next
        
        pList = VarPtr(lList(LBound(lList)))
        pTypes = VarPtr(iTypes(LBound(iTypes)))
        
    End If

    resultCall = DispCallFunc(ByVal 0&, _
                              pFunc, _
                              CC_STDCALL, _
                              lRetType, _
                              UBound(vParams) - LBound(vParams) + 1, _
                              ByVal pTypes, _
                              ByVal pList, _
                              CallByPointer)

    If resultCall Then Err.Raise 5: Exit Function
    
End Function

