Attribute VB_Name = "modHook"
' // modHook.bas - hijacking of API functions
' // © Krivous Anatoly Anatolevich (The trick), 2016

Option Explicit

Private Type tSplicepData
    bData(4)            As Byte     ' // Original data
    lAddressOfOrigin    As Long     ' // Address of original function
    lAddressOfNew       As Long     ' // New function address
    bIsPaused           As Boolean  ' // Determines whether hijack is paused or not
End Type

Private msdHookArray()  As tSplicepData ' // Intercepted functions
Private mlHooksCount    As Long         ' // Number of intercepted functions

' // Intercept the function using the splicing-method
Public Function HookFunction( _
                ByVal lpSourceAddress As Long, _
                ByVal lpDestinationAddress As Long) As Boolean
    Dim lIndex  As Long
    
    On Error GoTo error_handler
    
    lIndex = GetFreeIndex()
    
    ' // Change page attributes
    If VirtualProtect(ByVal lpSourceAddress, 5, PAGE_EXECUTE_READWRITE, 0) = 0 Then Exit Function
    
    ' // Copy original instructions to the new place
    memcpy msdHookArray(lIndex).bData(0), ByVal lpSourceAddress, 5
    
    ' // Place JMP instruction to interception function
    GetMem1 &HE9, ByVal lpSourceAddress
    
    ' // Place relative address to interception function
    GetMem4 CLng(lpDestinationAddress - lpSourceAddress - 5), ByVal lpSourceAddress + 1
    
    msdHookArray(lIndex).lAddressOfOrigin = lpSourceAddress
    msdHookArray(lIndex).lAddressOfNew = lpDestinationAddress
    
    mlHooksCount = mlHooksCount + 1
    
    HookFunction = True
    
    Exit Function
    
error_handler:
    
    ErrorLog "modHook::HookFunction"

End Function

' // Pause hook
Public Function PauseHook( _
                ByVal pAddress As Long) As Boolean
    Dim lIndex  As Long
    
    On Error GoTo error_handler
    
    lIndex = FindHook(pAddress)
    If lIndex < 0 Then Exit Function
    
    If msdHookArray(lIndex).bIsPaused Then Exit Function
    
    ' // Copy original instructions to the old place
    memcpy ByVal msdHookArray(lIndex).lAddressOfOrigin, msdHookArray(lIndex).bData(0), 5
    
    msdHookArray(lIndex).bIsPaused = True
    PauseHook = True
     
    Exit Function
    
error_handler:
    
    ErrorLog "modHook::PauseHook"

End Function

' // Resume hook
Public Function ResumeHook( _
                ByVal pAddress As Long) As Boolean
    Dim lIndex  As Long
    
    On Error GoTo error_handler
    
    lIndex = FindHook(pAddress)
    If lIndex < 0 Then Exit Function
    
    If Not msdHookArray(lIndex).bIsPaused Then Exit Function
    
    ' // Place JMP instruction to interception function
    GetMem1 &HE9, ByVal msdHookArray(lIndex).lAddressOfOrigin
    ' // Place relative address to interception function
    GetMem4 CLng(msdHookArray(lIndex).lAddressOfNew - msdHookArray(lIndex).lAddressOfOrigin - 5), _
            ByVal msdHookArray(lIndex).lAddressOfOrigin + 1
    
    msdHookArray(lIndex).bIsPaused = False
    ResumeHook = True
     
    Exit Function
    
error_handler:
    
    ErrorLog "modHook::ResumeHook"

End Function

' // Remove interception
Public Function UnhookFunction( _
                ByVal pAddress As Long) As Boolean
    Dim lIndex  As Long
    
    On Error GoTo error_handler
    
    lIndex = FindHook(pAddress)
    If lIndex < 0 Then Exit Function
    
    ' // Copy original instructions to the old place
    memcpy ByVal msdHookArray(lIndex).lAddressOfOrigin, msdHookArray(lIndex).bData(0), 5
    
    msdHookArray(lIndex).lAddressOfOrigin = 0
    
    UnhookFunction = True

    Exit Function
    
error_handler:
    
    ErrorLog "modHook::UnhookFunction"

End Function

' // Get free slot in array
Private Function GetFreeIndex() As Long
    Dim lIndex  As Long
    
    For lIndex = 0 To mlHooksCount - 1
        
        If msdHookArray(lIndex).lAddressOfOrigin = 0 Then
        
            GetFreeIndex = lIndex
            Exit Function
            
        End If
        
    Next
    
    GetFreeIndex = mlHooksCount
    
    ReDim Preserve msdHookArray(mlHooksCount + 10)
    
End Function

' // Find slot index by address
Private Function FindHook( _
                 ByVal pAddress As Long) As Long
    Dim lIndex  As Long
    
    For lIndex = 0 To mlHooksCount - 1
        
        If msdHookArray(lIndex).lAddressOfOrigin = pAddress Then
        
            FindHook = lIndex
            Exit Function
            
        End If
        
    Next

    FindHook = -1
    
End Function

