# stdError

## Spec

* Tracks stack
* Supercedes `Err`.
* Should be able to handle errors despite position in code (e.g. Classes / Modules).
* Overcome [error option functionality](https://stackoverflow.com/questions/38132790/debugging-errors-in-vba-classes-in-excel)
* minimal boiler plate
* provides logging of CriticalErrors, Errors, Warning, Info
* provides rendering of logging to HTML, RTF(?), TXT, Debug Window

## Hurdles:

* Automated Stack tracing on error
    * [Might be useful](https://www.vbforums.com/showthread.php?896754-Get-module-(or-and)-class-but-also-sub-function-names&p=5571739&viewfull=1#post5571739)









































## Appendix A

```vb
' //
' // Get calling procedure name
' // The result executable should be compiled with debug symbols
' // by The trick 2022
' //

Option Explicit
Option Base 0

Private Enum PTR    ' // Alias (thanks OlimilO1402)
    [_]
End Enum

Private Const MAX_SYM_NAME                                  As Long = 2000
Private Const MAX_PATH                                      As Long = 260
Private Const SIZEOF_SYMBOL_INFO                            As Long = 88
Private Const GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS        As Long = 4
Private Const GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT  As Long = 2

Private Type SYMBOL_INFO
    SizeOfStruct            As Long
    TypeIndex               As Long
    Reserved(1)             As Currency
    Index                   As Long
    Size                    As Long
    ModBase                 As Currency
    Flags                   As Long
    lPad0                   As Long
    Value                   As Currency
    Address                 As Currency
    Register                As Long
    Scope                   As Long
    Tag                     As Long
    NameLen                 As Long
    MaxNameLen              As Long
    iName(MAX_SYM_NAME - 1) As Integer
End Type

Private Declare Function SymInitialize Lib "dbghelp" _
                         Alias "SymInitializeW" ( _
                         ByVal hProcess As OLE_HANDLE, _
                         ByVal UserSearchPath As Any, _
                         ByVal fInvadeProcess As Long) As Long
Private Declare Function SymFromAddr Lib "dbghelp" _
                         Alias "SymFromAddrW" ( _
                         ByVal hProcess As OLE_HANDLE, _
                         ByVal Address As Currency, _
                         ByRef Displacement As Currency, _
                         ByRef Symbol As SYMBOL_INFO) As Long
Private Declare Function SymLoadModuleEx Lib "dbghelp" _
                         Alias "SymLoadModuleExW" ( _
                         ByVal hProcess As OLE_HANDLE, _
                         ByVal hFile As OLE_HANDLE, _
                         ByVal ImageName As PTR, _
                         ByVal ModuleName As PTR, _
                         ByVal BaseOfDll As Currency, _
                         ByVal DllSize As Long, _
                         ByRef Data As Any, _
                         ByVal Flags As Long) As Long
Private Declare Function GetModuleFileName Lib "kernel32" _
                         Alias "GetModuleFileNameW" ( _
                         ByVal hModule As Long, _
                         ByVal lpFileName As PTR, _
                         ByVal nSize As Long) As Long
Private Declare Function GetModuleHandleEx Lib "kernel32" _
                         Alias "GetModuleHandleExW" ( _
                         ByVal dwFlags As Long, _
                         ByVal lpModuleName As PTR, _
                         ByRef phModule As Any) As Long
Private Declare Function SysAllocString Lib "oleaut32" ( _
                         ByRef pOlechar As Any) As Long
Private Declare Function EbSetMode Lib "vba6" ( _
                         ByVal Mode As Long) As Long
Private Declare Function EbGetCallstackCount Lib "vba6" ( _
                         ByRef lCount As Long) As Long
Private Declare Function EbGetCallstackFunction Lib "vba6" ( _
                         ByVal lIndex As Long, _
                         ByVal pProject As PTR, _
                         ByVal pModule As PTR, _
                         ByVal pFunction As PTR, _
                         ByRef lRet As Long) As Long
    
Private Declare Sub GetMem4 Lib "msvbvm60" ( _
                    ByRef pAddr As Any, _
                    ByRef pRetVal As Any)
Private Declare Sub PutMemPtr Lib "msvbvm60" _
                    Alias "PutMem4" ( _
                    ByRef pAddr As Any, _
                    ByVal pNewVal As PTR)

Private m_bInintialized As Boolean

Public Function GetCallingProcName( _
                Optional ByVal lReserved As Long) As String
    Dim tSymInfo    As SYMBOL_INFO
    Dim cAddr       As Currency
    Dim cDisp       As Currency
    Dim bIsInIDE    As Boolean
    Dim lStackCount As Long
    Dim sProject    As String
    Dim sModule     As String
    Dim sFunction   As String
    
    Debug.Assert MakeTrue(bIsInIDE)
    
    If bIsInIDE Then
        
        EbSetMode 2
        
        If EbGetCallstackCount(lStackCount) >= 0 Then
            If lStackCount > 1 Then
                If EbGetCallstackFunction(1, VarPtr(sProject), VarPtr(sModule), VarPtr(sFunction), 0) >= 0 Then
                    GetCallingProcName = sModule & "::" & sFunction
                End If
            End If
        End If
        
        EbSetMode 1
        
        Exit Function
        
    End If
    
    If Not m_bInintialized Then
        If SymInitialize(VarPtr(m_bInintialized), ByVal 0&, 0) = 0 Then
            Exit Function
        ElseIf SymLoadModuleEx(VarPtr(m_bInintialized), 0, StrPtr(GetExecutableName), 0, 0@, 0, ByVal 0&, 0) = 0 Then
            Exit Function
        Else
            m_bInintialized = True
        End If
    End If
    
    tSymInfo.SizeOfStruct = SIZEOF_SYMBOL_INFO
    tSymInfo.MaxNameLen = MAX_SYM_NAME
    
    GetMem4 ByVal VarPtr(lReserved) - 4, cAddr
    
    If SymFromAddr(VarPtr(m_bInintialized), cAddr, cDisp, tSymInfo) = 0 Then
        Exit Function
    End If
    
    PutMemPtr ByVal VarPtr(GetCallingProcName), SysAllocString(tSymInfo.iName(0))
    
End Function

Private Function MakeTrue( _
                 ByRef bValue As Boolean) As Boolean
    MakeTrue = True
    bValue = True
End Function

Private Function GetExecutableName() As String
    Dim sRet    As String
    Dim lSize   As Long
    Dim hMod    As PTR
    
    If GetModuleHandleEx(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS Or GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT, _
                         AddressOf GetCallingProcName, hMod) = 0 Then
        Exit Function
    End If
    
    sRet = Space$(MAX_PATH)
    lSize = GetModuleFileName(hMod, StrPtr(sRet), Len(sRet))
    
    If lSize Then
        GetExecutableName = Left$(sRet, lSize)
    End If

End Function
```

Usage: 

```vb
MsgBox GetCallingProcName
```

...

You can walk through call stack using RtlCaptureStackBackTrace function like:


```vb
Private Declare Function RtlCaptureStackBackTrace Lib "kernel32" ( _
                         ByVal FramesToSkip As Long, _
                         ByVal FramesToCapture As Long, _
                         ByRef BackTrace As Any, _
                         ByRef BackTraceHash As Any) As Integer

Private Sub Form_Load()
    Dim lPtrs() As Long
    ReDim lPtrs(10)
    RtlCaptureStackBackTrace 0, UBound(lPtrs) + 1, lPtrs(0), ByVal 0&
End Sub
```

But you don't need this if you need a function like LogError. Just copy code from GetCallingProcName to your function and add your logic inside this function. Of course you could using RtlCaptureStackBackTrace and get the all call stack like:

```vb
Public Function GetCallStack() As String
    Dim tSymInfo    As SYMBOL_INFO
    Dim cAddr       As Currency
    Dim cDisp       As Currency
    Dim bIsInIDE    As Boolean
    Dim lStackCount As Long
    Dim sProject    As String
    Dim sModule     As String
    Dim sFunction   As String
    Dim lIndex      As Long
    Dim pAddr()     As PTR
    
    Debug.Assert MakeTrue(bIsInIDE)
    
    If bIsInIDE Then
        EbSetMode 2
        If EbGetCallstackCount(lStackCount) >= 0 Then
            For lIndex = 1 To lStackCount - 1
                If EbGetCallstackFunction(lIndex, VarPtr(sProject), VarPtr(sModule), VarPtr(sFunction), 0) >= 0 Then
                    GetCallStack = GetCallStack & sModule & "::" & sFunction & vbNewLine
                    sProject = vbNullString
                    sModule = vbNullString
                    sFunction = vbNullString
                End If
            Next
        End If
        EbSetMode 1
        Exit Function
    End If
    
    If Not m_bInintialized Then
        If SymInitialize(VarPtr(m_bInintialized), ByVal 0&, 0) = 0 Then
            Exit Function
        ElseIf SymLoadModuleEx(VarPtr(m_bInintialized), 0, StrPtr(GetExecutableName), 0, 0@, 0, ByVal 0&, 0) = 0 Then
            Exit Function
        Else
            m_bInintialized = True
        End If
    End If
    
    tSymInfo.SizeOfStruct = SIZEOF_SYMBOL_INFO
    tSymInfo.MaxNameLen = MAX_SYM_NAME
    
    ReDim pAddr(31)
    
    lStackCount = RtlCaptureStackBackTrace(1, UBound(pAddr) + 1, pAddr(0), ByVal 0&)
    For lIndex = 0 To UBound(pAddr)
        GetMem4 pAddr(lIndex), cAddr
        If SymFromAddr(VarPtr(m_bInintialized), cAddr, cDisp, tSymInfo) Then
            PutMemPtr ByVal VarPtr(sFunction), SysAllocString(tSymInfo.iName(0))
            GetCallStack = GetCallStack & sFunction & vbNewLine
            sFunction = vbNullString
        Else
            GetCallStack = GetCallStack & "<unknown>" & vbNewLine
        End If
    Next
End Function
```

One issue here is that `EbGetCallstackCount` and `EbGetCallstackFunction`