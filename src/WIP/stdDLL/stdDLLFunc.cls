VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdDLLFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Implements stdICallable

#if Win64 then
    Private Const vbLongPtr = vbLongLong
    Private Const PTR_SIZE = 8
#else
    Private Const vbLongPtr = vbLong
    Private Const PTR_SIZE = 4
#end if

#if VBA7 then
    Private Declare PtrSafe Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As LongPtr, ByVal offsetinVft As LongPtr, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByVal paTypes As LongPtr, ByVal paValues As LongPtr, ByRef retVAR As Variant) As Long
    Private Declare PtrSafe Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As LongPtr
    Private Declare PtrSafe Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
    Private Declare PtrSafe Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
    Private Declare PtrSafe Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As LongPtr) As Long
#else
    Private Enum LongPtr
      [_]
    End Enum
    Private Declare Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As LongPtr, ByVal offsetinVft As LongPtr, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByVal paTypes As LongPtr, ByVal paValues As LongPtr, ByRef retVAR As Variant) As Long
    Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As LongPtr
    Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
    Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
    Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As LongPtr) As Long
#end if


Private Enum CALLINGCONVENTION_ENUM
  ' http://msdn.microsoft.com/en-us/library/system.runtime.interopservices.comtypes.callconv%28v=vs.110%29.aspx
  CC_FASTCALL = 0&
  CC_CDECL
  CC_PASCAL
  CC_MACPASCAL
  CC_STDCALL                        ' typical windows APIs
  CC_FPFASTCALL
  CC_SYSCALL
  CC_MPWCDECL
  CC_MPWPASCAL
End Enum
Private Enum STRINGPARAMS_ENUM
    STR_ANSI
    STR_UNICODE
End Enum

Private Type TThis
    CallType as CALLINGCONVENTION_ENUM
    StringParams as STRINGPARAMS_ENUM
    DLL as string
    FuncName as string
    ReturnType as vbVarType
    ParamTypes as vbVarType()

    'internal
    DllHandle as LongPtr
    FuncHandle as LongPtr
End Type
Private This as TThis

'Creates a new stdDLLFunc object with the given parameters. The function will be called using StdCall convention and ANSI string parameters.
'@constructor
'@param dll - The name of the DLL to load
'@param func - The name of the function to call
'@param returnType - The type of the return value
'@param paramTypes - The types of the parameters
'@returns - A new stdDLLFunc object
Public Function CreateStdCallA(ByVal dll as string, ByVal func as string, ByVal returnType as vbVarType, ParamArray paramTypes as vbVarType()) as stdDLLFunc
    set CreateStdCallA = New stdDLLFunc
    Call CreateStdCallA.protInit(CC_STDCALL, STR_ANSI, dll, func, returnType, paramTypes)
End Function

'Creates a new stdDLLFunc object with the given parameters. The function will be called using CDECL convention and ANSI string parameters.
'@constructor
'@param dll - The name of the DLL to load
'@param func - The name of the function to call
'@param returnType - The type of the return value
'@param paramTypes - The types of the parameters
'@returns - A new stdDLLFunc object
Public Function CreateCDeclA(ByVal dll as string, ByVal func as string, ByVal returnType as vbVarType, ParamArray paramTypes as vbVarType()) as stdDLLFunc
    set CreateCDeclA = New stdDLLFunc
    Call CreateCDeclA.protInit(CC_CDECL, STR_ANSI, dll, func, returnType, paramTypes)
End Function

'Creates a new stdDLLFunc object with the given parameters. The function will be called using StdCall convention and unicode string parameters.
'@constructor
'@param dll - The name of the DLL to load
'@param func - The name of the function to call
'@param returnType - The type of the return value
'@param paramTypes - The types of the parameters
'@returns - A new stdDLLFunc object
Public Function CreateStdCallW(ByVal dll as string, ByVal func as string, ByVal returnType as vbVarType, ParamArray paramTypes as vbVarType()) as stdDLLFunc
    set CreateStdCallW = New stdDLLFunc
    Call CreateStdCallW.protInit(CC_STDCALL, STR_UNICODE, dll, func, returnType, paramTypes)
End Function

'Creates a new stdDLLFunc object with the given parameters. The function will be called using CDECL convention and unicode string parameters.
'@constructor
'@param dll - The name of the DLL to load
'@param func - The name of the function to call
'@param returnType - The type of the return value
'@param paramTypes - The types of the parameters
'@returns - A new stdDLLFunc object
Public Function CreateCDeclW(ByVal dll as string, ByVal func as string, ByVal returnType as vbVarType, ParamArray paramTypes as vbVarType()) as stdDLLFunc
    set CreateCDeclW = New stdDLLFunc
    Call CreateCDeclW.protInit(CC_CDECL, STR_UNICODE, dll, func, returnType, paramTypes)
End Function

'Creates a new stdDLLFunc object with the given parameters.
'@constructor
'@param callType - The calling convention to use
'@param stringParams - The string parameters to use
'@param dll - The name of the DLL to load
'@param func - The name of the function to call
'@param returnType - The type of the return value
'@param paramTypes - The types of the parameters
'@returns - A new stdDLLFunc object
Public Function Create(ByVal callType as Long, ByVal stringParams as Long, ByVal dll as string, ByVal func as string, ByVal returnType as vbVarType, ParamArray paramTypes as vbVarType()) as stdDLLFunc
    set Create = New stdDLLFunc
    Call Create.protInit(callType, stringParams, dll, func, returnType, paramTypes)
End Function

'Initialises the stdDLLFunc object
'@protected
'@param callType - The calling convention to use
'@param stringParams - The string parameters to use
'@param dll - The name of the DLL to load
'@param funcName - The name of the function to call
'@param returnType - The type of the return value
'@param paramTypes - The types of the parameters
Public Sub protInit(ByVal callType as CALLINGCONVENTION_ENUM, ByVal stringParams as STRINGPARAMS_ENUM, ByRef dll as string, ByRef funcName as string, ByRef returnType as vbVarType, ByRef paramTypes as vbVarType())
    If Not (returnType And &HFFFF0000) = 0& then Err.Raise 5, "stdDLLFunc#RunEx", "Return type is not a valid type"
    If stringParams < STR_ANSI or stringParams > STR_UNICODE then Err.Raise 5, "stdDLLFunc#RunEx", "Unknown string format provided in initializer"
    
    This.CallType = callType
    This.StringParams = stringParams
    This.DLL = dll
    This.FuncName = funcName
    This.ReturnType = returnType
    This.ParamTypes = paramTypes

    this.DllHandle = GetModuleHandle(this.DLL)
    if this.DllHandle = 0 then
        this.DllHandle = LoadLibrary(this.DLL)
    end if
    this.FuncHandle = GetProcAddress(this.DllHandle, this.FuncName)

    If this.DllHandle = 0 then Err.Raise 5, "stdDLLFunc#RunEx", "Failed to load DLL"
    If this.FuncHandle = 0 then Err.Raise 5, "stdDLLFunc#RunEx", "Failed to get function address"
End Sub

'Runs the function
'@param params - The parameters to run the function with
'@returns - The result of the function
Public Function Run(ParamArray params() as variant) as variant
    Dim vResult as variant
    Call CopyVariant(Run, RunEx(params))
    Return vResult
End Function

'Runs the function with a array of parameters
'@param params - The parameters to run the function with
'@returns - The result of the function
Public Function RunEx(ByVal params as variant) as variant
    Const VT_BYREF = &H4000

    Dim iUB as Long: iUB = UBound(params)
    Dim iLB as Long: iLB = LBound(params)
    Dim argCount as Long: argCount = Abs(iUB - iLB + 1)

    'Patch strings to ensure they are passed as pointers
    Dim i as long
    For i = iLB to iUB
        if This.ParamTypes(i) = vbString then
            Dim dest as LongPtr
            Call CopyMemory(VarPtr(dest), VarPtr(params(i)), 2)
            If (dest And VT_BYREF) = 0& Then
                dest = VarPtr(params(i)) + 8&
            Else
                Call CopyMemory(VarPtr(dest), VarPtr(params(i)) + 8&, 4)
            End If
            
            'Map to pointer
            Dim lStrPtr as LongPtr
            Call CopyMemory(VarPtr(lStrPtr), dest, PTR_SIZE)
            If lStrPtr > 0& Then                    
                Dim sText as string
                If This.StringParams = STR_ANSI Then  
                    'Convert Unicode to ANSI
                    sText = params(pIndex) 
                    ' then re-write the passed String to ANSI
                    Call FillMemory(lStrPtr, LenB(sText), 0)
                    sText = StrConv(sText, vbFromUnicode)
                    Call CopyMemory(lStrPtr, StrPtr(sText), LenB(sText))
                End If
            End If
            params(i) = lStrPtr
        end if
    Next
    
    'Call function
    Dim lRet as Long, vResult as variant: lRet = DispCallFunc(this.funcHandle, 0&, this.callType, this.returnType, UBound(this.paramTypes) + 1&, VarPtr(this.paramTypes(0)), VarPtr(params(iLB)), vResult)
    if lRet = 0& Then
        if isObject(vResult) then
            set RunEx = vResult
        else
            let RunEx = vResult
        end if
    else
        Call SetLastError(lRet)
        Err.Raise lRet, "stdDLLFunc#RunEx", "Failed to call function"
    end if
End Function

'Converts an ANSI string pointer to a string
'@param ANSIpointer - The pointer to the ANSI string
'@returns - The ANSI string stored at the pointer
'@remark ANSIpointer must be a pointer to an ANSI string (1 byte per character)
Public Function PointerToStringA(ByVal ANSIpointer As LongPtr) As String
    Dim lSize As Long, sANSI As String
    If Not ANSIpointer = 0& Then
        lSize = lstrlenA(ANSIpointer)
        If lSize > 0& Then
            sANSI = String$(lSize \ 2& + 1&, vbNullChar)
            Call CopyMemory(StrPtr(sANSI), ANSIpointer, lSize)
            PointerToStringA = Left$(StrConv(sANSI, vbUnicode), lSize)
        End If
    End If
End Function

'Converts a unicode string pointer to a string
'@param UnicodePointer - The pointer to the unicode string
'@returns - The unicode string stored at the pointer
'@remark UnicodePointer must be a pointer to an unicode string (2 bytes per character)
Public Function PointerToStringW(ByVal UnicodePointer As LongPtr) As String
    Dim lSize As Long
    If Not UnicodePointer = 0& Then
        lSize = lstrlenW(UnicodePointer)
        If lSize > 0& Then
            PointerToStringW = Space$(lSize)
            Call CopyMemory(StrPtr(PointerToStringW), UnicodePointer, lSize * 2&)
        End If
    End If
End Function





Private Sub Class_Terminate()
    If Not this.DllHandle = 0& Then
        Call FreeLibrary(this.DllHandle)
    End If
End Sub


