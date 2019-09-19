VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsThunks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'///////////////////////////////////////////////////////////////////////////

'   The _Readme.rtf file has all the documentation you need for the various
'   thunks included in this code page. Documentation herein is a courtesty
'   and is not complete.

'   The two private functions listed below are support functions.
'   If adding thunks to a stand-alone class or usercontrol, i.e., single-file
'   project, then you must include these two functions and the APIs below

'   The CreateTasker_[xxxx] functions that follow are self-contained and
'   can be individually added to a project as desired.

'   This class can simply be added to any project. If you don't want to use
'   a class, then you can copy and paste the entirety of this class' content,
'   as is, to a bas-module of your choosing.

'///////////////////////////////////////////////////////////////////////////

Private Declare Function thunkCreateDispTypeInfo Lib "OleAut32.dll" Alias "CreateDispTypeInfo" (ByVal pidata As Long, ByVal lcid As Long, ByRef pptinfo As IUnknown) As Long
Private Declare Function thunkCreateStdDispatch Lib "OleAut32.dll" Alias "CreateStdDispatch" (ByVal ptinfo As Long, ByVal pvThis As Long, ByVal ptinfo As Long, ByVal ppunkStdDisp As Long) As Long
Private Declare Function thunkVirtualQuery Lib "kernel32.dll" Alias "VirtualQuery" (ByRef lpAddress As Any, ByRef lpBuffer As Any, ByVal dwLength As Long) As Long
Private Declare Sub thunkCopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub thunkCoTaskMemFree Lib "ole32.dll" Alias "CoTaskMemFree" (ByVal pv As Long)
Private Declare Function thunkCoTaskMemAlloc Lib "ole32.dll" Alias "CoTaskMemAlloc" (ByVal cb As Long) As Long
Private Declare Function thunkVirtualAlloc Lib "kernel32.dll" Alias "VirtualAlloc" (ByVal lpAddress As Long, ByVal dwSize As Long, Optional ByVal flAllocationType As Long = &H1000, Optional ByVal flProtect As Long = &H40) As Long
Private Declare Function thunkGetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function thunkGetProcAddress Lib "kernel32.dll" Alias "GetProcAddress" (ByVal hModule As Long, ByRef lpProcName As Any) As Long
' these two are only needed when including the subclass tasker...
Private Declare Function thunkLoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function thunkFreeLibrary Lib "kernel32.dll" Alias "FreeLibrary" (ByVal hLibModule As Long) As Long

Private Function pvCreateITypeInfo(arrMethods As Long, ParamArray pParams()) As stdole.IUnknown

    ' pParams() passed in triples (return type, method name, params if any)
    '   -- notice that a number always precedes list of strings
    '   for simplicity, all params & method vartypes are either Longs or Variants
    '   if name is prefixed with ! then the vartype is variant else long
    '   if name is prefixed with * then the vartype is long ByRef vs ByVal
    
    ' VTable ordinal for 1st method always starts with #3; after IUnknown:Release
    ' DispID always begins with 0 and is incremented, 1st passed method is default method
    
    ' This routine creates a JIT ITypeInfo interface so we can use IDispatch with our
    '   VTable-only IUnknown object created in the various thunk-creation routines.
    '   The end result is passing this typelib along with our VTable-Only IUnknown to
    '   CreateStdDispatch API which creates IDispatch and 'hooks' it into our IUnknown.
    '   That makes the IUnknown a VB-compatible object allowing exposed methods/properties.
    ' What is passed here are the methods/properties and their parameter information
    
    ' the array containing the required structures is serialized like:
    '   header 8 bytes (variable: arrMethods)
    '       4 bytes: pointer to first method structure (mOffset)
    '       4 bytes: method structures count
    '   method structures 28 bytes each (variable: mOffset)
    '       4 bytes: pointer to method name (sOffset+xxx)
    '       4 bytes: pointer to ParamData array if any (pOffset+xxx)
    '       4 bytes: DispatchID (dispID)
    '       4 bytes: method ordinal in VTable, zero-bound
    '       4 bytes: method calling convention
    '       4 bytes: number of items in ParamData array
    '       2 bytes: method type (sub, function, property)
    '       2 bytes: method vartype
    '   param structures 8 bytes each (variable: pOffset)
    '       4 bytes: pointer to parameter name (sOffset+xxx)
    '       2 bytes: param vartype
    '       2 bytes: padding
    '   unique unicode strings, double null terminated, dword aligned (variable: sOffset)
    
    Dim mOffset As Long, pOffset As Long, strOffset As Long
    Dim mCount As Long, pCount As Long, p As Long, n As Long
    Dim colNames As Collection, zMethods(0 To 8) As Long
    Dim bIsParam As Boolean, sName As String
    
    '/// pass 1: tally string sizes, count methods & parameters for array sizing
    Set colNames = New Collection
    For p = 0 To UBound(pParams)
        If VarType(pParams(p)) = vbString Then
            sName = pParams(p)
            If Asc(sName) < 65 Then sName = Mid$(sName, 2)
            On Error Resume Next    ' strings prefixed with 3 char hex offet into array
            colNames.Add Right$("00" & Hex$(strOffset), 3) & sName, sName
            If Err Then
                Err.Clear
            Else
                n = (LenB(sName) + 5) And &HFFFFFFFC
                strOffset = strOffset + n
            End If
            On Error GoTo 0
            If bIsParam = True Then pCount = pCount + 1 Else bIsParam = True
        Else
            bIsParam = False: mCount = mCount + 1
        End If
    Next
    '/// calc overall array size, create array & set offsets
    p = strOffset + mCount * 28 + pCount * 8 + 8
    arrMethods = thunkCoTaskMemAlloc(p)
    Debug.Assert arrMethods <> 0
    
    mOffset = ((arrMethods Xor &H80000000) + 8) Xor &H80000000 ' array position where methods start
    pOffset = ((mOffset Xor &H80000000) + mCount * 28) Xor &H80000000 ' array position where parameters start
    strOffset = ((pOffset Xor &H80000000) + pCount * 8) Xor &H80000000 ' array position where strings start
    zMethods(0) = mOffset: zMethods(1) = mCount
    thunkCopyMemory ByVal arrMethods, zMethods(0), 8    ' copy base structure (header)
    
    '/// copy strings to array
    p = strOffset
    For pCount = 1 To colNames.Count
        sName = Mid$(colNames(pCount), 4)
        n = LenB(sName) + 2
        thunkCopyMemory ByVal p, ByVal StrPtr(sName), n
        p = ((p Xor &H80000000) + ((n + 3) And &HFFFFFFFC)) Xor &H80000000
    Next
    
    '/// pass 2: copy methods & param structures to array
    Erase zMethods(): p = 0
    zMethods(3) = 3: zMethods(4) = 4
    For pCount = 0 To UBound(pParams)
        If VarType(pParams(pCount)) = vbString Then
            sName = pParams(pCount)
            If Asc(sName) < 65 Then sName = Mid$(sName, 2)
            n = CLng("&H" & Left$(colNames(sName), 3))  ' relative offset into array
            If bIsParam = True Then
                zMethods(7) = ((strOffset Xor &H80000000) + n) Xor &H80000000 ' param name offset
                Select Case Asc(pParams(pCount))
                    Case 33: zMethods(8) = vbVariant            ' ! prefix (variant)
                    Case 42: zMethods(8) = vbLong Or &H4000&    ' * prefix (long ByRef)
                    Case Else: zMethods(8) = vbLong             ' no special prefix (long)
                End Select
                thunkCopyMemory ByVal pOffset, zMethods(7), 8
                If zMethods(1) = 0 Then zMethods(1) = pOffset 'set param offset
                zMethods(5) = zMethods(5) + 1               ' increment nr params
                pOffset = ((pOffset Xor &H80000000) + 8) Xor &H80000000
            Else
                zMethods(0) = ((strOffset Xor &H80000000) + n) Xor &H80000000 ' method name offset
                If zMethods(0) <> p Then
                    If p <> 0 Then zMethods(2) = zMethods(2) + 1 ' increment DispID
                    p = zMethods(0)
                End If
                bIsParam = True
            End If
        Else
            If pCount <> 0 Then
                thunkCopyMemory ByVal mOffset, zMethods(0), 28
                zMethods(1) = 0: zMethods(5) = 0 ' reset param offset & param count
                zMethods(3) = zMethods(3) + 1
                mOffset = ((mOffset Xor &H80000000) + 28) Xor &H80000000
            End If
            zMethods(6) = &H30000 Or pParams(pCount)
            bIsParam = False
        End If
    Next
    thunkCopyMemory ByVal mOffset, zMethods(0), 28      ' copy final method
    Set colNames = Nothing
    If thunkCreateDispTypeInfo(arrMethods, &H800, pvCreateITypeInfo) = 0 Then Exit Function
    
ExitRoutine:
    If arrMethods <> 0 Then thunkCoTaskMemFree arrMethods: arrMethods = 0

End Function

Private Function pvGetAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
    
    ' redesigned but based on Paul Caton's zAddrOf method that can find function
    '   pointers within VB modules (forms, classes, etc).
    ' Redesigning includes combining 2 routines into 1, including additional known
    '   VB offsets, and use of VirtualQuery over IsBadCodePtr API.
    
    Dim bSub As Byte, bVal As Byte
    Dim nAddr As Long, vOffset As Long
    Dim nMethod As Long, n As Long
    Dim maxAddr As Long, MBI() As Long  ' faux MEMORY_BASIC_INFORMATION structure
    Dim nAnchor As Long
    
    If nOrdinal < 1 Then Exit Function
    If oCallback Is Nothing Then Exit Function
    
    thunkCopyMemory nAddr, ByVal ObjPtr(oCallback), 4 ' VTable
    For n = 1 To 4
        Select Case n
        Case 1: vOffset = &H1C                  ' known VB class offset
        Case 2: vOffset = &H6F8                 ' known form offset
        Case 3: vOffset = &H710                 ' known property page offset
        Case 4: vOffset = &H7A4                 ' known usercontrol offset
        End Select
        vOffset = ((vOffset Xor &H80000000) + nAddr) Xor &H80000000 ' start offset
        ' should find a method within 8 attempts...
        maxAddr = ((vOffset Xor &H80000000) + 32) Xor &H80000000
        ' walk the VTable looking for first method to be found
        Do While vOffset < maxAddr
            thunkCopyMemory nMethod, ByVal vOffset, 4 ' get function address at VTable entry
            If nMethod <> 0 Then                ' zero = implemented, skip
                thunkCopyMemory bVal, ByVal nMethod, 1 ' get the 1st byte of that method
                If bVal = &H33 Or bVal = &HE9 Then
                    nAnchor = vOffset           ' bingo. got first method, used as anchor
                    bSub = bVal                 ' cache type of code (native vs. pCode)
                    Exit Do                     ' done looking for anchor
                End If
            End If
            vOffset = ((vOffset Xor &H80000000) + 4) Xor &H80000000
        Loop
        If bSub <> 0 Then Exit For
    Next
    If nAnchor = 0 Then Exit Function           ' failure
    
    ReDim MBI(0 To 6)
    ' start with next method & walk til end of object...
    nMethod = ((nAnchor Xor &H80000000) + 4) Xor &H80000000
    ' 1024/4=256 methods. Increase if you must...
    maxAddr = ((nMethod Xor &H80000000) + 1024) Xor &H80000000
    ' walk the VTable looking for last method to be found
    Do While nMethod < maxAddr
        thunkCopyMemory nAddr, ByVal nMethod, 4 ' get function pointer for VTable entry
        MBI(5) = 0
        thunkVirtualQuery ByVal nAddr, MBI(0), 28 ' query for properties
        If (MBI(5) And &HEE) = 0 Then           ' can't be read
            thunkCopyMemory pvGetAddressOf, ByVal nMethod - (nOrdinal * 4), 4
            Exit Do
        ElseIf (MBI(5) And &H101) <> 0 Then     ' else is a page guard or has no access?
            thunkCopyMemory pvGetAddressOf, ByVal nMethod - (nOrdinal * 4), 4
            Exit Do
        End If
        thunkCopyMemory bVal, ByVal nAddr, 1    ' get function's first byte
        If bVal <> bSub Then                    ' anything other than our anchor's byte, done
            thunkCopyMemory pvGetAddressOf, ByVal nMethod - (nOrdinal * 4), 4
            Exit Do
        End If                                  ' move to next VTable entry...
        nMethod = ((nMethod Xor &H80000000) + 4) Xor &H80000000
    Loop
    If (nMethod - nAnchor) / 4 < nOrdinal Then pvGetAddressOf = 0

End Function

Public Function CreateTasker_Callback(VbHost As Object, _
                                    ByVal FunctionOrdinal As Long, _
                                    ByVal ParamsDWordCount As Byte, _
                                    Optional ByVal CallbackAbortValue As Long = 0, _
                                    Optional ByVal IsCDecl As Boolean = False, _
                                    Optional AddIDEsafety As Boolean = True) As Object

    ' returned object has these methods
    ' Method Name   Sample Call             Purpose/Notes on next line
    ' ----------------------------------------------------------------------------
    ' AddressOf     lAddr=oTasker.AddressOf
    '               Returns thunk address for callbacks (Default method)
    ' Tag           oTasker.Tag=1908: Print oTasker.Tag
    '               returns/sets user-defined thunk tag (Long values only)
    
    ' this function's parameters
    '   VbHost :: the host that this thunk object is declared in
    '   FunctionOrdinal :: last private method in host is #1, second to last is #2, etc
    '   ParamsDWordCount :: number of 4-byte blocks used by all callback parameters
    '       In most cases, this matches the number of callback parameters
    '       Use this rule of thumb: Every typical VarType uses 1 DWord unless...
    '           it's Currency,Date,Double and in that case, uses 2 DWords
    '       When in doubt: ((ParamByteSize + 3) And &HFFC) / 4
    '   CallbackAbortValue :: which value to return to stop callbacks
    '   IsCDecl :: set to true only if the callback sender expectes CDecl
    '   AddIDEsafety :: ignored unless in IDE
    
    Dim z_Sc() As Long, hTLBstructs As Long
    Dim hThunk As Long, hData As Long
    Dim oIDispatch As Long, lEbMode As Long
    Dim oIUnk As stdole.IUnknown
    Dim oITypeInfo As stdole.IUnknown
    Dim hMod As New Collection
    
    Const THUNK_SIZE As Long = 110          ' size of the thunk array (in longs)
    Const VTable_SIZE As Long = 6           ' nr of VTable entries (in longs)
    Const DATA_SIZE As Long = 16            ' Object data size (in longs)
    
    FunctionOrdinal = pvGetAddressOf(VbHost, FunctionOrdinal)
    Debug.Assert FunctionOrdinal <> 0
    ' if invalid function index was passed or VbHost is Nothing then abort
    If FunctionOrdinal = 0 Then Exit Function
    
    ' call local routine to create a temporary typelib for adhoc IDispatch interface
    ' 1=function, 2=propGet, 4=propLet
    Set oITypeInfo = pvCreateITypeInfo(hTLBstructs, _
                                    2, "AddressOf", 2, "Tag", 4, "Tag", "Value")
    Debug.Assert (ObjPtr(oITypeInfo) <> 0)
    If oITypeInfo Is Nothing Then GoTo AbortCleanup
    
    hData = thunkCoTaskMemAlloc(DATA_SIZE * 4) ' memory for our ObjPtr
    Debug.Assert hData <> 0
    If hData = 0 Then GoTo AbortCleanup
    
    hThunk = thunkVirtualAlloc(0, (THUNK_SIZE + VTable_SIZE) * 4)
    Debug.Assert hThunk <> 0
    If hThunk = 0 Then GoTo AbortCleanup    ' thunk creation failed
    
    Debug.Print "thunk: "; hThunk
                                        
    GoSub SetAPIreferences  ' Gosub? Just wanted to declutter top half of routine
    
    ReDim z_Sc(0 To THUNK_SIZE + VTable_SIZE - 1) ' the thunk
    z_Sc(0) = &HBB60D231: z_Sc(1) = &H12345678: z_Sc(2) = &H9E9E589: z_Sc(3) = &H90000001: z_Sc(4) = &HBB60C031: z_Sc(5) = &H12345678: z_Sc(6) = &HB9E9E589: z_Sc(7) = &H90000000: z_Sc(8) = &H8BE58960: z_Sc(9) = &H458B245D: z_Sc(10) = &H4488B28: z_Sc(11) = &H8108480B: z_Sc(12) = &HC0F9&: z_Sc(13) = &H81377500: z_Sc(14) = &HC78&: z_Sc(15) = &H2E754600: z_Sc(16) = &HC985088B: z_Sc(17) = &H558B0C75: z_Sc(18) = &HFF1A892C: z_Sc(19) = &HC0310443: z_Sc(20) = &H438B21EB: z_Sc(21) = &H74C0851C: z_Sc(22) = &HF98115: z_Sc(23) = &H75000204: z_Sc(24) = &H2C75FF0D: z_Sc(25) = &H502875FF: z_Sc(26) = &H10FF008B: z_Sc(27) = &H2B805EB: z_Sc(28) = &H89800040: z_Sc(29) = &HC2611C45
    z_Sc(30) = &HF689000C: z_Sc(31) = &H424548B: z_Sc(32) = &H4004428B: z_Sc(33) = &HC2044289: z_Sc(34) = &HF6890004: z_Sc(35) = &H424548B: z_Sc(36) = &H4804428B: z_Sc(37) = &H85044289: z_Sc(38) = &H600C75C0: z_Sc(39) = &H5D8BE589: z_Sc(40) = &HD6E824: z_Sc(41) = &HC2610000: z_Sc(42) = &HF6890004: z_Sc(43) = &H245C8B53: z_Sc(44) = &H14438B08: z_Sc(45) = &H4C25B: z_Sc(46) = &H245C8B53: z_Sc(47) = &H18438B08: z_Sc(48) = &H4C25B: z_Sc(49) = &H245C8B53: z_Sc(50) = &H24448B08: z_Sc(51) = &H1843890C: z_Sc(52) = &HC25BC031: z_Sc(53) = &HF6890008: z_Sc(54) = &H87B83: z_Sc(55) = &HC2610474: z_Sc(56) = &H75FF0010: z_Sc(57) = &H2475FF2C: z_Sc(58) = &HFF3053FF: z_Sc(59) = &H53FF2073
    z_Sc(60) = &H28438B34: z_Sc(61) = &H8B144589: z_Sc(62) = &HFF53147B: z_Sc(63) = &H45C73453: z_Sc(64) = &H800030: z_Sc(65) = &H2C45C700: z_Sc(67) = &H61287D89: z_Sc(68) = &H24048958: z_Sc(69) = &HF689E2FF: z_Sc(70) = &H8B0843FF: z_Sc(71) = &H45893C43: z_Sc(72) = &H107B831C: z_Sc(73) = &HE8077500: z_Sc(74) = &H74&: z_Sc(75) = &H438B31EB: z_Sc(76) = &H74C08524: z_Sc(77) = &H83D0FF07: z_Sc(78) = &H237501F8: z_Sc(79) = &H501C458D: z_Sc(80) = &H81384B8B: z_Sc(81) = &HFFFFE1: z_Sc(82) = &HFC0EE300: z_Sc(83) = &H2924758D: z_Sc(84) = &H243C8DCC: z_Sc(85) = &HF302E9C1: z_Sc(86) = &HC73FFA5: z_Sc(87) = &HFF1053FF: z_Sc(88) = &H4B8B084B: z_Sc(89) = &H78C98538
    z_Sc(90) = &HFFE1810F: z_Sc(91) = &H890000FF: z_Sc(92) = &H5A61184D: z_Sc(93) = &HE2FFCC01: z_Sc(94) = &HF689C361: z_Sc(95) = &HC43C7: z_Sc(96) = &HC7000000: z_Sc(97) = &H1043&: z_Sc(98) = &H438B0000: z_Sc(99) = &H74C0851C: z_Sc(100) = &H8B5006: z_Sc(101) = &HE80850FF: z_Sc(102) = &H4&: z_Sc(103) = &H768DC3: z_Sc(104) = &H8314438B: z_Sc(105) = &H685010C0: z_Sc(106) = &H3E8&: z_Sc(107) = &H6A0C73FF: z_Sc(108) = &H2C53FF00: z_Sc(109) = &HC30C4389

    '/// patch thunk & build VTable -- same order as sent to pvCreateITypeInfo
    z_Sc(1) = hData: z_Sc(5) = hData        ' 2 non-VTable functions, pass objptr to them
    z_Sc(THUNK_SIZE + 0) = ((hThunk Xor &H80000000) + &H20) Xor &H80000000 ' IUnknown:QueryInterface(QI)
    z_Sc(THUNK_SIZE + 1) = ((hThunk Xor &H80000000) + &H7C) Xor &H80000000 ' IUnknown:AddRef
    z_Sc(THUNK_SIZE + 2) = ((hThunk Xor &H80000000) + &H8C) Xor &H80000000 ' IUnknown:Release
    z_Sc(THUNK_SIZE + 3) = ((hThunk Xor &H80000000) + &HAC) Xor &H80000000 ' AddressOf
    z_Sc(THUNK_SIZE + 4) = ((hThunk Xor &H80000000) + &HB8) Xor &H80000000 ' Tag Get
    z_Sc(THUNK_SIZE + 5) = ((hThunk Xor &H80000000) + &HC4) Xor &H80000000 ' Tag Let
    thunkCopyMemory ByVal hThunk, z_Sc(0), (THUNK_SIZE + VTable_SIZE) * 4
    Erase z_Sc()
    
    ' /// create the VB object's VTable
    ReDim z_Sc(DATA_SIZE - 1)               ' update VTable location
    z_Sc(0) = ((hThunk Xor &H80000000) + THUNK_SIZE * 4) Xor &H80000000
    z_Sc(1) = 1                             ' ref counter
    z_Sc(2) = 0                             ' recursion counter
    z_Sc(3) = ObjPtr(VbHost)                ' passed VB host
    z_Sc(4) = FunctionOrdinal               ' host function pointer
    z_Sc(5) = hThunk                        ' this thunk address
    z_Sc(6) = 0                             ' thunk tag
    z_Sc(7) = 0                             ' IDispatch::IUnknown
    z_Sc(8) = hTLBstructs                   ' TypeLib structs
    z_Sc(9) = lEbMode                       ' optional IDE safety EbMode API or null
    z_Sc(10) = thunkGetProcAddress(hMod("k32"), ByVal "VirtualFree")
    z_Sc(11) = thunkGetProcAddress(hMod("u32"), ByVal "SetTimer")
    z_Sc(12) = thunkGetProcAddress(hMod("u32"), ByVal "KillTimer")
    z_Sc(13) = thunkGetProcAddress(hMod("o32"), ByVal "CoTaskMemFree")
    z_Sc(14) = ParamsDWordCount * 4         ' number bytes used by callback params
    If IsCDecl = True Or ParamsDWordCount = 0 Then z_Sc(14) = z_Sc(14) Or &H80000000
    z_Sc(15) = CallbackAbortValue           ' callback cancellation value
    thunkCopyMemory ByVal hData, z_Sc(0), DATA_SIZE * 4
    thunkCopyMemory oIUnk, hData, 4
    Erase z_Sc()
    Set hMod = Nothing
    
    Debug.Print "data section: "; hData
    
    ' create the adhoc IDispatch. Enables VB to recognize this as an object
    thunkCreateStdDispatch hData, hData, ObjPtr(oITypeInfo), VarPtr(oIDispatch)
    Set oITypeInfo = Nothing
    
    Debug.Assert (oIDispatch <> 0)
    If oIDispatch <> 0 Then                 ' if failure, we unwind
        ' copy the IDispatch pointer to the object data
        thunkCopyMemory ByVal ((hData Xor &H80000000) + 28) Xor &H80000000, oIDispatch, 4
        Set CreateTasker_Callback = oIUnk
    End If
    Set oIUnk = Nothing ' will self-delete thunk & memory allocations if routine failed
    Exit Function
    
SetAPIreferences:       ' modules used by the thunk
    hMod.Add thunkGetModuleHandle("user32.dll"), "u32"
    hMod.Add thunkGetModuleHandle("kernel32.dll"), "k32"
    hMod.Add thunkGetModuleHandle("ole32.dll"), "o32"
    If AddIDEsafety = True Then
        On Error Resume Next
        Debug.Print 1 / 0
        If Err Then
            Err.Clear: On Error GoTo 0
            hMod.Add thunkGetModuleHandle("vba6.dll"), "vba"
            ' note: if next line errors, change vba6.dll to vba5.dll
            lEbMode = thunkGetProcAddress(hMod("vba"), ByVal "EbMode")
        End If
        On Error GoTo 0
    End If
    Return
    
AbortCleanup:
    If hData <> 0 Then thunkCoTaskMemFree hData
    If hTLBstructs <> 0 Then
        Set oITypeInfo = Nothing
        thunkCoTaskMemFree hTLBstructs
    End If

End Function

Public Function CreateTasker_Subclass(VbHost As Object, _
                                    ByVal FunctionOrdinal As Long, _
                                    ByVal InitHwnd As Long, _
                                    Optional UseExtendedVersion As Boolean = False, _
                                    Optional AddIDEsafety As Boolean = True) As Object
                                    
    ' returned object has these methods (non-extended version)
    ' Method Name       Sample Call             Purpose/Notes on next line
    ' ----------------------------------------------------------------------------
    ' Pause             no parameters. oTasker.Pause
    '                   temporarily stops sending messages to your project
    ' LastInChainA      lRet=oTasker.LastInChainA(hWnd, uMsg, wParam, lParam)
    '                   skips chain, calls DefWindowProcA
    '                   function should only be used inside subclass procedure as needed
    ' LastInChainW      lRet=oTasker.LastInChainW(hWnd, uMsg, wParam, lParam)
    '                   skips chain, calls DefWindowProcW
    '                   function should only be used inside subclass procedure as needed
    ' NextInChain       lRet=oTasker.NextInChain(hWnd, uMsg, wParam, lParam)
    '                   passes message to next in subclass chain
    '                   function should only be used inside subclass procedure as needed
    ' Resume            no parameters. oTasker.Resume
    '                   resumes sending messages to your project
    ' State             Default method, no parameters. Print oTasker.State
    '                   returns -1 if currently paused
    '                       ## else number of actively subclassed hWnds
    ' Tag               oTasker.Tag=1908: Print oTasker.Tag
    '                   returns/sets user-defined thunk tag (Long values only)
    ' TagHwnd           oTasker(Me.hWnd)=1234: Print oTasker(Me.Hwnd)
    '                   returns/sets a Long value tag for a specific hWnd
    
    ' Extended version only
    ' Method Name       Sample Call             Purpose/Notes on next line
    ' ----------------------------------------------------------------------------
    ' AddFilterMessage  oTasker.AddFilterMessage 0|1, Array(msg1, msg2): oTasker.AddFilterMessage 0, WM_KILLFOCUS
    '                   appends messages to the thunk filter.
    '                   returns number of filter messages or -1 if invalid parameter
    '                   pass last parameter as single message, Array(...), or Long array
    '                   pass 1st parameter: 0 = inclusive filter else non-zero = exclusive list
    '                   inclusive: only forward messages in the list
    '                   exclusive: only forward messages not in the list
    ' AddWindow         oTasker.AddWindow Array(Text1.hWnd, Text2.hWnd, Text3.hWnd)
    '                   subclass a list of hWnds or single hWnd
    '                   returns number of subclassed windows or -1 if invalid parameter
    '                   pass last parameter as single message, Array(...), or Long array
    ' IsFiltered        no parameters. Print oTasker.IsFiltered
    '                   returns filter pointer if a thunk filter exists
    ' IsWindowManaged   oTasker.IsWindowManaged(Me.hWnd)
    '                   returns subclass-list pointer if passed hWnd is subclassed
    ' RemoveFilter      no parameters. Resets thunk filter, always returns zero
    ' RemoveWindow      oTasker.RemoveWindow -1: oTasker.RemoveWindow VarPtr(Array(Text1.hWnd, Me.Hwnd))
    '                   removes subclassing for list of hWnds
    '                   returns number of subclassed windows or -1 if invalid parameter
    '                   pass last parameter as single message, Array(...), or Long array
    '                   passing -1 unsubclasses all hWnds
    
    ' this function's parameters
    '   VbHost :: the host that this thunk object is declared in
    '   FunctionOrdinal :: last private method in host is #1, second to last is #2, etc
    '   InitHwnd :: hWnd required if UseExtendedVersion=False, subclass passed hWnd
    '   UseExtendedVersion :: add the extra methods to the object or don't
    '   AddIDEsafety :: ignored unless in IDE
    
    Dim z_Sc() As Long, n As Long
    Dim hThunk As Long, hData As Long
    Dim oIDispatch As Long, lEbMode As Long
    Dim hTLBstructs As Long, oObj As Object
    Dim oIUnk As stdole.IUnknown
    Dim oITypeInfo As stdole.IUnknown
    Dim hMod As New Collection
    Dim lThunkSize As Long, lTableSize As Long
    
    Const THUNK_SIZE As Long = 280          ' size of the thunk array (in longs)
    Const VTable_SIZE As Long = 13          ' nr of VTable entries (in longs)
    Const THUNK_SIZE_EX As Long = 510       ' size of the thunk array (in longs)
    Const VTable_SIZE_EX As Long = 20       ' nr of VTable entries (in longs)
    Const DATA_SIZE As Long = 26            ' Object data size (in longs)
    
    FunctionOrdinal = pvGetAddressOf(VbHost, FunctionOrdinal)
    Debug.Assert FunctionOrdinal <> 0
    ' if invalid function index was passed or VbHost is Nothing then abort
    If FunctionOrdinal = 0 Then Exit Function
    
    ' call local routine to create a temporary typelib for adhoc IDispatch interface
    ' 1=function, 2=propGet, 4=propLet
    If UseExtendedVersion = False Then
        Debug.Assert (InitHwnd <> 0)
        If InitHwnd = 0 Then Exit Function  ' required if not using extended version
        lThunkSize = THUNK_SIZE: lTableSize = VTable_SIZE
        Set oITypeInfo = pvCreateITypeInfo(hTLBstructs, _
                        1, "State", 2, "Tag", 4, "Tag", "Value", _
                        2, "TagHwnd", "hWnd", 4, "TagHwnd", "hWnd", "Value", _
                        1, "Pause", 1, "Resume", _
                        1, "NextInChain", "hWnd", "uMsg", "wParam", "lParam", _
                        1, "LastInChainA", "hWnd", "uMsg", "wParam", "lParam", _
                        1, "LastInChainW", "hWnd", "uMsg", "wParam", "lParam")
    Else
        lThunkSize = THUNK_SIZE_EX: lTableSize = VTable_SIZE_EX
        Set oITypeInfo = pvCreateITypeInfo(hTLBstructs, _
                        1, "State", 2, "Tag", 4, "Tag", "Value", _
                        2, "TagHwnd", "hWnd", 4, "TagHwnd", "hWnd", "Value", _
                        1, "Pause", 1, "Resume", _
                        1, "NextInChain", "hWnd", "uMsg", "wParam", "lParam", _
                        1, "LastInChainA", "hWnd", "uMsg", "wParam", "lParam", _
                        1, "LastInChainW", "hWnd", "uMsg", "wParam", "lParam", _
                        1, "IsWindowManaged", "hWnd", 1, "IsFiltered", _
                        1, "AddWindow", "!pItems", 1, "RemoveWindow", "!pItems", _
                        1, "RemoveFilter", 1, "AddFilterMessage", "fType", "!pItems", _
                        1, "GetWindows", "pCount", "pItems")
    End If
    Debug.Assert (ObjPtr(oITypeInfo) <> 0)
    If oITypeInfo Is Nothing Then GoTo AbortCleanup
    
    GoSub SetAPIreferences  ' Gosub? Just wanted to declutter top half of routine
    ' set references before creating thunk. If 1 reference fails, must exit function
    
    hData = thunkCoTaskMemAlloc(DATA_SIZE * 4) ' memory for our ObjPtr
    Debug.Assert hData <> 0
    If hData = 0 Then GoTo AbortCleanup
    
    hThunk = thunkVirtualAlloc(0, (lThunkSize + lTableSize) * 4)
    Debug.Assert hThunk <> 0
    If hThunk = 0 Then GoTo AbortCleanup    ' thunk creation failed
    
    Debug.Print "thunk: "; hThunk
                                        
    ReDim z_Sc(0 To lThunkSize + lTableSize - 1) ' the thunk
    z_Sc(0) = &HBB60D231: z_Sc(1) = &H12345678: z_Sc(2) = &HD9E9E589: z_Sc(3) = &H90000001: z_Sc(4) = &HBB60C031: z_Sc(5) = &H12345678: z_Sc(6) = &H25E9E589: z_Sc(7) = &H90000002: z_Sc(8) = &H8BE58960: z_Sc(9) = &H458B245D: z_Sc(10) = &H4488B28: z_Sc(11) = &H8108480B: z_Sc(12) = &HC0F9&: z_Sc(13) = &H81377500: z_Sc(14) = &HC78&: z_Sc(15) = &H2E754600: z_Sc(16) = &HC985088B: z_Sc(17) = &H558B0C75: z_Sc(18) = &HFF1A892C: z_Sc(19) = &HC0310443: z_Sc(20) = &H438B21EB: z_Sc(21) = &H74C0851C: z_Sc(22) = &HF98115: z_Sc(23) = &H75000204: z_Sc(24) = &H2C75FF0D: z_Sc(25) = &H502875FF: z_Sc(26) = &H10FF008B: z_Sc(27) = &H2B805EB: z_Sc(28) = &H89800040: z_Sc(29) = &HC2611C45
    z_Sc(30) = &HF689000C: z_Sc(31) = &H424548B: z_Sc(32) = &H4004428B: z_Sc(33) = &HC2044289: z_Sc(34) = &HF6890004: z_Sc(35) = &H424548B: z_Sc(36) = &H4804428B: z_Sc(37) = &H85044289: z_Sc(38) = &H600C75C0: z_Sc(39) = &H5D8BE589: z_Sc(40) = &H2D6E824: z_Sc(41) = &HC2610000: z_Sc(42) = &HF6890004: z_Sc(43) = &H245C8B53: z_Sc(44) = &H60538B08: z_Sc(45) = &H6C2F7: z_Sc(46) = &H5740000: z_Sc(47) = &HEBFFC883: z_Sc(48) = &H58438B18: z_Sc(49) = &H1174C085: z_Sc(50) = &H1C2F7: z_Sc(51) = &H7740000: z_Sc(52) = &H1B8&: z_Sc(53) = &H8B02EB00: z_Sc(54) = &H4C25B00: z_Sc(55) = &H768D00: z_Sc(56) = &H245C8B53: z_Sc(57) = &H18438B08: z_Sc(58) = &H4C25B: z_Sc(59) = &H245C8B53
    z_Sc(60) = &H24448B08: z_Sc(61) = &H1843890C: z_Sc(62) = &HC25BC031: z_Sc(63) = &HF6890008: z_Sc(64) = &H8960C031: z_Sc(65) = &H245D8BE5: z_Sc(66) = &H8D28558B: z_Sc(67) = &H8B501C45: z_Sc(68) = &HD0311443: z_Sc(69) = &H50D03150: z_Sc(70) = &H4C53FF52: z_Sc(71) = &H8C261: z_Sc(72) = &H8960C031: z_Sc(73) = &H245D8BE5: z_Sc(74) = &H28558BFC: z_Sc(75) = &H307E8: z_Sc(76) = &H74C08500: z_Sc(77) = &H2C75FF13: z_Sc(78) = &H3114438B: z_Sc(79) = &HD03152D0: z_Sc(80) = &H53FF5250: z_Sc(81) = &H1C458940: z_Sc(82) = &HCC261: z_Sc(83) = &H245C8B53: z_Sc(84) = &H60438B08: z_Sc(85) = &H8902C883: z_Sc(86) = &HC25B6043: z_Sc(87) = &HF6890004: z_Sc(88) = &H8960C031: z_Sc(89) = &H245D8BE5
    z_Sc(90) = &HF760538B: z_Sc(91) = &H80C2&: z_Sc(92) = &HFC1E7400: z_Sc(93) = &H80F281: z_Sc(94) = &H31520000: z_Sc(95) = &H5C538BFF: z_Sc(96) = &HE85C7B89: z_Sc(97) = &H1A8&: z_Sc(98) = &H75C0855A: z_Sc(99) = &HEBD23112: z_Sc(100) = &H4C2F70E: z_Sc(101) = &H75000000: z_Sc(102) = &H2CA8306: z_Sc(103) = &H8902F283: z_Sc(104) = &H55896053: z_Sc(105) = &H4C2611C: z_Sc(106) = &H768D00: z_Sc(107) = &H53E58955: z_Sc(108) = &H8B085D8B: z_Sc(109) = &HB1E84853: z_Sc(110) = &H5B000001: z_Sc(111) = &H14C25D: z_Sc(112) = &H53E58955: z_Sc(113) = &H8B085D8B: z_Sc(114) = &H9DE85053: z_Sc(115) = &H5B000001: z_Sc(116) = &H14C25D: z_Sc(117) = &H53E58955: z_Sc(118) = &H8B085D8B: z_Sc(119) = &H89E85453
    z_Sc(120) = &H5B000001: z_Sc(121) = &H14C25D: z_Sc(122) = &HE80843FF: z_Sc(123) = &HA0&: z_Sc(124) = &H3775F685: z_Sc(125) = &H1CFE8: z_Sc(126) = &H1C458D00: z_Sc(127) = &H14458D50: z_Sc(128) = &H1875FF50: z_Sc(129) = &HFF3875FF: z_Sc(130) = &H75FF3075: z_Sc(131) = &H2875FF2C: z_Sc(132) = &HFF2475FF: z_Sc(133) = &H53FF0C73: z_Sc(134) = &H18458B10: z_Sc(135) = &H674C085: z_Sc(136) = &HFF008B50: z_Sc(137) = &H75390850: z_Sc(138) = &HFF127514: z_Sc(139) = &H75FF3075: z_Sc(140) = &H2875FF2C: z_Sc(141) = &HFF2475FF: z_Sc(142) = &H45894853: z_Sc(143) = &H84BFF1C: z_Sc(144) = &H18C261: z_Sc(145) = &H74084339: z_Sc(146) = &H10C26104: z_Sc(147) = &H2C75FF00: z_Sc(148) = &HFF2475FF: z_Sc(149) = &H4B8B3053
    z_Sc(150) = &H5104E364: z_Sc(151) = &H6A3C53FF: z_Sc(152) = &H2073FF00: z_Sc(153) = &H8B3853FF: z_Sc(154) = &H45892843: z_Sc(155) = &H147B8B14: z_Sc(156) = &HFF53006A: z_Sc(157) = &H45C73853: z_Sc(158) = &H800030: z_Sc(159) = &H2C45C700: z_Sc(161) = &H61287D89: z_Sc(162) = &H24048958: z_Sc(163) = &HF689E2FF: z_Sc(164) = &H8BDE89FC: z_Sc(165) = &H4A96043: z_Sc(166) = &H74000000: z_Sc(167) = &H10EE807: z_Sc(168) = &H3AEB0000: z_Sc(169) = &H2A9&: z_Sc(170) = &H31227500: z_Sc(171) = &H5C7B8BF6: z_Sc(172) = &H974FF85: z_Sc(173) = &H407E8: z_Sc(174) = &H75F68500: z_Sc(175) = &H24438B10: z_Sc(176) = &H974C085: z_Sc(177) = &HF883D0FF: z_Sc(178) = &H89027401: z_Sc(179) = &H287D83DE
    z_Sc(180) = &H8B0B7502: z_Sc(181) = &H7B8B2455: z_Sc(182) = &H2E858: z_Sc(183) = &H90C30000: z_Sc(184) = &HFF85C031: z_Sc(185) = &HD2854774: z_Sc(186) = &H49E84374: z_Sc(187) = &H85000001: z_Sc(188) = &H503A74C0: z_Sc(189) = &H3114438B: z_Sc(190) = &HD03150D0: z_Sc(191) = &H53FF5250: z_Sc(192) = &HC0855A44: z_Sc(193) = &H43F72774: z_Sc(194) = &H160&: z_Sc(195) = &H8B197500: z_Sc(196) = &HFE3490F: z_Sc(197) = &HCA390F89: z_Sc(198) = &H4C8B137F: z_Sc(199) = &HC89048F: z_Sc(200) = &H510AEB97: z_Sc(201) = &H3853FF57: z_Sc(202) = &H7B89FF31: z_Sc(203) = &HF689C358: z_Sc(204) = &HD285C031: z_Sc(205) = &HFDE83374: z_Sc(206) = &H85000000: z_Sc(207) = &H522A75C0: z_Sc(208) = &H14438B50: z_Sc(209) = &H3150D031
    z_Sc(210) = &HFF5250D0: z_Sc(211) = &H855A4053: z_Sc(212) = &HF71674C0: z_Sc(213) = &H16043: z_Sc(214) = &H5740000: z_Sc(215) = &HEB585389: z_Sc(216) = &H410F8B08: z_Sc(217) = &H14890F89: z_Sc(218) = &HF689C38F: z_Sc(219) = &HFF1875FF: z_Sc(220) = &H75FF1475: z_Sc(221) = &HC75FF10: z_Sc(222) = &H90C3D2FF: z_Sc(223) = &H4604B83: z_Sc(224) = &HC43C7: z_Sc(225) = &HE8000000: z_Sc(226) = &H7C&: z_Sc(227) = &H855C7B8B: z_Sc(228) = &H310A74FF: z_Sc(229) = &H33EE8C9: z_Sc(230) = &H7B890000: z_Sc(231) = &H1C438B5C: z_Sc(232) = &H674C085: z_Sc(233) = &HFF008B50: z_Sc(234) = &H1E80850: z_Sc(235) = &HC3000000: z_Sc(236) = &H8314438B: z_Sc(237) = &H6A5010C0: z_Sc(238) = &HC73FF64: z_Sc(239) = &H53FF006A
    z_Sc(240) = &HC43892C: z_Sc(241) = &H768DC3: z_Sc(242) = &HC710EC83: z_Sc(243) = &H4002404: z_Sc(244) = &H44C70002: z_Sc(245) = &H424&: z_Sc(246) = &H44C70000: z_Sc(247) = &HC00824: z_Sc(248) = &H44C70000: z_Sc(249) = &HC24&: z_Sc(250) = &H148D4600: z_Sc(251) = &H1845C724: z_Sc(253) = &H8D18458D: z_Sc(254) = &H8B525108: z_Sc(255) = &H8B501C43: z_Sc(256) = &H8310FF00: z_Sc(257) = &H90C310C4: z_Sc(258) = &H85587B8B: z_Sc(259) = &HFC2574FF: z_Sc(260) = &HA960438B: z_Sc(261) = &H1&: z_Sc(262) = &HFA890974: z_Sc(263) = &HFFFEBFE8: z_Sc(264) = &H8B11EBFF: z_Sc(265) = &H768D0F: z_Sc(266) = &H518F148B: z_Sc(267) = &HFFFEAFE8: z_Sc(268) = &HF4E259FF: z_Sc(269) = &H768DC3
    z_Sc(270) = &H1E74FF85: z_Sc(271) = &H83604B8B: z_Sc(272) = &H7E301E1: z_Sc(273) = &H1275FA39: z_Sc(274) = &H8BC3C889: z_Sc(275) = &H8951570F: z_Sc(276) = &H4C783D0: z_Sc(277) = &H5F58AFF2: z_Sc(278) = &HC0310374: z_Sc(279) = &HC3C829C3


    '/// patch thunk & build VTable -- same order as sent to pvCreateITypeInfo
    z_Sc(1) = hData: z_Sc(5) = hData        ' 2 non-VTable functions, pass objptr to them
    z_Sc(lThunkSize + 0) = ((hThunk Xor &H80000000) + &H20) Xor &H80000000 ' IUnknown:QueryInterface(QI)
    z_Sc(lThunkSize + 1) = ((hThunk Xor &H80000000) + &H7C) Xor &H80000000 ' IUnknown:AddRef
    z_Sc(lThunkSize + 2) = ((hThunk Xor &H80000000) + &H8C) Xor &H80000000 ' IUnknown:Release
    z_Sc(lThunkSize + 3) = ((hThunk Xor &H80000000) + &HAC) Xor &H80000000 ' State
    z_Sc(lThunkSize + 4) = ((hThunk Xor &H80000000) + &HE0) Xor &H80000000 ' Tag Get
    z_Sc(lThunkSize + 5) = ((hThunk Xor &H80000000) + &HEC) Xor &H80000000 ' Tag Let
    z_Sc(lThunkSize + 6) = ((hThunk Xor &H80000000) + &H100) Xor &H80000000 ' TagHwnd Get
    z_Sc(lThunkSize + 7) = ((hThunk Xor &H80000000) + &H120) Xor &H80000000 ' TagHwnd Let
    z_Sc(lThunkSize + 8) = ((hThunk Xor &H80000000) + &H14C) Xor &H80000000 ' Pause
    z_Sc(lThunkSize + 9) = ((hThunk Xor &H80000000) + &H160) Xor &H80000000 ' Resume
    z_Sc(lThunkSize + 10) = ((hThunk Xor &H80000000) + &H1AC) Xor &H80000000 ' NextInChain
    z_Sc(lThunkSize + 11) = ((hThunk Xor &H80000000) + &H1C0) Xor &H80000000 ' LastInChainA
    z_Sc(lThunkSize + 12) = ((hThunk Xor &H80000000) + &H1D4) Xor &H80000000 ' LastInChainW
    
    '/// if extended version never wanted, can remove this entire IF block
    If UseExtendedVersion Then
        z_Sc(280) = &H8BE58960: z_Sc(281) = &H8BFC245D: z_Sc(282) = &H7B8B2855: z_Sc(283) = &HFFC6E858: z_Sc(284) = &HC085FFFF: z_Sc(285) = &H7D890374: z_Sc(286) = &H8C2611C: z_Sc(287) = &H768D00: z_Sc(288) = &H245C8B53: z_Sc(289) = &H5C438B08: z_Sc(290) = &H4C25B: z_Sc(291) = &HFFFFFFB8: z_Sc(292) = &HE58960FF: z_Sc(293) = &H8D245D8B: z_Sc(294) = &H79E82855: z_Sc(295) = &H83000002: z_Sc(296) = &H7400147D: z_Sc(297) = &H7D83FC6F: z_Sc(298) = &H3E740314: z_Sc(299) = &H6474C985
        z_Sc(300) = &H7B8BD689: z_Sc(301) = &H21EE858: z_Sc(302) = &HC0850000: z_Sc(303) = &H7B891274: z_Sc(304) = &H4EE8358: z_Sc(305) = &H518E148B: z_Sc(306) = &HFFFE63E8: z_Sc(307) = &HF4E259FF: z_Sc(308) = &H3147D81: z_Sc(309) = &H74000020: z_Sc(310) = &H4C6832A: z_Sc(311) = &HC931F789: z_Sc(312) = &H1F3E8: z_Sc(313) = &H587B8B00: z_Sc(314) = &H7B8B19EB: z_Sc(315) = &H1B958: z_Sc(316) = &HE1E80000: z_Sc(317) = &H85000001: z_Sc(318) = &H891974C0: z_Sc(319) = &H2DE8587B: z_Sc(320) = &H8BFFFFFE: z_Sc(321) = &H1C4D890F: z_Sc(322) = &H875C985: z_Sc(323) = &H1C7E8: z_Sc(324) = &H587B8900: z_Sc(325) = &H14C261: z_Sc(326) = &HFFFFFFB8: z_Sc(327) = &HE58960FF: z_Sc(328) = &H83245D8B: z_Sc(329) = &H7400587B
        z_Sc(330) = &H28558D75: z_Sc(331) = &H1E7E8: z_Sc(332) = &H147D8300: z_Sc(333) = &HFC677400: z_Sc(334) = &H3147D83: z_Sc(335) = &HC9853C74: z_Sc(336) = &HD6895C74: z_Sc(337) = &H83587B8B: z_Sc(338) = &H45C704EE: z_Sc(339) = &H1C&: z_Sc(340) = &H768D00: z_Sc(341) = &H518E148B: z_Sc(342) = &HFFFD83E8: z_Sc(343) = &HF4E259FF: z_Sc(344) = &H3147D81: z_Sc(345) = &H74000020: z_Sc(346) = &H4C6832C: z_Sc(347) = &HC931F789: z_Sc(348) = &H163E8: z_Sc(349) = &H587B8B00: z_Sc(350) = &H45C71BEB: z_Sc(351) = &H1C&: z_Sc(352) = &HFFFA8300: z_Sc(353) = &H7DE80775: z_Sc(354) = &HEBFFFFFE: z_Sc(355) = &H587B8B11: z_Sc(356) = &HFFFD4BE8: z_Sc(357) = &H74FF85FF: z_Sc(358) = &H890F8B05: z_Sc(359) = &HC2611C4D
        z_Sc(360) = &HF6890014: z_Sc(361) = &H245C8B53: z_Sc(362) = &H5C438B08: z_Sc(363) = &HB74C085: z_Sc(364) = &HFF50006A: z_Sc(365) = &HC0313853: z_Sc(366) = &H5B5C4389: z_Sc(367) = &H900004C2: z_Sc(368) = &HFFFFFFB8: z_Sc(369) = &HE58960FF: z_Sc(370) = &H8D245D8B: z_Sc(371) = &H45E82C55: z_Sc(372) = &H83000001: z_Sc(373) = &HF00147D: z_Sc(374) = &HA084&: z_Sc(375) = &H147D8300: z_Sc(376) = &H85087403: z_Sc(377) = &H92840FC9: z_Sc(378) = &H31000000: z_Sc(379) = &H5C43F7C0: z_Sc(381) = &H458B0D75: z_Sc(382) = &H74C08528: z_Sc(383) = &H1B805: z_Sc(384) = &HF7400000: z_Sc(385) = &H20001445: z_Sc(386) = &H48740000: z_Sc(387) = &HC289D689: z_Sc(388) = &H415C7B8B: z_Sc(389) = &HBFE8&
        z_Sc(390) = &H74C08500: z_Sc(391) = &H5C7B8923: z_Sc(392) = &H1078B49: z_Sc(393) = &H890789C8: z_Sc(394) = &HD2851C45: z_Sc(395) = &H57890374: z_Sc(396) = &HC829FC04: z_Sc(397) = &HC102C083: z_Sc(398) = &HC70102E0: z_Sc(399) = &H5EA5F356: z_Sc(400) = &H3147D81: z_Sc(401) = &H74000020: z_Sc(402) = &H31F78934: z_Sc(403) = &H86E8C9: z_Sc(404) = &H29EB0000: z_Sc(405) = &H2B9C689: z_Sc(406) = &H8B000000: z_Sc(407) = &H75E85C7B: z_Sc(408) = &H85000000: z_Sc(409) = &H891674C0: z_Sc(410) = &HF8B5C7B: z_Sc(411) = &H850F8941: z_Sc(412) = &H890374F6: z_Sc(413) = &H54890477: z_Sc(414) = &H4D89048F: z_Sc(415) = &H18C2611C: z_Sc(416) = &H768D00: z_Sc(417) = &H8960C031: z_Sc(418) = &H245D8BE5: z_Sc(419) = &H8558738B
        z_Sc(420) = &H8B2874F6: z_Sc(421) = &HC0852845: z_Sc(422) = &H7752178: z_Sc(423) = &H4D890E8B: z_Sc(424) = &H8B18EB1C: z_Sc(425) = &HFF852C7D: z_Sc(426) = &HE8B1174: z_Sc(427) = &H27DC839: z_Sc(428) = &H89FCC189: z_Sc(429) = &HC6831C4D: z_Sc(430) = &H61A5F304: z_Sc(431) = &H90000CC2: z_Sc(432) = &H778B0F8B: z_Sc(433) = &H28458B04: z_Sc(434) = &HF208C783: z_Sc(435) = &H830574AF: z_Sc(436) = &H3EB02F6: z_Sc(437) = &HC301F683: z_Sc(438) = &H5152C031: z_Sc(439) = &HD75C985: z_Sc(440) = &H3174FF85: z_Sc(441) = &H53FF5751: z_Sc(442) = &HEBFF3138: z_Sc(443) = &HFF854128: z_Sc(444) = &H178B1474: z_Sc(445) = &HE1C1D101: z_Sc(446) = &HFF575102: z_Sc(447) = &HC0853853: z_Sc(448) = &HC7891374: z_Sc(449) = &HE1C10FEB
        z_Sc(450) = &H53FF5102: z_Sc(451) = &H74C08534: z_Sc(452) = &H89388904: z_Sc(453) = &HC35A59C7: z_Sc(454) = &H8FE8&: z_Sc(455) = &H14458900: z_Sc(456) = &H840FC085: z_Sc(457) = &H82&: z_Sc(458) = &H7403F883: z_Sc(459) = &H74D2857D: z_Sc(460) = &H66D68979: z_Sc(461) = &H8366068B: z_Sc(462) = &H6C7501F8: z_Sc(463) = &H85104E8B: z_Sc(464) = &H816774C9: z_Sc(465) = &HFFF9&: z_Sc(466) = &H8B5D7F00: z_Sc(467) = &H768B0446: z_Sc(468) = &H74F6850C: z_Sc(469) = &H10F88353: z_Sc(470) = &HF2890474: z_Sc(471) = &H31FC4CEB: z_Sc(472) = &HFF72E8FF: z_Sc(473) = &HC085FFFF: z_Sc(474) = &HD2313E74: z_Sc(475) = &H66068B66: z_Sc(476) = &H75FFFCA9: z_Sc(477) = &HF883661E: z_Sc(478) = &H660F7403: z_Sc(479) = &H7502F883
        z_Sc(480) = &H468B6612: z_Sc(481) = &HC0BF0F08: z_Sc(482) = &H468B03EB: z_Sc(483) = &H17048908: z_Sc(484) = &H8304C283: z_Sc(485) = &HD4E210C6: z_Sc(486) = &HE9C1D189: z_Sc(487) = &H8904E302: z_Sc(488) = &HE807EBFA: z_Sc(489) = &HFFFFFF30: z_Sc(490) = &H90C3C931: z_Sc(491) = &H8B66C031: z_Sc(492) = &H4000A902: z_Sc(493) = &H5750000: z_Sc(494) = &HEB08C283: z_Sc(495) = &H4000350D: z_Sc(496) = &H528B0000: z_Sc(497) = &HCF88308: z_Sc(498) = &HC3DE274: z_Sc(499) = &H74000020: z_Sc(500) = &H20033D22: z_Sc(501) = &H1B740000: z_Sc(502) = &H7403F883: z_Sc(503) = &H2F88316: z_Sc(504) = &HC0310474: z_Sc(505) = &H8B660FEB: z_Sc(506) = &HD2BF0F12: z_Sc(507) = &H3B8&: z_Sc(508) = &H8B02EB00: z_Sc(509) = &HC312

        z_Sc(lThunkSize + 13) = ((hThunk Xor &H80000000) + &H460) Xor &H80000000 ' IsWindowManaged
        z_Sc(lThunkSize + 14) = ((hThunk Xor &H80000000) + &H480) Xor &H80000000 ' IsFiltered
        z_Sc(lThunkSize + 15) = ((hThunk Xor &H80000000) + &H48C) Xor &H80000000 ' AddWindow
        z_Sc(lThunkSize + 16) = ((hThunk Xor &H80000000) + &H518) Xor &H80000000 ' RemoveWindow
        z_Sc(lThunkSize + 17) = ((hThunk Xor &H80000000) + &H5A4) Xor &H80000000 ' RemoveFilter
        z_Sc(lThunkSize + 18) = ((hThunk Xor &H80000000) + &H5C0) Xor &H80000000 ' AddFilterMessage
        z_Sc(lThunkSize + 19) = ((hThunk Xor &H80000000) + &H684) Xor &H80000000 ' GetWindows
    End If
    thunkCopyMemory ByVal hThunk, z_Sc(0), (lThunkSize + lTableSize) * 4
    
    ' /// create the VB object's VTable
    ReDim z_Sc(DATA_SIZE - 1)               ' update VTable location
    z_Sc(0) = ((hThunk Xor &H80000000) + lThunkSize * 4) Xor &H80000000
    z_Sc(1) = 1                             ' ref counter
    z_Sc(2) = 0                             ' recursion counter
    z_Sc(3) = ObjPtr(VbHost)               ' passed VB host
    z_Sc(4) = FunctionOrdinal               ' host function pointer
    z_Sc(5) = hThunk                        ' this thunk address
    z_Sc(6) = 0                             ' thunk tag
    z_Sc(7) = 0                             ' Dispatch::IUnknown
    z_Sc(8) = hTLBstructs                   ' TypeLib structs
    z_Sc(9) = lEbMode                       ' optional IDE safety EbMode API or null
    z_Sc(10) = thunkGetProcAddress(hMod("k32"), ByVal "VirtualFree")
    z_Sc(11) = thunkGetProcAddress(hMod("u32"), ByVal "SetTimer")
    z_Sc(12) = thunkGetProcAddress(hMod("u32"), ByVal "KillTimer")
    z_Sc(13) = thunkGetProcAddress(hMod("o32"), ByVal "CoTaskMemAlloc")
    z_Sc(14) = thunkGetProcAddress(hMod("o32"), ByVal "CoTaskMemRealloc")
    z_Sc(15) = thunkGetProcAddress(hMod("k32"), ByVal "FreeLibrary")
    z_Sc(16) = hMod("SetWindowSubclass")
    z_Sc(17) = hMod("RemoveWindowSubclass")
    z_Sc(18) = hMod("DefSubclassProc")
    z_Sc(19) = hMod("GetWindowSubclass")
    z_Sc(20) = thunkGetProcAddress(hMod("u32"), ByVal "DefWindowProcA")
    z_Sc(21) = thunkGetProcAddress(hMod("u32"), ByVal "DefWindowProcW")
    z_Sc(22) = 0                            ' Client list
    If UseExtendedVersion = False Then
        z_Sc(23) = InitHwnd                 ' Filter
        z_Sc(24) = &H81 ' state: 0=normal,1=base version,2=paused,4=unloading,128=initializing
    End If
    z_Sc(25) = hMod("cc32")                 ' common controls hMod reference, released with FreeLibrary
    thunkCopyMemory ByVal hData, z_Sc(0), DATA_SIZE * 4
    thunkCopyMemory oIUnk, hData, 4
    Erase z_Sc()
    Set hMod = Nothing
    
    Debug.Print "data section: "; hData
    
    ' create the adhoc IDispatch. Enables VB to recognize this as an object
    thunkCreateStdDispatch hData, hData, ObjPtr(oITypeInfo), VarPtr(oIDispatch)
    Set oITypeInfo = Nothing
    
    Debug.Assert (oIDispatch <> 0)
    If oIDispatch <> 0 Then                 ' if failure, we unwind
        ' copy the IDispatch pointer to the object data & initialize subclassing if applies
        thunkCopyMemory ByVal ((hData Xor &H80000000) + 28) Xor &H80000000, oIDispatch, 4
        Set oObj = oIUnk
        If UseExtendedVersion = False Then  ' unloads thunk if failure
            oObj.resume
            If oObj.State <> 0 Then Set CreateTasker_Subclass = oObj
        Else
            Set CreateTasker_Subclass = oObj
            oObj.AddWindow InitHwnd
        End If
        Set oObj = Nothing
    End If
    Set oIUnk = Nothing ' will self-delete thunk & memory allocations if routine failed
    Exit Function
    
SetAPIreferences:       ' modules used by the thunk
    n = thunkLoadLibrary("comctl32.dll")
    Debug.Assert n <> 0                 ' gotta have this for subclassing
    If n = 0 Then GoTo AbortCleanup
    hMod.Add n, "cc32"
    n = thunkGetProcAddress(hMod("cc32"), ByVal "SetWindowSubclass")
    If n = 0 Then                   ' try by ordinal
        n = thunkGetProcAddress(hMod("cc32"), ByVal 410&)
        Debug.Assert n <> 0         ' can't load common controls subclassing library!!!
        If n = 0 Then GoTo AbortCleanup
        hMod.Add n, "SetWindowSubclass"
        hMod.Add thunkGetProcAddress(hMod("cc32"), ByVal 411&), "GetWindowSubclass"
        hMod.Add thunkGetProcAddress(hMod("cc32"), ByVal 412&), "RemoveWindowSubclass"
        hMod.Add thunkGetProcAddress(hMod("cc32"), ByVal 413&), "DefSubclassProc"
    Else
        hMod.Add n, "SetWindowSubclass"
        hMod.Add thunkGetProcAddress(hMod("cc32"), ByVal "GetWindowSubclass"), "GetWindowSubclass"
        hMod.Add thunkGetProcAddress(hMod("cc32"), ByVal "RemoveWindowSubclass"), "RemoveWindowSubclass"
        hMod.Add thunkGetProcAddress(hMod("cc32"), ByVal "DefSubclassProc"), "DefSubclassProc"
    End If
    hMod.Add thunkGetModuleHandle("user32.dll"), "u32"
    hMod.Add thunkGetModuleHandle("kernel32.dll"), "k32"
    hMod.Add thunkGetModuleHandle("ole32.dll"), "o32"
    If AddIDEsafety = True Then
        On Error Resume Next
        Debug.Print 1 / 0
        If Err Then
            Err.Clear: On Error GoTo 0
            hMod.Add thunkGetModuleHandle("vba6.dll"), "vba"
            ' note: if next line errors, change vba6.dll to vba5.dll
            lEbMode = thunkGetProcAddress(hMod("vba"), ByVal "EbMode")
        End If
        On Error GoTo 0
    End If
    Return
    
AbortCleanup:
    If hData <> 0 Then thunkCoTaskMemFree hData
    If hTLBstructs <> 0 Then
        Set oITypeInfo = Nothing
        thunkCoTaskMemFree hTLBstructs
    End If
    On Error Resume Next
    n = hMod("cc32")
    If Err Then Err.Clear Else thunkFreeLibrary n
    On Error GoTo 0
    
End Function

Public Function CreateTasker_CustomClass(VbHost As Object, _
                                    ByVal FunctionOrdinal As Long, _
                                    Optional ByVal UseUnicode As Boolean = False, _
                                    Optional UseExtendedVersion As Boolean = False, _
                                    Optional AddIDEsafety As Boolean = True) As Object
                                    
    ' returned object has these methods (non-extended version)
    ' Method Name       Sample Call             Purpose/Notes on next line
    ' ----------------------------------------------------------------------------
    ' AddressOf         lAddr=oTasker.AddressOf
    '                   returns thunk address for RegisterClassEx
    ' Atom              oTasker.Atom=return value of RegisterClassEx: Print oTasker.Atom
    '                   returns/sets class atom assigned by Windows
    '                   passing Atom performs UnregisterClass API call when thunk unloads
    ' Pause             no parameters. oTasker.Pause
    '                   temporarily stops sending messages to your project
    ' NextInChain       lRet=oTasker.NextInChain(hWnd, uMsg, wParam, lParam)
    '                   passes message to next in subclass chain
    '                   function should only be used inside subclass procedure as needed
    ' Resume            no parameters. oTasker.Resume
    '                   resumes sending messages to your project
    ' State             Default method, no parameters. Print oTasker.State
    '                   returns -1 if currently paused
    '                       ## else number of actively subclassed hWnds
    ' Tag               oTasker.Tag=1908: Print oTasker.Tag
    
    ' Extended version only
    ' Method Name       Sample Call             Purpose/Notes on next line
    ' ----------------------------------------------------------------------------
    ' AddFilterMessage  oTasker.AddFilterMessage 0|1, Array(msg1, msg2): oTasker.AddFilterMessage 0, WM_KILLFOCUS
    '                   appends messages to the thunk filter.
    '                   returns number of filter messages or -1 if invalid parameter
    '                   pass last parameter as single message, Array(...), or Long array
    '                   pass 1st parameter: 0 = inclusive filter else non-zero = exclusive list
    '                   inclusive: only forward messages in the list
    '                   exclusive: only forward messages not in the list
    ' IsFiltered        Print oTasker.IsFiltered
    '                   returns non-zero if a filter exists
    ' IsWindowManaged   oTasker.IsWindowManaged(Me.hWnd)
    '                   returns subclass-list pointer if passed hWnd is subclassed
    ' RemoveFilter      no parameters. Resets thunk filter, always returns zero
    ' RemoveWindow      oTasker.RemoveWindow -1: oTasker.RemoveWindow VarPtr(Array(Text1.hWnd, Me.Hwnd))
    '                   removes subclassing for list of hWnds
    '                   returns number of subclassed windows or -1 if invalid parameter
    '                   pass last parameter as single message, Array(...), or Long array
    '                   passing -1 unsubclasses all hWnds
    ' TagHwnd           oTasker(Me.hWnd)=1234: Print oTasker(Me.Hwnd)
    '                   returns/sets a Long value tag for a specific cusotm class hWnd
    
    ' this function's parameters
    '   VbHost :: the host that this thunk object is declared in
    '   FunctionOrdinal :: last private method in host is #1, second to last is #2, etc
    '   UseUnicode :: sets whether ANSI or wide-character SetWindowsHookEx is used
    '   UseExtendedVersion :: add the extra methods to the object or don't
    '   AddIDEsafety :: ignored unless in IDE
    
    Dim hTLBstructs As Long, z_Sc() As Long
    Dim hThunk As Long, hData As Long
    Dim oIDispatch As Long, lEbMode As Long
    Dim lThunkSize As Long, lTableSize As Long
    Dim oIUnk As stdole.IUnknown
    Dim oITypeInfo As stdole.IUnknown
    Dim hMod As New Collection
    
    Const THUNK_SIZE As Long = 267          ' size of the thunk array (in longs)
    Const VTable_SIZE As Long = 12          ' nr of VTable entries (in longs)
    Const THUNK_SIZE_EX As Long = 480       ' size of the thunk array (in longs)
    Const VTable_SIZE_EX As Long = 20       ' nr of VTable entries (in longs)
    Const DATA_SIZE As Long = 24            ' Object data size (in longs)
    
    FunctionOrdinal = pvGetAddressOf(VbHost, FunctionOrdinal)
    Debug.Assert FunctionOrdinal <> 0
    ' if invalid function index was passed or VbHost is Nothing then abort
    If FunctionOrdinal = 0 Then Exit Function
    
    ' call local routine to create a temporary typelib for adhoc IDispatch interface
    ' 1=function, 2=propGet, 4=propLet
    If UseExtendedVersion = False Then
        lThunkSize = THUNK_SIZE: lTableSize = VTable_SIZE
        Set oITypeInfo = pvCreateITypeInfo(hTLBstructs, _
                        1, "State", 1, "AddressOf", 2, "Atom", 4, "Atom", "Value", _
                        2, "Tag", 4, "Tag", "Value", 1, "Pause", 1, "Resume", _
                        1, "NextInChain", "hWnd", "uMsg", "wParam", "lParam")
    Else
        lThunkSize = THUNK_SIZE_EX: lTableSize = VTable_SIZE_EX
        Set oITypeInfo = pvCreateITypeInfo(hTLBstructs, _
                        1, "State", 1, "AddressOf", 2, "Atom", 4, "Atom", "Value", _
                        2, "Tag", 4, "Tag", "Value", 1, "Pause", 1, "Resume", _
                        1, "NextInChain", "hWnd", "uMsg", "wParam", "lParam", _
                        1, "IsWindowManaged", "hWnd", 1, "IsFiltered", _
                        1, "RemoveWindow", "!pItems", 1, "RemoveFilter", _
                        2, "TagHwnd", "hWnd", 4, "TagHwnd", "hWnd", "Value", _
                        1, "AddFilterMessage", "fType", "!pItems", _
                        1, "GetWindows", "pCount", "pItems")
    End If
    Debug.Assert (ObjPtr(oITypeInfo) <> 0)
    If oITypeInfo Is Nothing Then GoTo AbortCleanup
    
    hData = thunkCoTaskMemAlloc(DATA_SIZE * 4) ' memory for our ObjPtr
    Debug.Assert hData <> 0
    If hData = 0 Then GoTo AbortCleanup
    
    hThunk = thunkVirtualAlloc(0, (lThunkSize + lTableSize) * 4)
    Debug.Assert hThunk <> 0
    If hThunk = 0 Then GoTo AbortCleanup    ' thunk creation failed
   
    Debug.Print "thunk: "; hThunk
                                        
    GoSub SetAPIreferences  ' Gosub? Just wanted to declutter top half of routine
    
    ReDim z_Sc(0 To lThunkSize + lTableSize - 1) ' the thunk
    z_Sc(0) = &HBB60D231: z_Sc(1) = &H12345678: z_Sc(2) = &H5DE9E589: z_Sc(3) = &H90000001: z_Sc(4) = &HBB60C031: z_Sc(5) = &H12345678: z_Sc(6) = &HA9E9E589: z_Sc(7) = &H90000001: z_Sc(8) = &H8BE58960: z_Sc(9) = &H458B245D: z_Sc(10) = &H4488B28: z_Sc(11) = &H8108480B: z_Sc(12) = &HC0F9&: z_Sc(13) = &H81377500: z_Sc(14) = &HC78&: z_Sc(15) = &H2E754600: z_Sc(16) = &HC985088B: z_Sc(17) = &H558B0C75: z_Sc(18) = &HFF1A892C: z_Sc(19) = &HC0310443: z_Sc(20) = &H438B21EB: z_Sc(21) = &H74C0851C: z_Sc(22) = &HF98115: z_Sc(23) = &H75000204: z_Sc(24) = &H2C75FF0D: z_Sc(25) = &H502875FF: z_Sc(26) = &H10FF008B: z_Sc(27) = &H2B805EB: z_Sc(28) = &H89800040: z_Sc(29) = &HC2611C45
    z_Sc(30) = &HF689000C: z_Sc(31) = &H424548B: z_Sc(32) = &H4004428B: z_Sc(33) = &HC2044289: z_Sc(34) = &HF6890004: z_Sc(35) = &H424548B: z_Sc(36) = &H4804428B: z_Sc(37) = &H85044289: z_Sc(38) = &H600C75C0: z_Sc(39) = &H5D8BE589: z_Sc(40) = &H25AE824: z_Sc(41) = &HC2610000: z_Sc(42) = &HF6890004: z_Sc(43) = &H245C8B53: z_Sc(44) = &H50438B08: z_Sc(45) = &H6A9&: z_Sc(46) = &H83057400: z_Sc(47) = &H9EBFFC8: z_Sc(48) = &H8548438B: z_Sc(49) = &H8B0274C0: z_Sc(50) = &H4C25B00: z_Sc(51) = &H768D00: z_Sc(52) = &H245C8B53: z_Sc(53) = &H14438B08: z_Sc(54) = &H4C25B: z_Sc(55) = &H245C8B53: z_Sc(56) = &H54438B08: z_Sc(57) = &H4C25B: z_Sc(58) = &H245C8B53: z_Sc(59) = &H24448B08
    z_Sc(60) = &H5443890C: z_Sc(61) = &HC25BC031: z_Sc(62) = &HF6890008: z_Sc(63) = &H245C8B53: z_Sc(64) = &H18438B08: z_Sc(65) = &H4C25B: z_Sc(66) = &H245C8B53: z_Sc(67) = &H24448B08: z_Sc(68) = &H1843890C: z_Sc(69) = &HC25BC031: z_Sc(70) = &HF6890008: z_Sc(71) = &H245C8B53: z_Sc(72) = &H50438B08: z_Sc(73) = &H8902C883: z_Sc(74) = &HC25B5043: z_Sc(75) = &HF6890004: z_Sc(76) = &H8BE58960: z_Sc(77) = &H438B245D: z_Sc(78) = &H4A950: z_Sc(79) = &H6750000: z_Sc(80) = &H8302C883: z_Sc(81) = &H438902F0: z_Sc(82) = &H1C458950: z_Sc(83) = &H4C261: z_Sc(84) = &H53E58955: z_Sc(85) = &HFF085D8B: z_Sc(86) = &H75FF1875: z_Sc(87) = &H1075FF14: z_Sc(88) = &HFF0C75FF: z_Sc(89) = &H5D5B3C53
    z_Sc(90) = &H900014C2: z_Sc(91) = &HFF1C5589: z_Sc(92) = &HB9E80843: z_Sc(93) = &H85000000: z_Sc(94) = &HE83475F6: z_Sc(95) = &H1E0&: z_Sc(96) = &H501C458D: z_Sc(97) = &H5014458D: z_Sc(98) = &HFF1875FF: z_Sc(99) = &H75FF3075: z_Sc(100) = &H2875FF2C: z_Sc(101) = &HFF2475FF: z_Sc(102) = &H53FF0C73: z_Sc(103) = &H18458B10: z_Sc(104) = &H674C085: z_Sc(105) = &HFF008B50: z_Sc(106) = &H75390850: z_Sc(107) = &HFF127514: z_Sc(108) = &H75FF3075: z_Sc(109) = &H2875FF2C: z_Sc(110) = &HFF2475FF: z_Sc(111) = &H45893C53: z_Sc(112) = &H84BFF1C: z_Sc(113) = &H10C261: z_Sc(114) = &H87B83: z_Sc(115) = &H7B831075: z_Sc(116) = &HE740048: z_Sc(117) = &H1C7E8: z_Sc(118) = &H16AE800: z_Sc(119) = &HC2610000
    z_Sc(120) = &H75FF0010: z_Sc(121) = &H2475FF2C: z_Sc(122) = &H313053FF: z_Sc(123) = &H207B8BC9: z_Sc(124) = &H1F7E8: z_Sc(125) = &H54438B00: z_Sc(126) = &HA74C085: z_Sc(127) = &HFFC0B70F: z_Sc(128) = &HFF505873: z_Sc(129) = &H438B4453: z_Sc(130) = &H14458928: z_Sc(131) = &H6A147B8B: z_Sc(132) = &H53FF5300: z_Sc(133) = &H3045C738: z_Sc(134) = &H8000&: z_Sc(135) = &H2C45C7: z_Sc(136) = &H89000000: z_Sc(137) = &H5861287D: z_Sc(138) = &HFF240489: z_Sc(139) = &H768DE2: z_Sc(140) = &H8BDE89FC: z_Sc(141) = &H6A95043: z_Sc(142) = &H75000000: z_Sc(143) = &H8BF63122: z_Sc(144) = &HFF854C7B: z_Sc(145) = &H3DE80974: z_Sc(146) = &H85000004: z_Sc(147) = &H8B1075F6: z_Sc(148) = &HC0852443: z_Sc(149) = &HD0FF0974
    z_Sc(150) = &H7401F883: z_Sc(151) = &H81DE8902: z_Sc(152) = &H82287D: z_Sc(153) = &H4A750000: z_Sc(154) = &H7B8BDE89: z_Sc(155) = &H24558B48: z_Sc(156) = &H15BE8: z_Sc(157) = &H74C08500: z_Sc(158) = &H490F8B78: z_Sc(159) = &HF891BE3: z_Sc(160) = &H48F548B: z_Sc(161) = &H8B871489: z_Sc(162) = &HFF855C7B: z_Sc(163) = &H548B6374: z_Sc(164) = &H1489048F: z_Sc(165) = &HEB0F8987: z_Sc(166) = &H14EE858: z_Sc(167) = &H7B890000: z_Sc(168) = &H5C7B8B48: z_Sc(169) = &H4974FF85: z_Sc(170) = &H13FE8: z_Sc(171) = &H5C7B8900: z_Sc(172) = &H7D833FEB: z_Sc(173) = &H39750128: z_Sc(174) = &H83487B8B: z_Sc(175) = &H29E802C1: z_Sc(176) = &H85000001: z_Sc(177) = &H892A74C0: z_Sc(178) = &HF8B487B: z_Sc(179) = &H8B0F8941
    z_Sc(180) = &H14892455: z_Sc(181) = &H5C7B8B8F: z_Sc(182) = &H1574FF85: z_Sc(183) = &H10BE8: z_Sc(184) = &H74C08500: z_Sc(185) = &H5C7B890C: z_Sc(186) = &H4C70F89: z_Sc(187) = &H8F&: z_Sc(188) = &H5043F700: z_Sc(189) = &H4&: z_Sc(190) = &H49E80574: z_Sc(191) = &HC3000000: z_Sc(192) = &H4504B83: z_Sc(193) = &HC43C7: z_Sc(194) = &H8B000000: z_Sc(195) = &HFF854C7B: z_Sc(196) = &HC9310A74: z_Sc(197) = &HD3E8&: z_Sc(198) = &H4C7B8900: z_Sc(199) = &H855C7B8B: z_Sc(200) = &H310A74FF: z_Sc(201) = &HC2E8C9: z_Sc(202) = &H7B890000: z_Sc(203) = &H6EE85C: z_Sc(204) = &H438B0000: z_Sc(205) = &H74C0851C: z_Sc(206) = &H8B5006: z_Sc(207) = &HE80850FF: z_Sc(208) = &H4&: z_Sc(209) = &H768DC3
    z_Sc(210) = &H8314438B: z_Sc(211) = &H685010C0: z_Sc(212) = &H1F4&: z_Sc(213) = &H6A0C73FF: z_Sc(214) = &H2C53FF00: z_Sc(215) = &HC30C4389: z_Sc(216) = &HC710EC83: z_Sc(217) = &H4002404: z_Sc(218) = &H44C70002: z_Sc(219) = &H424&: z_Sc(220) = &H44C70000: z_Sc(221) = &HC00824: z_Sc(222) = &H44C70000: z_Sc(223) = &HC24&: z_Sc(224) = &H148D4600: z_Sc(225) = &H1845C724: z_Sc(227) = &H8D18458D: z_Sc(228) = &H8B525108: z_Sc(229) = &H8B501C43: z_Sc(230) = &H8310FF00: z_Sc(231) = &H90C310C4: z_Sc(232) = &H85487B8B: z_Sc(233) = &HFC1574FF: z_Sc(234) = &H8B487B8B: z_Sc(235) = &H768D0F: z_Sc(236) = &H518F148B: z_Sc(237) = &H7E8&: z_Sc(238) = &HF4E25900: z_Sc(239) = &H768DC3
    z_Sc(240) = &HA74D285: z_Sc(241) = &H6A006A: z_Sc(242) = &HFF52106A: z_Sc(243) = &H90C34053: z_Sc(244) = &HF74FF85: z_Sc(245) = &H51570F8B: z_Sc(246) = &HC783D089: z_Sc(247) = &H58AFF204: z_Sc(248) = &H3103745F: z_Sc(249) = &HC829C3C0: z_Sc(250) = &H768DC3: z_Sc(251) = &H5152C031: z_Sc(252) = &HD75C985: z_Sc(253) = &H3174FF85: z_Sc(254) = &H53FF5751: z_Sc(255) = &HEBFF3138: z_Sc(256) = &HFF854128: z_Sc(257) = &H178B1474: z_Sc(258) = &HE1C1D101: z_Sc(259) = &HFF575102: z_Sc(260) = &HC0853853: z_Sc(261) = &HC7891374: z_Sc(262) = &HE1C10FEB: z_Sc(263) = &H53FF5102: z_Sc(264) = &H74C08534: z_Sc(265) = &H89388904: z_Sc(266) = &HC35A59C7

    '/// patch thunk & build VTable -- same order as sent to pvCreateITypeInfo
    z_Sc(1) = hData: z_Sc(5) = hData        ' 2 non-VTable functions, pass objptr to them
    z_Sc(lThunkSize + 0) = ((hThunk Xor &H80000000) + &H20) Xor &H80000000 ' IUnknown:QueryInterface(QI)
    z_Sc(lThunkSize + 1) = ((hThunk Xor &H80000000) + &H7C) Xor &H80000000 ' IUnknown:AddRef
    z_Sc(lThunkSize + 2) = ((hThunk Xor &H80000000) + &H8C) Xor &H80000000 ' IUnknown:Release
    z_Sc(lThunkSize + 3) = ((hThunk Xor &H80000000) + &HAC) Xor &H80000000 ' State
    z_Sc(lThunkSize + 4) = ((hThunk Xor &H80000000) + &HD0) Xor &H80000000 ' AddressOf
    z_Sc(lThunkSize + 5) = ((hThunk Xor &H80000000) + &HDC) Xor &H80000000 ' Atom Get
    z_Sc(lThunkSize + 6) = ((hThunk Xor &H80000000) + &HE8) Xor &H80000000 ' Atom Let
    z_Sc(lThunkSize + 7) = ((hThunk Xor &H80000000) + &HFC) Xor &H80000000 ' Tag Get
    z_Sc(lThunkSize + 8) = ((hThunk Xor &H80000000) + &H108) Xor &H80000000 ' Tag Let
    z_Sc(lThunkSize + 9) = ((hThunk Xor &H80000000) + &H11C) Xor &H80000000 ' Pause
    z_Sc(lThunkSize + 10) = ((hThunk Xor &H80000000) + &H130) Xor &H80000000 ' Resume
    z_Sc(lThunkSize + 11) = ((hThunk Xor &H80000000) + &H150) Xor &H80000000 ' NextInChain
    
    '/// if extended version never wanted, can remove this entire IF block
    If UseExtendedVersion Then
        z_Sc(lThunkSize + 12) = ((hThunk Xor &H80000000) + &H42C) Xor &H80000000 ' IsWindowManaged
        z_Sc(lThunkSize + 13) = ((hThunk Xor &H80000000) + &H44C) Xor &H80000000 ' IsFiltered
        z_Sc(lThunkSize + 14) = ((hThunk Xor &H80000000) + &H458) Xor &H80000000 ' RemoveWindow
        z_Sc(lThunkSize + 15) = ((hThunk Xor &H80000000) + &H4F0) Xor &H80000000 ' RemoveFilter
        z_Sc(lThunkSize + 16) = ((hThunk Xor &H80000000) + &H50C) Xor &H80000000 ' TagHwnd Get
        z_Sc(lThunkSize + 17) = ((hThunk Xor &H80000000) + &H538) Xor &H80000000 ' TagHwnd Let
        z_Sc(lThunkSize + 18) = ((hThunk Xor &H80000000) + &H588) Xor &H80000000 ' AddFilterMessage
        z_Sc(lThunkSize + 19) = ((hThunk Xor &H80000000) + &H64C) Xor &H80000000 ' GetWindows
        
        z_Sc(267) = &H8BE58960: z_Sc(268) = &H8BFC245D: z_Sc(269) = &H7B8B2855
        z_Sc(270) = &HFF92E848: z_Sc(271) = &HC085FFFF: z_Sc(272) = &H7D890374: z_Sc(273) = &H8C2611C: z_Sc(274) = &H768D00: z_Sc(275) = &H245C8B53: z_Sc(276) = &H4C438B08: z_Sc(277) = &H4C25B: z_Sc(278) = &HFFFFFFB8: z_Sc(279) = &HE58960FF: z_Sc(280) = &H83245D8B: z_Sc(281) = &H7400487B: z_Sc(282) = &H28558D7E: z_Sc(283) = &H22FE8: z_Sc(284) = &H147D8300: z_Sc(285) = &HFC707400: z_Sc(286) = &H3147D83: z_Sc(287) = &HC9854574: z_Sc(288) = &HD6896574: z_Sc(289) = &H83487B8B: z_Sc(290) = &H45C704EE: z_Sc(291) = &H1C&: z_Sc(292) = &H768D00: z_Sc(293) = &H518E148B: z_Sc(294) = &HFFFF33E8: z_Sc(295) = &H74C085FF: z_Sc(296) = &HFF1AE805: z_Sc(297) = &HE259FFFF: z_Sc(298) = &H147D81EB: z_Sc(299) = &H2003&
        z_Sc(300) = &HC6833574: z_Sc(301) = &H31F78904: z_Sc(302) = &HFF2EE8C9: z_Sc(303) = &H7B8BFFFF: z_Sc(304) = &HC724EB48: z_Sc(305) = &H1C45&: z_Sc(306) = &HFA830000: z_Sc(307) = &HE80775FF: z_Sc(308) = &HFFFFFECC: z_Sc(309) = &H7B8B11EB: z_Sc(310) = &HFEF2E848: z_Sc(311) = &HC085FFFF: z_Sc(312) = &HD9E80574: z_Sc(313) = &H89FFFFFE: z_Sc(314) = &HC2611C7D: z_Sc(315) = &HF6890014: z_Sc(316) = &H245C8B53: z_Sc(317) = &H4C438B08: z_Sc(318) = &HB74C085: z_Sc(319) = &HFF50006A: z_Sc(320) = &HC0313853: z_Sc(321) = &H5B4C4389: z_Sc(322) = &H900004C2: z_Sc(323) = &H8960C031: z_Sc(324) = &H245D8BE5: z_Sc(325) = &H487B8BFC: z_Sc(326) = &HE828558B: z_Sc(327) = &HFFFFFEB0: z_Sc(328) = &HD74C085: z_Sc(329) = &H855C7B8B
        z_Sc(330) = &H8B0674FF: z_Sc(331) = &H45898704: z_Sc(332) = &H8C2611C: z_Sc(333) = &H768D00: z_Sc(334) = &H8960C031: z_Sc(335) = &H245D8BE5: z_Sc(336) = &H487B8BFC: z_Sc(337) = &HE828558B: z_Sc(338) = &HFFFFFE84: z_Sc(339) = &H3274C085: z_Sc(340) = &H855C7B8B: z_Sc(341) = &H8B2275FF: z_Sc(342) = &HA8B4853: z_Sc(343) = &H89E85041: z_Sc(344) = &H85FFFFFE: z_Sc(345) = &H1A7458C0: z_Sc(346) = &H495C7B89: z_Sc(347) = &HF6890F89: z_Sc(348) = &H8F04C7: z_Sc(349) = &HE2000000: z_Sc(350) = &H1C7D89F7: z_Sc(351) = &H892C558B: z_Sc(352) = &HC2618714: z_Sc(353) = &HF689000C: z_Sc(354) = &HFFFFFFB8: z_Sc(355) = &HE58960FF: z_Sc(356) = &H8D245D8B: z_Sc(357) = &H5E82C55: z_Sc(358) = &H83000001: z_Sc(359) = &HF00147D
        z_Sc(360) = &HA084&: z_Sc(361) = &H147D8300: z_Sc(362) = &H85087403: z_Sc(363) = &H92840FC9: z_Sc(364) = &H31000000: z_Sc(365) = &H4C43F7C0: z_Sc(367) = &H458B0D75: z_Sc(368) = &H74C08528: z_Sc(369) = &H1B805: z_Sc(370) = &HF7400000: z_Sc(371) = &H20001445: z_Sc(372) = &H48740000: z_Sc(373) = &HC289D689: z_Sc(374) = &H414C7B8B: z_Sc(375) = &HFFFE0BE8: z_Sc(376) = &H74C085FF: z_Sc(377) = &H4C7B8923: z_Sc(378) = &H1078B49: z_Sc(379) = &H890789C8: z_Sc(380) = &HD2851C45: z_Sc(381) = &H57890374: z_Sc(382) = &HC829FC04: z_Sc(383) = &HC102C083: z_Sc(384) = &HC70102E0: z_Sc(385) = &H5EA5F356: z_Sc(386) = &H3147D81: z_Sc(387) = &H74000020: z_Sc(388) = &H31F78934: z_Sc(389) = &HFDD2E8C9
        z_Sc(390) = &H29EBFFFF: z_Sc(391) = &H2B9C689: z_Sc(392) = &H8B000000: z_Sc(393) = &HC1E84C7B: z_Sc(394) = &H85FFFFFD: z_Sc(395) = &H891674C0: z_Sc(396) = &HF8B4C7B: z_Sc(397) = &H850F8941: z_Sc(398) = &H890374F6: z_Sc(399) = &H54890477: z_Sc(400) = &H4D89048F: z_Sc(401) = &H18C2611C: z_Sc(402) = &H768D00: z_Sc(403) = &H8960C031: z_Sc(404) = &H245D8BE5: z_Sc(405) = &H8548738B: z_Sc(406) = &H8B2874F6: z_Sc(407) = &HC0852845: z_Sc(408) = &H7752178: z_Sc(409) = &H4D890E8B: z_Sc(410) = &H8B18EB1C: z_Sc(411) = &HFF852C7D: z_Sc(412) = &HE8B1174: z_Sc(413) = &H27DC839: z_Sc(414) = &H89FCC189: z_Sc(415) = &HC6831C4D: z_Sc(416) = &H61A5F304: z_Sc(417) = &H90000CC2: z_Sc(418) = &H778B0F8B: z_Sc(419) = &H28458B04
        z_Sc(420) = &HF208C783: z_Sc(421) = &H830574AF: z_Sc(422) = &H3EB02F6: z_Sc(423) = &HC301F683: z_Sc(424) = &H8FE8&: z_Sc(425) = &H14458900: z_Sc(426) = &H840FC085: z_Sc(427) = &H82&: z_Sc(428) = &H7403F883: z_Sc(429) = &H74D2857D: z_Sc(430) = &H66D68979: z_Sc(431) = &H8366068B: z_Sc(432) = &H6C7501F8: z_Sc(433) = &H85104E8B: z_Sc(434) = &H816774C9: z_Sc(435) = &HFFF9&: z_Sc(436) = &H8B5D7F00: z_Sc(437) = &H768B0446: z_Sc(438) = &H74F6850C: z_Sc(439) = &H10F88353: z_Sc(440) = &HF2890474: z_Sc(441) = &HFF314CEB: z_Sc(442) = &HFFFCFFE8: z_Sc(443) = &H74C085FF: z_Sc(444) = &HFCD2313F: z_Sc(445) = &H66068B66: z_Sc(446) = &H75FFFCA9: z_Sc(447) = &HF883661E: z_Sc(448) = &H660F7403: z_Sc(449) = &H7502F883
        z_Sc(450) = &H468B6612: z_Sc(451) = &HC0BF0F08: z_Sc(452) = &H468B03EB: z_Sc(453) = &H17048908: z_Sc(454) = &H8304C283: z_Sc(455) = &HD4E210C6: z_Sc(456) = &HE9C1D189: z_Sc(457) = &H8904E302: z_Sc(458) = &HE807EBFA: z_Sc(459) = &HFFFFFCBC: z_Sc(460) = &H90C3C931: z_Sc(461) = &H8B66C031: z_Sc(462) = &H4000A902: z_Sc(463) = &H5750000: z_Sc(464) = &HEB08C283: z_Sc(465) = &H4000350D: z_Sc(466) = &H528B0000: z_Sc(467) = &HCF88308: z_Sc(468) = &HC3DE274: z_Sc(469) = &H74000020: z_Sc(470) = &H20033D22: z_Sc(471) = &H1B740000: z_Sc(472) = &H7403F883: z_Sc(473) = &H2F88316: z_Sc(474) = &HC0310474: z_Sc(475) = &H8B660FEB: z_Sc(476) = &HD2BF0F12: z_Sc(477) = &H3B8&: z_Sc(478) = &H8B02EB00: z_Sc(479) = &HC312
    End If
    thunkCopyMemory ByVal hThunk, z_Sc(0), (lThunkSize + lTableSize) * 4
    Erase z_Sc()
    
    ' /// create the VB object's VTable
    ReDim z_Sc(DATA_SIZE - 1)               ' update VTable location
    z_Sc(0) = ((hThunk Xor &H80000000) + lThunkSize * 4) Xor &H80000000
    z_Sc(1) = 1                             ' ref counter
    z_Sc(2) = 0                             ' recursion counter
    z_Sc(3) = ObjPtr(VbHost)               ' passed VB host
    z_Sc(4) = FunctionOrdinal               ' host function pointer
    z_Sc(5) = hThunk                        ' this thunk address
    z_Sc(6) = 0                             ' thunk tag
    z_Sc(7) = 0                             ' Dispatch::IUnknown
    z_Sc(8) = hTLBstructs                   ' TypeLib structs
    z_Sc(9) = lEbMode                       ' optional IDE safety EbMode API or null
    z_Sc(10) = thunkGetProcAddress(hMod("k32"), ByVal "VirtualFree")
    z_Sc(11) = thunkGetProcAddress(hMod("u32"), ByVal "SetTimer")
    z_Sc(12) = thunkGetProcAddress(hMod("u32"), ByVal "KillTimer")
    z_Sc(13) = thunkGetProcAddress(hMod("o32"), ByVal "CoTaskMemAlloc")
    z_Sc(14) = thunkGetProcAddress(hMod("o32"), ByVal "CoTaskMemRealloc")
    If UseUnicode = True Then
        z_Sc(15) = thunkGetProcAddress(hMod("u32"), ByVal "DefWindowProcW")
    Else
        z_Sc(15) = thunkGetProcAddress(hMod("u32"), ByVal "DefWindowProcA")
    End If
    z_Sc(16) = thunkGetProcAddress(hMod("u32"), ByVal "PostMessageA")
    z_Sc(17) = thunkGetProcAddress(hMod("u32"), ByVal "UnregisterClassW")
    z_Sc(18) = 0                            ' Client List
    z_Sc(19) = 0                            ' Filter
    If UseExtendedVersion = False Then z_Sc(20) = 1 ' state: 0=normal,1=base version,2=paused,4=unloading
    z_Sc(21) = 0                            ' Atom
    z_Sc(22) = App.hInstance                ' App.hInstance
    z_Sc(23) = 0                            ' hwnd Tags list
    thunkCopyMemory ByVal hData, z_Sc(0), DATA_SIZE * 4
    thunkCopyMemory oIUnk, hData, 4
    Erase z_Sc()
    Set hMod = Nothing
    
    Debug.Print "data section: "; hData
    
    ' create the adhoc IDispatch. Enables VB to recognize this as an object
    thunkCreateStdDispatch hData, hData, ObjPtr(oITypeInfo), VarPtr(oIDispatch)
    Set oITypeInfo = Nothing
    
    Debug.Assert (oIDispatch <> 0)
    If oIDispatch <> 0 Then                 ' if failure, we unwind
        ' copy the IDispatch pointer to the object data
        thunkCopyMemory ByVal ((hData Xor &H80000000) + 28) Xor &H80000000, oIDispatch, 4
        Set CreateTasker_CustomClass = oIUnk
    End If
    Set oIUnk = Nothing ' will self-delete thunk & memory allocations if routine failed
    Exit Function
    
SetAPIreferences:   ' modules used by the thunk
    hMod.Add thunkGetModuleHandle("user32.dll"), "u32"
    hMod.Add thunkGetModuleHandle("kernel32.dll"), "k32"
    hMod.Add thunkGetModuleHandle("ole32.dll"), "o32"
    If AddIDEsafety = True Then
        On Error Resume Next
        Debug.Print 1 / 0
        If Err Then
            Err.Clear: On Error GoTo 0
            hMod.Add thunkGetModuleHandle("vba6.dll"), "vba"
            ' note: if next line errors, change vba6.dll to vba5.dll
            lEbMode = thunkGetProcAddress(hMod("vba"), ByVal "EbMode")
        End If
        On Error GoTo 0
    End If
    Return

AbortCleanup:
    If hData <> 0 Then thunkCoTaskMemFree hData
    If hTLBstructs <> 0 Then
        Set oITypeInfo = Nothing
        thunkCoTaskMemFree hTLBstructs
    End If
    
End Function

Public Function CreateTasker_Hook(VbHost As Object, _
                                    ByVal FunctionOrdinal As Long, _
                                    ByVal HookType As Long, _
                                    Optional ByVal UseUnicode As Boolean = False, _
                                    Optional UseExtendedVersion As Boolean = False, _
                                    Optional AddIDEsafety As Boolean = True) As Object
                                    
    ' returned object has these methods
    ' Method Name       Sample Call             Purpose/Notes on next line
    ' ----------------------------------------------------------------------------
    ' HookType          Print oTasker.HookType
    '                   returns hook type, whether hook is running or not
    ' NextInChain       lRtn=oTasker.NextInChain(nCode,wParam,lParam)
    '                   Passes message to next in hook chain
    '                   function should only be used inside hook procedure as needed
    ' Pause             oTasker.Pause
    '                   Uninstalls hook
    ' Resume            oTasker.Resume
    '                   Reinstalls hook, current hook type
    ' State             Default method, no parameters. Print oTasker.State
    '                   Returns non-zero if hook is running
    ' Tag               oTasker.Tag=1908: Print oTasker.Tag
    '                   User-defined thunk tag (Long values only)
    
    ' Extended version only
    ' Method Name       Sample Call             Purpose/Notes on next line
    ' ----------------------------------------------------------------------------
    ' AddFilterMessage  oTasker.AddFilterMessage 0|1, 1|2|3, Array(msg1, msg2): oTasker.AddFilterMessage 0, 2, WM_KILLFOCUS
    '                   appends messages to the thunk filter.
    '                   returns number of filter messages or -1 if invalid parameter
    '                   pass last parameter as single message, Array(...), or Long array
    '                   pass middle parameter: 1=nCode filter, 2=wParam filter, 3=lParam filter
    '                   pass 1st parameter: 0 = inclusive filter else non-zero = exclusive list
    '                   inclusive: only forward messages in the list
    '                   exclusive: only forward messages not in the list
    ' IsFiltered        no parameters. Print oTasker.IsFiltered
    ' RemoveFilter      no parameters. Resets thunk filter, always returns zero
    
    ' this function's parameters
    '   VbHost :: the host that this thunk object is declared in
    '   FunctionOrdinal :: last private method in host is #1, second to last is #2, etc
    '   HookType :: required, type of hook. See SetWindowsHookEx for constants & details
    '   UseUnicode :: sets whether ANSI or wide-character SetWindowsHookEx is used
    '   UseExtendedVersion :: add the extra methods to the object or don't
    '   AddIDEsafety :: ignored unless in IDE
    
    Dim z_Sc() As Long, hTLBstructs As Long
    Dim lThunkSize As Long, lTableSize As Long
    Dim hThunk As Long, hData As Long
    Dim oIDispatch As Long, lEbMode As Long
    Dim oIUnk As stdole.IUnknown, oObj As Object
    Dim oITypeInfo As stdole.IUnknown
    Dim hMod As New Collection
    
    Const THUNK_SIZE As Long = 172          ' size of the thunk array (in longs)
    Const VTable_SIZE As Long = 10          ' nr of VTable entries (in longs)
    Const THUNK_SIZE_EX As Long = 324       ' size of the thunk array (in longs)
    Const VTable_SIZE_EX As Long = 13       ' nr of VTable entries (in longs)
    Const DATA_SIZE As Long = 22            ' Object data size (in longs)
    
    FunctionOrdinal = pvGetAddressOf(VbHost, FunctionOrdinal)
    Debug.Assert FunctionOrdinal <> 0
    ' if invalid function index was passed or VbHost is Nothing then abort
    If FunctionOrdinal = 0 Then Exit Function
    
    ' call local routine to create a temporary typelib for adhoc IDispatch interface
    ' 1=function, 2=propGet, 4=propLet
    If UseExtendedVersion = False Then
        lThunkSize = THUNK_SIZE: lTableSize = VTable_SIZE
        Set oITypeInfo = pvCreateITypeInfo(hTLBstructs, _
                        1, "State", 2, "Tag", 4, "Tag", "Value", 1, "Pause", 1, "Resume", _
                        2, "HookType", 1, "NextInChain", "nCode", "wParam", "lParam")
    Else
        lThunkSize = THUNK_SIZE_EX: lTableSize = VTable_SIZE_EX
        Set oITypeInfo = pvCreateITypeInfo(hTLBstructs, _
                        1, "State", 2, "Tag", 4, "Tag", "Value", 1, "Pause", 1, "Resume", _
                        2, "HookType", 1, "NextInChain", "nCode", "wParam", "lParam", _
                        1, "IsFiltered", 1, "RemoveFilter", _
                        1, "AddFilterMessage", "fType", "pOrdinal", "!pItems")
    End If
    Debug.Assert (ObjPtr(oITypeInfo) <> 0)
    If oITypeInfo Is Nothing Then GoTo AbortCleanup
    
    hData = thunkCoTaskMemAlloc(DATA_SIZE * 4) ' memory for our ObjPtr
    Debug.Assert hData <> 0
    If hData = 0 Then GoTo AbortCleanup
    
    hThunk = thunkVirtualAlloc(0, (lThunkSize + lTableSize) * 4)
    Debug.Assert hThunk <> 0
    If hThunk = 0 Then GoTo AbortCleanup    ' thunk creation failed
    
    Debug.Print "thunk: "; hThunk
    
    GoSub SetAPIreferences  ' Gosub? Just wanted to declutter top half of routine
    
    ReDim z_Sc(0 To lThunkSize + lTableSize - 1) ' the thunk
    z_Sc(0) = &HBB60D231: z_Sc(1) = &H12345678: z_Sc(2) = &H15E9E589: z_Sc(3) = &H90000001: z_Sc(4) = &HBB60C031: z_Sc(5) = &H12345678: z_Sc(6) = &H89E9E589: z_Sc(7) = &H90000001: z_Sc(8) = &H8BE58960: z_Sc(9) = &H458B245D: z_Sc(10) = &H4488B28: z_Sc(11) = &H8108480B: z_Sc(12) = &HC0F9&: z_Sc(13) = &H81377500: z_Sc(14) = &HC78&: z_Sc(15) = &H2E754600: z_Sc(16) = &HC985088B: z_Sc(17) = &H558B0C75: z_Sc(18) = &HFF1A892C: z_Sc(19) = &HC0310443: z_Sc(20) = &H438B21EB: z_Sc(21) = &H74C0851C: z_Sc(22) = &HF98115: z_Sc(23) = &H75000204: z_Sc(24) = &H2C75FF0D: z_Sc(25) = &H502875FF: z_Sc(26) = &H10FF008B: z_Sc(27) = &H2B805EB: z_Sc(28) = &H89800040: z_Sc(29) = &HC2611C45
    z_Sc(30) = &HF689000C: z_Sc(31) = &H424548B: z_Sc(32) = &H4004428B: z_Sc(33) = &HC2044289: z_Sc(34) = &HF6890004: z_Sc(35) = &H424548B: z_Sc(36) = &H4804428B: z_Sc(37) = &H85044289: z_Sc(38) = &H600C75C0: z_Sc(39) = &H5D8BE589: z_Sc(40) = &H192E824: z_Sc(41) = &HC2610000: z_Sc(42) = &HF6890004: z_Sc(43) = &H245C8B53: z_Sc(44) = &H4C438B08: z_Sc(45) = &H4C25B: z_Sc(46) = &H245C8B53: z_Sc(47) = &H18438B08: z_Sc(48) = &H4C25B: z_Sc(49) = &H245C8B53: z_Sc(50) = &H24448B08: z_Sc(51) = &H1843890C: z_Sc(52) = &HC25BC031: z_Sc(53) = &HF6890008: z_Sc(54) = &H245C8B53: z_Sc(55) = &H122E808: z_Sc(56) = &HC25B0000: z_Sc(57) = &HF6890004: z_Sc(58) = &H245C8B53: z_Sc(59) = &H126E808
    z_Sc(60) = &HC25B0000: z_Sc(61) = &HF6890004: z_Sc(62) = &H245C8B53: z_Sc(63) = &H48438B08: z_Sc(64) = &H4C25B: z_Sc(65) = &H53E58955: z_Sc(66) = &H31085D8B: z_Sc(67) = &H4C4B8BC0: z_Sc(68) = &H75FF0DE3: z_Sc(69) = &H1075FF14: z_Sc(70) = &H510C75FF: z_Sc(71) = &H5B4053FF: z_Sc(72) = &H10C25D: z_Sc(73) = &H7589F631: z_Sc(74) = &H843FF1C: z_Sc(75) = &H75107339: z_Sc(76) = &HB6E807: z_Sc(77) = &H56EB0000: z_Sc(78) = &H8524458B: z_Sc(79) = &H8B4F78C0: z_Sc(80) = &HFF85547B: z_Sc(81) = &H71E80974: z_Sc(82) = &H85000002: z_Sc(83) = &H8B3F75F6: z_Sc(84) = &HC0852443: z_Sc(85) = &HD0FF0774: z_Sc(86) = &H7501F883: z_Sc(87) = &H10EE831: z_Sc(88) = &H458D0000: z_Sc(89) = &H458D501C
    z_Sc(90) = &H75FF5014: z_Sc(91) = &H2C75FF18: z_Sc(92) = &HFF2875FF: z_Sc(93) = &H73FF2475: z_Sc(94) = &H1053FF0C: z_Sc(95) = &H8518458B: z_Sc(96) = &H500674C0: z_Sc(97) = &H50FF008B: z_Sc(98) = &H14753908: z_Sc(99) = &H75FF1275: z_Sc(100) = &H2875FF2C: z_Sc(101) = &HFF2475FF: z_Sc(102) = &H53FF4C73: z_Sc(103) = &H1C458940: z_Sc(104) = &H61084BFF: z_Sc(105) = &H90000CC2: z_Sc(106) = &H87B83: z_Sc(107) = &HC2610474: z_Sc(108) = &H75FF0010: z_Sc(109) = &H2475FF2C: z_Sc(110) = &H6A3053FF: z_Sc(111) = &H2073FF00: z_Sc(112) = &H8B3853FF: z_Sc(113) = &H45892843: z_Sc(114) = &H147B8B14: z_Sc(115) = &HFF53006A: z_Sc(116) = &H45C73853: z_Sc(117) = &H800030: z_Sc(118) = &H2C45C700
    z_Sc(120) = &H61287D89: z_Sc(121) = &H24048958: z_Sc(122) = &HF689E2FF: z_Sc(123) = &H8314438B: z_Sc(124) = &H6A5010C0: z_Sc(125) = &HC73FF64: z_Sc(126) = &H53FF006A: z_Sc(127) = &HC43892C: z_Sc(128) = &H768DC3: z_Sc(129) = &H854C438B: z_Sc(130) = &H500B74C0: z_Sc(131) = &HC74453FF: z_Sc(132) = &H4C43&: z_Sc(133) = &H90C30000: z_Sc(134) = &H4B8BC031: z_Sc(135) = &H8B16E348: z_Sc(136) = &HCA854C53: z_Sc(137) = &H73FF0F75: z_Sc(138) = &HFF006A50: z_Sc(139) = &HFF511473: z_Sc(140) = &H43893C53: z_Sc(141) = &HF689C34C: z_Sc(142) = &H1043C7: z_Sc(143) = &HC7000000: z_Sc(144) = &HC43&: z_Sc(145) = &HB9E80000: z_Sc(146) = &H8BFFFFFF: z_Sc(147) = &HFF85547B: z_Sc(148) = &HC9310A74: z_Sc(149) = &H197E8
    z_Sc(150) = &H547B8900: z_Sc(151) = &H851C438B: z_Sc(152) = &H500674C0: z_Sc(153) = &H50FF008B: z_Sc(154) = &HFF7EE808: z_Sc(155) = &H90C3FFFF: z_Sc(156) = &HC710EC83: z_Sc(157) = &H4002404: z_Sc(158) = &H44C70002: z_Sc(159) = &H424&: z_Sc(160) = &H44C70000: z_Sc(161) = &HC00824: z_Sc(162) = &H44C70000: z_Sc(163) = &HC24&: z_Sc(164) = &H148D4600: z_Sc(165) = &H1845C724: z_Sc(167) = &H8D18458D: z_Sc(168) = &H8B525108: z_Sc(169) = &H8B501C43: z_Sc(170) = &H8310FF00: z_Sc(171) = &H90C310C4

    '/// patch thunk & build VTable -- same order as sent to pvCreateITypeInfo
    z_Sc(1) = hData: z_Sc(5) = hData        ' 2 non-VTable functions, pass objptr to them
    z_Sc(lThunkSize + 0) = ((hThunk Xor &H80000000) + &H20) Xor &H80000000 ' IUnknown:QueryInterface(QI)
    z_Sc(lThunkSize + 1) = ((hThunk Xor &H80000000) + &H7C) Xor &H80000000 ' IUnknown:AddRef
    z_Sc(lThunkSize + 2) = ((hThunk Xor &H80000000) + &H8C) Xor &H80000000 ' IUnknown:Release
    z_Sc(lThunkSize + 3) = ((hThunk Xor &H80000000) + &HAC) Xor &H80000000 ' State
    z_Sc(lThunkSize + 4) = ((hThunk Xor &H80000000) + &HB8) Xor &H80000000 ' Tag Get
    z_Sc(lThunkSize + 5) = ((hThunk Xor &H80000000) + &HC4) Xor &H80000000 ' Tag Let
    z_Sc(lThunkSize + 6) = ((hThunk Xor &H80000000) + &HD8) Xor &H80000000 ' Pause
    z_Sc(lThunkSize + 7) = ((hThunk Xor &H80000000) + &HE8) Xor &H80000000 ' Resume
    z_Sc(lThunkSize + 8) = ((hThunk Xor &H80000000) + &HF8) Xor &H80000000 ' HookType
    z_Sc(lThunkSize + 9) = ((hThunk Xor &H80000000) + &H104) Xor &H80000000 ' NextInChain
    
    '/// if extended version never wanted, can remove this entire IF block
    If UseExtendedVersion = True Then
        z_Sc(172) = &H245C8B53: z_Sc(173) = &H54438B08: z_Sc(174) = &H4C25B: z_Sc(175) = &H245C8B53: z_Sc(176) = &H54438B08: z_Sc(177) = &HB74C085: z_Sc(178) = &HFF50006A: z_Sc(179) = &HC0313853
        z_Sc(180) = &H5B544389: z_Sc(181) = &H900004C2: z_Sc(182) = &HFFFFFFB8: z_Sc(183) = &HE58960FF: z_Sc(184) = &H8B245D8B: z_Sc(185) = &HC0852C45: z_Sc(186) = &HC7840F: z_Sc(187) = &HFCA90000: z_Sc(188) = &HFFFFFFF: z_Sc(189) = &HBC85&: z_Sc(190) = &H30558D00: z_Sc(191) = &H12FE8: z_Sc(192) = &H147D8300: z_Sc(193) = &HAA840F00: z_Sc(194) = &H83000000: z_Sc(195) = &H7403147D: z_Sc(196) = &HFC98508: z_Sc(197) = &H9C84&: z_Sc(198) = &HF7C03100: z_Sc(199) = &H5443&: z_Sc(200) = &H17750000: z_Sc(201) = &HC12C458B: z_Sc(202) = &H8B5104E0: z_Sc(203) = &HC985284D: z_Sc(204) = &H1B90574: z_Sc(205) = &H41000000: z_Sc(206) = &HFC59C809: z_Sc(207) = &H1445F7: z_Sc(208) = &H74000020: z_Sc(209) = &H89D68947
        z_Sc(210) = &H547B8BC2: z_Sc(211) = &H9EE841: z_Sc(212) = &HC0850000: z_Sc(213) = &H7B892274: z_Sc(214) = &H78B4954: z_Sc(215) = &H789C801: z_Sc(216) = &H851C4589: z_Sc(217) = &H890374D2: z_Sc(218) = &HC8290457: z_Sc(219) = &HC102C083: z_Sc(220) = &HC70102E0: z_Sc(221) = &H5EA5F356: z_Sc(222) = &H3147D81: z_Sc(223) = &H74000020: z_Sc(224) = &H31F78934: z_Sc(225) = &H66E8C9: z_Sc(226) = &H29EB0000: z_Sc(227) = &H2B9C689: z_Sc(228) = &H8B000000: z_Sc(229) = &H55E8547B: z_Sc(230) = &H85000000: z_Sc(231) = &H891674C0: z_Sc(232) = &HF8B547B: z_Sc(233) = &H850F8941: z_Sc(234) = &H890374F6: z_Sc(235) = &H54890477: z_Sc(236) = &H4D89048F: z_Sc(237) = &H1CC2611C: z_Sc(238) = &H768D00: z_Sc(239) = &H778B0F8B
        z_Sc(240) = &H20C6F704: z_Sc(241) = &H74000000: z_Sc(242) = &H30FE830F: z_Sc(243) = &H458B057D: z_Sc(244) = &H8B08EB28: z_Sc(245) = &H3EB2C45: z_Sc(246) = &HFC24458B: z_Sc(247) = &H8303E683: z_Sc(248) = &HAFF208C7: z_Sc(249) = &HF6830574: z_Sc(250) = &H8303EB02: z_Sc(251) = &H90C301F6: z_Sc(252) = &H5152C031: z_Sc(253) = &HD75C985: z_Sc(254) = &H3174FF85: z_Sc(255) = &H53FF5751: z_Sc(256) = &HEBFF3138: z_Sc(257) = &HFF854128: z_Sc(258) = &H178B1474: z_Sc(259) = &HE1C1D101: z_Sc(260) = &HFF575102: z_Sc(261) = &HC0853853: z_Sc(262) = &HC7891374: z_Sc(263) = &HE1C10FEB: z_Sc(264) = &H53FF5102: z_Sc(265) = &H74C08534: z_Sc(266) = &H89388904: z_Sc(267) = &HC35A59C7: z_Sc(268) = &H8FE8&: z_Sc(269) = &H14458900
        z_Sc(270) = &H840FC085: z_Sc(271) = &H82&: z_Sc(272) = &H7403F883: z_Sc(273) = &H74D2857D: z_Sc(274) = &H66D68979: z_Sc(275) = &H8366068B: z_Sc(276) = &H6C7501F8: z_Sc(277) = &H85104E8B: z_Sc(278) = &H816774C9: z_Sc(279) = &HFFF9&: z_Sc(280) = &H8B5D7F00: z_Sc(281) = &H768B0446: z_Sc(282) = &H74F6850C: z_Sc(283) = &H10F88353: z_Sc(284) = &HF2890474: z_Sc(285) = &HFF314CEB: z_Sc(286) = &HFFFF73E8: z_Sc(287) = &H74C085FF: z_Sc(288) = &HFCD2313F: z_Sc(289) = &H66068B66: z_Sc(290) = &H75FFFCA9: z_Sc(291) = &HF883661E: z_Sc(292) = &H660F7403: z_Sc(293) = &H7502F883: z_Sc(294) = &H468B6612: z_Sc(295) = &HC0BF0F08: z_Sc(296) = &H468B03EB: z_Sc(297) = &H17048908: z_Sc(298) = &H8304C283: z_Sc(299) = &HD4E210C6
        z_Sc(300) = &HE9C1D189: z_Sc(301) = &H8904E302: z_Sc(302) = &HE807EBFA: z_Sc(303) = &HFFFFFF30: z_Sc(304) = &H90C3C931: z_Sc(305) = &H8B66C031: z_Sc(306) = &H4000A902: z_Sc(307) = &H5750000: z_Sc(308) = &HEB08C283: z_Sc(309) = &H4000350D: z_Sc(310) = &H528B0000: z_Sc(311) = &HCF88308: z_Sc(312) = &HC3DE274: z_Sc(313) = &H74000020: z_Sc(314) = &H20033D22: z_Sc(315) = &H1B740000: z_Sc(316) = &H7403F883: z_Sc(317) = &H2F88316: z_Sc(318) = &HC0310474: z_Sc(319) = &H8B660FEB: z_Sc(320) = &HD2BF0F12: z_Sc(321) = &H3B8&: z_Sc(322) = &H8B02EB00: z_Sc(323) = &HC312

        z_Sc(lThunkSize + 10) = ((hThunk Xor &H80000000) + &H2B0) Xor &H80000000 ' IsFiltered
        z_Sc(lThunkSize + 11) = ((hThunk Xor &H80000000) + &H2BD) Xor &H80000000 ' RemoveFilter
        z_Sc(lThunkSize + 12) = ((hThunk Xor &H80000000) + &H2D8) Xor &H80000000 ' AddFilterMessage
    End If
    thunkCopyMemory ByVal hThunk, z_Sc(0), (lThunkSize + lTableSize) * 4
    
    ' /// create the VB object's VTable
    ReDim z_Sc(DATA_SIZE - 1)               ' update VTable location
    z_Sc(0) = ((hThunk Xor &H80000000) + lThunkSize * 4) Xor &H80000000
    z_Sc(1) = 1                             ' ref counter
    z_Sc(2) = 0                             ' recursion counter
    z_Sc(3) = ObjPtr(VbHost)               ' passed VB host
    z_Sc(4) = FunctionOrdinal               ' host function pointer
    z_Sc(5) = hThunk                        ' this thunk address
    z_Sc(6) = 0                             ' thunk tag
    z_Sc(7) = 0                             ' Dispatch::IUnknown
    z_Sc(8) = hTLBstructs                   ' TypeLib structs
    z_Sc(9) = lEbMode                       ' optional IDE safety EbMode API or null
    z_Sc(10) = thunkGetProcAddress(hMod("k32"), ByVal "VirtualFree")
    z_Sc(11) = thunkGetProcAddress(hMod("u32"), ByVal "SetTimer")
    z_Sc(12) = thunkGetProcAddress(hMod("u32"), ByVal "KillTimer")
    z_Sc(13) = thunkGetProcAddress(hMod("o32"), ByVal "CoTaskMemAlloc")
    z_Sc(14) = thunkGetProcAddress(hMod("o32"), ByVal "CoTaskMemRealloc")
    If UseUnicode = True Then
        z_Sc(15) = thunkGetProcAddress(hMod("u32"), ByVal "SetWindowsHookExW")
    Else
        z_Sc(15) = thunkGetProcAddress(hMod("u32"), ByVal "SetWindowsHookExA")
    End If
    z_Sc(16) = thunkGetProcAddress(hMod("u32"), ByVal "CallNextHookEx")
    z_Sc(17) = thunkGetProcAddress(hMod("u32"), ByVal "UnhookWindowsHookEx")
    z_Sc(18) = HookType                     ' Hook type
    z_Sc(19) = 0                            ' Hook ID
    z_Sc(20) = App.ThreadID                 ' App.ThreadID
    z_Sc(21) = 0                            ' Filter
    
    thunkCopyMemory ByVal hData, z_Sc(0), DATA_SIZE * 4
    thunkCopyMemory oIUnk, hData, 4
    Erase z_Sc()
    Set hMod = Nothing
    
    Debug.Print "data section: "; hData
    
    ' create the adhoc IDispatch. Enables VB to recognize this as an object
    thunkCreateStdDispatch hData, hData, ObjPtr(oITypeInfo), VarPtr(oIDispatch)
    Set oITypeInfo = Nothing
    
    Debug.Assert (oIDispatch <> 0)
    If oIDispatch <> 0 Then                 ' if failure, we unwind
        ' copy the IDispatch pointer to the object data & initialize the hook
        thunkCopyMemory ByVal ((hData Xor &H80000000) + 28) Xor &H80000000, oIDispatch, 4
        Set oObj = oIUnk
        oObj.resume
        If oObj.State <> 0 Then Set CreateTasker_Hook = oObj
        Set oObj = Nothing
    End If
    Set oIUnk = Nothing ' will self-delete thunk & memory allocations if routine failed
    Exit Function
    
SetAPIreferences:       ' modules used by the thunk
    hMod.Add thunkGetModuleHandle("user32.dll"), "u32"
    hMod.Add thunkGetModuleHandle("kernel32.dll"), "k32"
    hMod.Add thunkGetModuleHandle("ole32.dll"), "o32"
    If AddIDEsafety = True Then
        On Error Resume Next
        Debug.Print 1 / 0
        If Err Then
            Err.Clear: On Error GoTo 0
            hMod.Add thunkGetModuleHandle("vba6.dll"), "vba"
            ' note: if next line errors, change vba6.dll to vba5.dll
            lEbMode = thunkGetProcAddress(hMod("vba"), ByVal "EbMode")
        End If
        On Error GoTo 0
    End If
    Return
    
AbortCleanup:
    If hData <> 0 Then thunkCoTaskMemFree hData
    If hTLBstructs <> 0 Then
        Set oITypeInfo = Nothing
        thunkCoTaskMemFree hTLBstructs
    End If
    
End Function

Public Function CreateTasker_Timer(VbHost As Object, _
                                    ByVal FunctionOrdinal As Long, _
                                    Optional ByVal InitInterval As Long = 0, _
                                    Optional AddIDEsafety As Boolean = True) As Object

    ' returned object has these methods
    ' Method Name   Sample Call             Purpose/Notes on next line
    ' ----------------------------------------------------------------------------
    ' Interval      oTasker.Interval=500: Print oTasker.Interval
    '               returns/sets user-defined timer interval
    '               note: values < 1 are valid intervals, see SetTimer documentation if needed
    ' Pause         oTasker.Pause
    '               stops timer
    ' Resume        oTasker.Resume
    '               starts timer using current Interval property
    ' State         Print oTasker.State
    '               returns non-zero if timer is running (Default method)
    ' Tag           oTasker.Tag=1908: Print oTasker.Tag
    '               returns/sets user-defined thunk tag (Long values only)
    
    ' this function's parameters
    '   VbHost :: the host that this thunk object is declared in
    '   FunctionOrdinal :: last private method in host is #1, second to last is #2, etc
    '   AddIDEsafety :: ignored unless in IDE
    '   InitInterval :: pass non-zero to auto-start the timer when function returns
    
    Dim z_Sc() As Long
    Dim hThunk As Long, hData As Long
    Dim oIDispatch As Long, lEbMode As Long
    Dim hTLBstructs As Long, oObj As Object
    Dim oIUnk As stdole.IUnknown
    Dim oITypeInfo As stdole.IUnknown
    Dim hMod As New Collection
    
    Const THUNK_SIZE As Long = 150          ' size of the thunk array (in longs)
    Const VTable_SIZE As Long = 10          ' nr of VTable entries (in longs)
    Const DATA_SIZE As Long = 16            ' Object data size (in longs)
    
    FunctionOrdinal = pvGetAddressOf(VbHost, FunctionOrdinal)
    Debug.Assert FunctionOrdinal <> 0
    ' if invalid function index was passed or VbHost is Nothing then abort
    If FunctionOrdinal = 0 Then Exit Function
    
    ' call local routine to create a temporary typelib for adhoc IDispatch interface
    ' 1=function, 2=propGet, 4=propLet
    Set oITypeInfo = pvCreateITypeInfo(hTLBstructs, _
                        1, "State", 2, "Tag", 4, "Tag", "Value", _
                        1, "Pause", 1, "Resume", 2, "Interval", 4, "Interval", "Value")
    Debug.Assert (ObjPtr(oITypeInfo) <> 0)
    If oITypeInfo Is Nothing Then GoTo AbortCleanup
    
    hData = thunkCoTaskMemAlloc(DATA_SIZE * 4) ' memory for our ObjPtr
    Debug.Assert hData <> 0
    If hData = 0 Then GoTo AbortCleanup
    
    hThunk = thunkVirtualAlloc(0, (THUNK_SIZE + VTable_SIZE) * 4)
    Debug.Assert hThunk <> 0
    If hThunk = 0 Then GoTo AbortCleanup        ' thunk creation failed
    
    Debug.Print "thunk: "; hThunk
                                        
    GoSub SetAPIreferences  ' Gosub? Just wanted to declutter top half of routine
    
    ReDim z_Sc(0 To THUNK_SIZE + VTable_SIZE - 1) ' the thunk
    z_Sc(0) = &HBB60D231: z_Sc(1) = &H12345678: z_Sc(2) = &H3DE9E589: z_Sc(3) = &H90000001: z_Sc(4) = &HBB60C031: z_Sc(5) = &H12345678: z_Sc(6) = &H79E9E589: z_Sc(7) = &H90000001: z_Sc(8) = &H8BE58960: z_Sc(9) = &H458B245D: z_Sc(10) = &H4488B28: z_Sc(11) = &H8108480B: z_Sc(12) = &HC0F9&: z_Sc(13) = &H81377500: z_Sc(14) = &HC78&: z_Sc(15) = &H2E754600: z_Sc(16) = &HC985088B: z_Sc(17) = &H558B0C75: z_Sc(18) = &HFF1A892C: z_Sc(19) = &HC0310443: z_Sc(20) = &H438B21EB: z_Sc(21) = &H74C0851C: z_Sc(22) = &HF98115: z_Sc(23) = &H75000204: z_Sc(24) = &H2C75FF0D: z_Sc(25) = &H502875FF: z_Sc(26) = &H10FF008B: z_Sc(27) = &H2B805EB: z_Sc(28) = &H89800040: z_Sc(29) = &HC2611C45
    z_Sc(30) = &HF689000C: z_Sc(31) = &H424548B: z_Sc(32) = &H4004428B: z_Sc(33) = &HC2044289: z_Sc(34) = &HF6890004: z_Sc(35) = &H424548B: z_Sc(36) = &H4804428B: z_Sc(37) = &H85044289: z_Sc(38) = &H600C75C0: z_Sc(39) = &H5D8BE589: z_Sc(40) = &H132E824: z_Sc(41) = &HC2610000: z_Sc(42) = &HF6890004: z_Sc(43) = &H245C8B53: z_Sc(44) = &H38438B08: z_Sc(45) = &H4C25B: z_Sc(46) = &H245C8B53: z_Sc(47) = &H18438B08: z_Sc(48) = &H4C25B: z_Sc(49) = &H245C8B53: z_Sc(50) = &H24448B08: z_Sc(51) = &H1843890C: z_Sc(52) = &HC25BC031: z_Sc(53) = &HF6890008: z_Sc(54) = &H245C8B53: z_Sc(55) = &H52E808: z_Sc(56) = &HC25B0000: z_Sc(57) = &HF6890004: z_Sc(58) = &H245C8B53: z_Sc(59) = &HE8C93108
    z_Sc(60) = &H24&: z_Sc(61) = &H4C25B: z_Sc(62) = &H245C8B53: z_Sc(63) = &H3C438B08: z_Sc(64) = &H4C25B: z_Sc(65) = &H245C8B53: z_Sc(66) = &H24448B08: z_Sc(67) = &H3C43890C: z_Sc(68) = &HC25BC031: z_Sc(69) = &HF6890008: z_Sc(70) = &H574C985: z_Sc(71) = &HEB3C4B89: z_Sc(72) = &H3C4B8B03: z_Sc(73) = &H511473FF: z_Sc(74) = &H6A3873FF: z_Sc(75) = &H2C53FF00: z_Sc(76) = &HC3384389: z_Sc(77) = &H8538438B: z_Sc(78) = &H500D74C0: z_Sc(79) = &H53FF006A: z_Sc(80) = &H3843C730: z_Sc(82) = &H768DC3: z_Sc(83) = &H830843FF: z_Sc(84) = &H7500107B: z_Sc(85) = &HA6E807: z_Sc(86) = &H30EB0000: z_Sc(87) = &H8524438B: z_Sc(88) = &HFF0774C0: z_Sc(89) = &H1F883D0
    z_Sc(90) = &HA9E82275: z_Sc(91) = &H8D000000: z_Sc(92) = &HFF501C45: z_Sc(93) = &H75FF1875: z_Sc(94) = &HC73FF30: z_Sc(95) = &H8B1053FF: z_Sc(96) = &HC0851845: z_Sc(97) = &H8B500674: z_Sc(98) = &H850FF00: z_Sc(99) = &H31084BFF: z_Sc(100) = &H1C4589C0: z_Sc(101) = &H10C261: z_Sc(102) = &H87B83: z_Sc(103) = &HC2610474: z_Sc(104) = &H75FF0010: z_Sc(105) = &H2475FF2C: z_Sc(106) = &HFF3053FF: z_Sc(107) = &H53FF2073: z_Sc(108) = &H28438B34: z_Sc(109) = &H8B144589: z_Sc(110) = &HFF53147B: z_Sc(111) = &H45C73453: z_Sc(112) = &H800030: z_Sc(113) = &H2C45C700: z_Sc(115) = &H61287D89: z_Sc(116) = &H24048958: z_Sc(117) = &HF689E2FF: z_Sc(118) = &HC43C7: z_Sc(119) = &HC7000000
    z_Sc(120) = &H1043&: z_Sc(121) = &H49E80000: z_Sc(122) = &H8BFFFFFF: z_Sc(123) = &HC0851C43: z_Sc(124) = &H8B500674: z_Sc(125) = &H850FF00: z_Sc(126) = &H3E8&: z_Sc(127) = &HF689C300: z_Sc(128) = &H8314438B: z_Sc(129) = &H6A5010C0: z_Sc(130) = &HC73FF32: z_Sc(131) = &H53FF006A: z_Sc(132) = &HC43892C: z_Sc(133) = &H768DC3: z_Sc(134) = &HC710EC83: z_Sc(135) = &H4002404: z_Sc(136) = &H44C70002: z_Sc(137) = &H424&: z_Sc(138) = &H44C70000: z_Sc(139) = &HC00824: z_Sc(140) = &H44C70000: z_Sc(141) = &HC24&: z_Sc(142) = &H148D4600: z_Sc(143) = &H1845C724: z_Sc(145) = &H8D18458D: z_Sc(146) = &H8B525108: z_Sc(147) = &H8B501C43: z_Sc(148) = &H8310FF00: z_Sc(149) = &HC310C4

    '/// patch thunk & build VTable -- same order as sent to pvCreateITypeInfo
    z_Sc(1) = hData: z_Sc(5) = hData        ' 2 non-VTable functions, pass objptr to them
    z_Sc(THUNK_SIZE + 0) = ((hThunk Xor &H80000000) + &H20) Xor &H80000000 ' IUnknown:QueryInterface(QI)
    z_Sc(THUNK_SIZE + 1) = ((hThunk Xor &H80000000) + &H7C) Xor &H80000000 ' IUnknown:AddRef
    z_Sc(THUNK_SIZE + 2) = ((hThunk Xor &H80000000) + &H8C) Xor &H80000000 ' IUnknown:Release
    z_Sc(THUNK_SIZE + 3) = ((hThunk Xor &H80000000) + &HAC) Xor &H80000000 ' State
    z_Sc(THUNK_SIZE + 4) = ((hThunk Xor &H80000000) + &HB8) Xor &H80000000 ' Tag Get
    z_Sc(THUNK_SIZE + 5) = ((hThunk Xor &H80000000) + &HC4) Xor &H80000000 ' Tag Let
    z_Sc(THUNK_SIZE + 6) = ((hThunk Xor &H80000000) + &HD8) Xor &H80000000 ' Pause
    z_Sc(THUNK_SIZE + 7) = ((hThunk Xor &H80000000) + &HE8) Xor &H80000000 ' Restart
    z_Sc(THUNK_SIZE + 8) = ((hThunk Xor &H80000000) + &HF8) Xor &H80000000 ' Interval Get
    z_Sc(THUNK_SIZE + 9) = ((hThunk Xor &H80000000) + &H104) Xor &H80000000 ' Interval Let
    thunkCopyMemory ByVal hThunk, z_Sc(0), (THUNK_SIZE + VTable_SIZE) * 4
    Erase z_Sc()
    
    ' /// create the VB object's VTable
    ReDim z_Sc(DATA_SIZE - 1)               ' update VTable location
    z_Sc(0) = ((hThunk Xor &H80000000) + THUNK_SIZE * 4) Xor &H80000000
    z_Sc(1) = 1                             ' ref counter
    z_Sc(2) = 0                             ' recursion counter
    z_Sc(3) = ObjPtr(VbHost)               ' passed VB host
    z_Sc(4) = FunctionOrdinal               ' host function pointer
    z_Sc(5) = hThunk                        ' this thunk address
    z_Sc(6) = 0                             ' thunk tag
    z_Sc(7) = 0                             ' Dispatch::IUnknown
    z_Sc(8) = hTLBstructs                   ' TypeLib structs
    z_Sc(9) = lEbMode                       ' optional IDE safety EbMode API or null
    z_Sc(10) = thunkGetProcAddress(hMod("k32"), ByVal "VirtualFree")
    z_Sc(11) = thunkGetProcAddress(hMod("u32"), ByVal "SetTimer")
    z_Sc(12) = thunkGetProcAddress(hMod("u32"), ByVal "KillTimer")
    z_Sc(13) = thunkGetProcAddress(hMod("o32"), ByVal "CoTaskMemFree")
    z_Sc(14) = 0                            ' Timer ID
    z_Sc(15) = InitInterval                 ' Timer interval
    thunkCopyMemory ByVal hData, z_Sc(0), DATA_SIZE * 4
    thunkCopyMemory oIUnk, hData, 4
    Erase z_Sc()
    Set hMod = Nothing
    
    Debug.Print "data section: "; hData
    
    ' create the adhoc IDispatch. Enables VB to recognize this as an object
    thunkCreateStdDispatch hData, hData, ObjPtr(oITypeInfo), VarPtr(oIDispatch)
    Set oITypeInfo = Nothing
    
    Debug.Assert (oIDispatch <> 0)
    If oIDispatch <> 0 Then                 ' if failure, we unwind
        ' copy the IDispatch pointer to the object data & initialize the timer
        thunkCopyMemory ByVal ((hData Xor &H80000000) + 28) Xor &H80000000, oIDispatch, 4
        Set oObj = oIUnk
        If InitInterval <> 0 Then oObj.resume
        Set CreateTasker_Timer = oObj: Set oObj = Nothing
    End If
    Set oIUnk = Nothing ' will self-delete thunk & memory allocations if routine failed
    Exit Function
    
SetAPIreferences:       ' modules used by the thunk
    hMod.Add thunkGetModuleHandle("user32.dll"), "u32"
    hMod.Add thunkGetModuleHandle("kernel32.dll"), "k32"
    hMod.Add thunkGetModuleHandle("ole32.dll"), "o32"
    If AddIDEsafety = True Then
        On Error Resume Next
        Debug.Print 1 / 0
        If Err Then
            Err.Clear: On Error GoTo 0
            hMod.Add thunkGetModuleHandle("vba6.dll"), "vba"
            ' note: if next line errors, change vba6.dll to vba5.dll
            lEbMode = thunkGetProcAddress(hMod("vba"), ByVal "EbMode")
        End If
        On Error GoTo 0
    End If
    Return
    
AbortCleanup:
    If hData <> 0 Then thunkCoTaskMemFree hData
    If hTLBstructs <> 0 Then
        Set oITypeInfo = Nothing
        thunkCoTaskMemFree hTLBstructs
    End If
    
End Function

Public Function CreateTasker_HookCOM(ByVal FunctionOrdinal As Long, _
                                    ByVal VTableOrdinal As Long, _
                                    ByVal pSize As Long, _
                                    ObjectCOM As Variant) As Variant

    ' This routine is called at least 3 times
    '   once to identify the host and hooked interface's VTable method count
    '   once for EACH method to be hooked
    '   once to finish and return the tasker object
    
    ' The parameters are dependent on the FunctionOrdinal value
    ' --------------------------------------------------------------------------
    ' FunctionOrdinal = 0 (identify the number of interface methods and tasker host)
    '   VTableOrdinal   ignored
    '   pSize           nr of VTable methods for primary interface, including IUnknown
    '                   ** max accepted value is 255 methods. min accepted is 4
    '   ObjectCOM       the host to recieve tasker messages, i.e., Me
    ' Returns zero if failure
    ' sample for a stdPicture: CreateTasker_HookCOM(0, 0, 18, Me)
    
    ' FunctionOrdinal = 1 to n: host function ordinal to receive messages for VTable ordinal
    '   VTableOrdinal   one-bound VTable ordinal of the default interface's method to be hooked
    '   pSize           nr of DWords used by that method's parameters, zero if none
    '                   ** max accepted value is 255 DWords
    '   ObjectCOM       dispInterface or primary interface if dispInterface doesn't apply
    ' Returns zero if failure
    ' sample for stdPicture.Render (9th method): CreateTasker_HookCOM(1, 9, 10, Image1.Picture)
    
    ' FunctionOrdinal = -1 or -2: -1 enables IDE-safety, -2 prevents it
    ' ** called when no more methods are to be hooked, finishes the Tasker & returns it
    '   VTableOrdinal   ignored
    '   pSize           ignored
    '   ObjectCOM       ignored, pass as Nothing
    ' Returns NOTHING if failure else returns the tasker object
    ' sample: CreateTasker_HookCOM(-1, 0, 0, Nothing)
    ' --------------------------------------------------------------------------
    
    ' returned object has these methods
    ' Method Name       Sample Call             Purpose/Notes on next line
    ' ----------------------------------------------------------------------------
    ' HookedObjPtr      oTasker.HookedObjPtr = ObjPtr(Me.Picture): Print oTasker.HookedObjPtr
    '                   returns/sets hooked object. setting to zero removes hook
    ' RefCount          Print oTasker.RefCount
    '                   returns reference count reported by hooked object
    ' Tag               oTasker.Tag=1908: Print oTasker.Tag
    '                   User-defined thunk tag (Long values only)
    ' ----------------------------------------------------------------------------
    
    Static nrDefMethods As Long
    Static pMethods() As Long, nrRedirects As Long
    Static oHost As Object, oIDisp As Variant
    
    Dim z_Sc() As Long, z_Bytes() As Byte
    Dim hThunk As Long, hData As Long
    Dim oIDispatch As Long, hRedirects As Long
    Dim oIUnk As stdole.IUnknown
    Dim oITypeInfo As stdole.IUnknown
    Dim hMod As New Collection
    Dim hTLBstructs As Long, lEbMode As Long
    Dim p As Long, n As Long, j As Long
    Dim ptrPrimary As Long, ptrDefault As Long
    Dim bWantIUnkRelease As Boolean
    
    Const THUNK_SIZE As Long = 259          ' size of the thunk array (in longs)
    Const VTable_SIZE As Long = 8           ' nr of VTable entries (in longs)
    Const DATA_SIZE As Long = 19            ' Object data size (in longs)
    Const STUB_SIZE As Long = 16            ' nr of bytes for a method stub
    Const MAX_LIMIT As Long = 255           ' reasonable limit to prevent bad parameter values
    
    Select Case FunctionOrdinal
    Case 0          ' initialize
        CreateTasker_HookCOM = 0
        nrRedirects = 0: Erase pMethods(): oIDisp = Empty: Set oHost = Nothing
        If pSize < 4 Or pSize > MAX_LIMIT Then Exit Function
        If IsObject(ObjectCOM) = False Then Exit Function
        If ObjectCOM Is Nothing Then Exit Function
        
        Set oHost = ObjectCOM
        ReDim pMethods(0 To 14)             ' pre-size for up to 5 redirected methods
        nrDefMethods = pSize
        CreateTasker_HookCOM = pSize
        Exit Function
        
    Case Is > 0     ' set hooked methods
        CreateTasker_HookCOM = 0
        If nrDefMethods = 0 Then GoTo releaseStatics
        If IsEmpty(ObjectCOM) = True Then GoTo releaseStatics
        If IsObject(ObjectCOM) = False Then
            ' must be an object or stdole.IUnknown (vbDataObject)
            If VarType(ObjectCOM) <> vbDataObject Then GoTo releaseStatics
        End If
        If ObjPtr(ObjectCOM) = 0 Then Exit Function
        If FunctionOrdinal < 1 Or pSize < 0 Or pSize > MAX_LIMIT Then GoTo releaseStatics
        If VTableOrdinal < 1 Or VTableOrdinal > &H7FFF Then GoTo releaseStatics
        If IsEmpty(oIDisp) = True Then Set oIDisp = ObjectCOM
        If Not (oIDisp Is ObjectCOM) Then GoTo releaseStatics
        
        n = nrRedirects * 3: nrRedirects = nrRedirects + 1
        If n > UBound(pMethods) Then ReDim Preserve pMethods(0 To n + 5)
        pMethods(n + 2) = pvGetAddressOf(oHost, FunctionOrdinal)
        Debug.Assert pMethods(n + 2) <> 0
        If pMethods(n + 2) = 0 Then GoTo releaseStatics
        pMethods(n + 0) = VTableOrdinal - 1 ' use zero-bound for thunk, not one-bound
        pMethods(n + 1) = pSize * 4 + 4     ' byte count vs long count: include COM pThis param
        CreateTasker_HookCOM = nrRedirects
        Exit Function
        
    Case -1, -2     ' finalize & release references
        Set CreateTasker_HookCOM = Nothing
        If nrRedirects = 0 Then GoTo releaseStatics
        If FunctionOrdinal = -1 Then lEbMode = FunctionOrdinal
        
        For p = 0 To nrRedirects * 3 - 1 Step 3 ' test for hooking IUnknown::Release
            If pMethods(p) = 2 Then bWantIUnkRelease = True: Exit For
        Next
    Case Else: GoTo releaseStatics
    End Select
    
    ' call local routine to create a temporary typelib for adhoc IDispatch interface
    ' 1=function, 2=propGet, 4=propLet
    Set oITypeInfo = pvCreateITypeInfo(hTLBstructs, 2, "Tag", 4, "Tag", "Value", _
                                                    2, "HookedObjPtr", 4, "HookedObjPtr", "pValue", _
                                                    2, "RefCount")
    Debug.Assert (ObjPtr(oITypeInfo) <> 0)
    If oITypeInfo Is Nothing Then GoTo AbortCleanup
    
    hData = thunkCoTaskMemAlloc(DATA_SIZE * 4)  ' memory for our ObjPtr
    Debug.Assert hData <> 0
    If hData = 0 Then GoTo AbortCleanup
    
    hRedirects = thunkCoTaskMemAlloc((nrDefMethods + 8) * 4) ' memory for 2 VTables + counter
    Debug.Assert hRedirects <> 0
    If hRedirects = 0 Then GoTo AbortCleanup
    
    ptrDefault = ObjPtr(oIDisp)                 ' dispInterface objPtr
    Set oIUnk = oIDisp                          ' get primary interface
    ptrPrimary = ObjPtr(oIUnk)                  ' primary interface objPtr
    If ptrPrimary = ptrDefault Then ptrDefault = 0 ' if same, set dispInterface ptr to zero
    n = Abs(bWantIUnkRelease) Xor 1             ' increment count for IUnknown::Release as needed
    If ptrDefault <> 0 Then n = n + 1           ' add stub for dispInterface Release
    n = nrRedirects + n                         ' total stub count
    
    hThunk = thunkVirtualAlloc(0, (THUNK_SIZE + VTable_SIZE) * 4 + n * STUB_SIZE)
    Debug.Assert hThunk <> 0
    If hThunk = 0 Then GoTo AbortCleanup    ' thunk creation failed
    
    Debug.Print "thunk: "; hThunk
                                        
    GoSub SetAPIreferences  ' Gosub? Just wanted to declutter top half of routine
    
    n = THUNK_SIZE + (n * STUB_SIZE) \ 4        ' size of thunk in Longs
    ReDim z_Sc(0 To n + VTable_SIZE - 1)        ' include oTasker VTable
    z_Sc(0) = &HBB60C031: z_Sc(1) = &H12345678: z_Sc(2) = &H95E9E589: z_Sc(3) = &H90000002: z_Sc(4) = &HB890C089: z_Sc(5) = &H12345678: z_Sc(6) = &H2CBE9: z_Sc(7) = &H768D00: z_Sc(8) = &H8BE58960: z_Sc(9) = &H458B245D: z_Sc(10) = &H4488B28: z_Sc(11) = &H8108480B: z_Sc(12) = &HC0F9&: z_Sc(13) = &H81377500: z_Sc(14) = &HC78&: z_Sc(15) = &H2E754600: z_Sc(16) = &HC985088B: z_Sc(17) = &H558B0C75: z_Sc(18) = &HFF1A892C: z_Sc(19) = &HC0310443: z_Sc(20) = &H438B21EB: z_Sc(21) = &H74C0851C: z_Sc(22) = &HF98115: z_Sc(23) = &H75000204: z_Sc(24) = &H2C75FF0D: z_Sc(25) = &H502875FF: z_Sc(26) = &H10FF008B: z_Sc(27) = &H2B805EB: z_Sc(28) = &H89800040: z_Sc(29) = &HC2611C45
    z_Sc(30) = &HF689000C: z_Sc(31) = &H424548B: z_Sc(32) = &H4004428B: z_Sc(33) = &HC2044289: z_Sc(34) = &HF6890004: z_Sc(35) = &H424548B: z_Sc(36) = &H4804428B: z_Sc(37) = &H85044289: z_Sc(38) = &H600C75C0: z_Sc(39) = &H5D8BE589: z_Sc(40) = &H156E824: z_Sc(41) = &HC2610000: z_Sc(42) = &HF6890004: z_Sc(43) = &H245C8B53: z_Sc(44) = &H18438B08: z_Sc(45) = &H4C25B: z_Sc(46) = &H245C8B53: z_Sc(47) = &H24448B08: z_Sc(48) = &H1843890C: z_Sc(49) = &HC25BC031: z_Sc(50) = &HF6890008: z_Sc(51) = &H245C8B53: z_Sc(52) = &H40438B08: z_Sc(53) = &H4C25B: z_Sc(54) = &H8960C031: z_Sc(55) = &H245D8BE5: z_Sc(56) = &H87B83: z_Sc(57) = &H81850F: z_Sc(58) = &H41E80000: z_Sc(59) = &HE8000001
    z_Sc(60) = &HA0&: z_Sc(61) = &H8528458B: z_Sc(62) = &H317074C0: z_Sc(63) = &H68D2&: z_Sc(64) = &HC0684600: z_Sc(65) = &H52000000: z_Sc(66) = &H24148D52: z_Sc(67) = &H8D384389: z_Sc(68) = &HFF89407B: z_Sc(69) = &H52573F8D: z_Sc(70) = &H89008B50: z_Sc(71) = &H10FF3C43: z_Sc(72) = &H8510C483: z_Sc(73) = &H8B3C75C0: z_Sc(74) = &H55894053: z_Sc(75) = &H38533B1C: z_Sc(76) = &H43C70E75: z_Sc(77) = &H3C&: z_Sc(78) = &H3843C700: z_Sc(80) = &H4389028B: z_Sc(81) = &H116E844: z_Sc(82) = &H538B0000: z_Sc(83) = &H487B8B40: z_Sc(84) = &H8920C783: z_Sc(85) = &H38538B3A: z_Sc(86) = &HF74D285: z_Sc(87) = &H891CEF83: z_Sc(88) = &H3108EB3A: z_Sc(89) = &H3C4389C0
    z_Sc(90) = &H61384389: z_Sc(91) = &H900008C2: z_Sc(92) = &H8B60C031: z_Sc(93) = &HE589245D: z_Sc(94) = &H8540538B: z_Sc(95) = &H8B1074D2: z_Sc(96) = &HFF52447B: z_Sc(97) = &H73FF0457: z_Sc(98) = &H857FF40: z_Sc(99) = &H611C4589: z_Sc(100) = &H900004C2: z_Sc(101) = &HE3384B8B: z_Sc(102) = &H3C438B0D: z_Sc(103) = &HC0310189: z_Sc(104) = &H89384389: z_Sc(105) = &H4B8B3C43: z_Sc(106) = &H8B0DE340: z_Sc(107) = &H1894443: z_Sc(108) = &H4389C031: z_Sc(109) = &H44438940: z_Sc(110) = &H768DC3: z_Sc(111) = &HC710EC83: z_Sc(112) = &H4002404: z_Sc(113) = &H44C70002: z_Sc(114) = &H424&: z_Sc(115) = &H44C70000: z_Sc(116) = &HC00824: z_Sc(117) = &H44C70000: z_Sc(118) = &HC24&: z_Sc(119) = &H148D4600
    z_Sc(120) = &H1845C724: z_Sc(122) = &H8D18458D: z_Sc(123) = &H89525108: z_Sc(124) = &H28B52F2: z_Sc(125) = &HC48310FF: z_Sc(126) = &HF689C310: z_Sc(127) = &HC43C7: z_Sc(128) = &HE8000000: z_Sc(129) = &HFFFFFF8C: z_Sc(130) = &H851C438B: z_Sc(131) = &H500674C0: z_Sc(132) = &H50FF008B: z_Sc(133) = &H2E808: z_Sc(134) = &H90C30000: z_Sc(135) = &H5014438B: z_Sc(136) = &H73FF326A: z_Sc(137) = &HFF006A10: z_Sc(138) = &H43892C53: z_Sc(139) = &HF689C310: z_Sc(140) = &H8B487B8B: z_Sc(141) = &HC0850C47: z_Sc(142) = &H47C72574: z_Sc(143) = &HC&: z_Sc(144) = &H44738B00: z_Sc(145) = &HC7830F8B: z_Sc(146) = &H4EE831C: z_Sc(147) = &H8B8E048B: z_Sc(148) = &HC2398F14: z_Sc(149) = &H4C70775
    z_Sc(150) = &H8F&: z_Sc(151) = &HC3EDE200: z_Sc(152) = &H8B3C4B8B: z_Sc(153) = &H15E3487B: z_Sc(154) = &H7B9CE89: z_Sc(155) = &H83000000: z_Sc(156) = &HA5F304C7: z_Sc(157) = &H8B487B8B: z_Sc(158) = &H47892847: z_Sc(159) = &H44738B0C: z_Sc(160) = &HC7830F8B: z_Sc(161) = &H4EE831C: z_Sc(162) = &H858F048B: z_Sc(163) = &H8B0675C0: z_Sc(164) = &H4898E04: z_Sc(165) = &H83F1E28F: z_Sc(166) = &H75003C7B: z_Sc(167) = &H1CEF8306: z_Sc(168) = &HC30C5F89: z_Sc(169) = &H87B83: z_Sc(170) = &HC2610474: z_Sc(171) = &H75FF0010: z_Sc(172) = &H2475FF2C: z_Sc(173) = &HFF3053FF: z_Sc(174) = &H53FF4873: z_Sc(175) = &H2073FF34: z_Sc(176) = &H8B3453FF: z_Sc(177) = &H45892843: z_Sc(178) = &H147B8B14: z_Sc(179) = &H3453FF53
    z_Sc(180) = &H3045C7: z_Sc(181) = &HC7000080: z_Sc(182) = &H2C45&: z_Sc(183) = &H7D890000: z_Sc(184) = &H89586128: z_Sc(185) = &HE2FF2404: z_Sc(186) = &H89C38960: z_Sc(187) = &HFFC031E5: z_Sc(188) = &HEC830843: z_Sc(189) = &H184D8B10: z_Sc(190) = &HE9C1C889: z_Sc(191) = &HF84D8910: z_Sc(192) = &HFFFF25: z_Sc(193) = &HF0458900: z_Sc(194) = &H8B44538B: z_Sc(195) = &HF983487B: z_Sc(196) = &H8B0E7502: z_Sc(197) = &H433B2445: z_Sc(198) = &H8B067440: z_Sc(199) = &HEF833C53: z_Sc(200) = &H8A048B24: z_Sc(201) = &H8DFC4589: z_Sc(202) = &H57208F7C: z_Sc(203) = &H78937FF: z_Sc(204) = &H7F02F983: z_Sc(205) = &HF04D8B2A: z_Sc(206) = &H3C8DCC29: z_Sc(207) = &H24758D24: z_Sc(208) = &HF302E9C1: z_Sc(209) = &HFC55FFA5
    z_Sc(210) = &H831C4589: z_Sc(211) = &H7502F87D: z_Sc(212) = &H75C0850E: z_Sc(213) = &HFED6E80A: z_Sc(214) = &H35E8FFFF: z_Sc(215) = &H83FFFFFE: z_Sc(216) = &HF000C7B: z_Sc(217) = &H8384&: z_Sc(218) = &H147D8300: z_Sc(219) = &H8B7D7400: z_Sc(220) = &HC0852443: z_Sc(221) = &HD0FF0774: z_Sc(222) = &H7501F883: z_Sc(223) = &H14458B6F: z_Sc(224) = &H8BF44589: z_Sc(225) = &H31E81C73: z_Sc(226) = &H8DFFFFFE: z_Sc(227) = &H8B501C45: z_Sc(228) = &HE983F04D: z_Sc(229) = &H290DE304: z_Sc(230) = &H243C8DCC: z_Sc(231) = &HC128758D: z_Sc(232) = &HA5F302E9: z_Sc(233) = &H73FFF631: z_Sc(234) = &H14758940: z_Sc(235) = &H5014458D: z_Sc(236) = &HFF1875FF: z_Sc(237) = &H55FF0C73: z_Sc(238) = &H558B50F4: z_Sc(239) = &H52028B18
    z_Sc(240) = &H580850FF: z_Sc(241) = &H3F87D83: z_Sc(242) = &H7539087C: z_Sc(243) = &H891D7414: z_Sc(244) = &H4BFF1C45: z_Sc(245) = &HF04D8B08: z_Sc(246) = &H8BEC7D8B: z_Sc(247) = &H789E845: z_Sc(248) = &H8918C483: z_Sc(249) = &H5A61184D: z_Sc(250) = &HE2FFCC01: z_Sc(251) = &H3F87D83: z_Sc(252) = &H4BFFE07C: z_Sc(253) = &HFC558B08: z_Sc(254) = &H8BEC7D8B: z_Sc(255) = &H789E845: z_Sc(256) = &H83145589: z_Sc(257) = &HFF6118C4: z_Sc(258) = &HE2

    '/// patch thunk & build VTable -- same order as sent to pvCreateITypeInfo
    z_Sc(1) = hData: z_Sc(5) = hData        ' patch 2 non-VTable methods
    z_Sc(n + 0) = ((hThunk Xor &H80000000) + &H20) Xor &H80000000 ' IUnknown:QueryInterface(QI)
    z_Sc(n + 1) = ((hThunk Xor &H80000000) + &H7C) Xor &H80000000 ' IUnknown:AddRef
    z_Sc(n + 2) = ((hThunk Xor &H80000000) + &H8C) Xor &H80000000 ' IUnknown:Release
    z_Sc(n + 3) = ((hThunk Xor &H80000000) + &HAC) Xor &H80000000 ' Tag Get
    z_Sc(n + 4) = ((hThunk Xor &H80000000) + &HB8) Xor &H80000000 ' Tag Let
    z_Sc(n + 5) = ((hThunk Xor &H80000000) + &HCC) Xor &H80000000 ' HookedObjPtr Get
    z_Sc(n + 6) = ((hThunk Xor &H80000000) + &HD8) Xor &H80000000 ' HookedObjPtr Let
    z_Sc(n + 7) = ((hThunk Xor &H80000000) + &H170) Xor &H80000000 ' RefCount
    thunkCopyMemory ByVal hThunk, z_Sc(0), (n + VTable_SIZE) * 4
    Erase z_Sc()
    
    ' /// create the VB object's VTable
    ReDim z_Sc(DATA_SIZE - 1)               ' update VTable location
    z_Sc(0) = ((hThunk Xor &H80000000) + n * 4) Xor &H80000000
    z_Sc(1) = 1                             ' ref counter
    z_Sc(2) = 0                             ' recursion counter
    z_Sc(3) = ObjPtr(oHost)                 ' passed VB host
    z_Sc(4) = 0                             '
    z_Sc(5) = hThunk                        ' this thunk address
    z_Sc(6) = 0                             ' thunk tag
    z_Sc(7) = 0                             ' IDispatch::IUnknown
    z_Sc(8) = hTLBstructs                   ' TypeLib structs
    z_Sc(9) = lEbMode                       ' optional IDE safety EbMode API or null
    z_Sc(10) = thunkGetProcAddress(hMod("k32"), ByVal "VirtualFree")
    z_Sc(11) = thunkGetProcAddress(hMod("u32"), ByVal "SetTimer")
    z_Sc(12) = thunkGetProcAddress(hMod("u32"), ByVal "KillTimer")
    z_Sc(13) = thunkGetProcAddress(hMod("o32"), ByVal "CoTaskMemFree")
    z_Sc(14) = ptrDefault                   ' dispInterface ObjPtr
    If ptrDefault = 0 Then  ' flag used by the thunk when dispInterface doesn't apply
        thunkCopyMemory ByVal ((hRedirects Xor &H80000000) + 12) Xor &H80000000, -1&, 4
    Else
        thunkCopyMemory z_Sc(15), ByVal ptrDefault, 4 ' original VTable address
        thunkCopyMemory ByVal ((hRedirects Xor &H80000000) + 4) Xor &H80000000, _
                        ByVal z_Sc(15), 28 ' original function pointers
    End If
    z_Sc(16) = ptrPrimary                   ' default ObjPtr
    thunkCopyMemory z_Sc(17), ByVal ptrPrimary, 4 ' original VTable address
    thunkCopyMemory ByVal ((hRedirects Xor &H80000000) + 32) Xor &H80000000, _
                    ByVal z_Sc(17), nrDefMethods * 4 ' original function pointers
    z_Sc(18) = hRedirects                  ' VTable redirection for both interfaces
    thunkCopyMemory ByVal hRedirects, nrDefMethods, 4 ' primary methods count
    
    thunkCopyMemory ByVal hData, z_Sc(0), DATA_SIZE * 4
    Erase z_Sc()
    Set hMod = Nothing
    Set oIUnk = Nothing
    
    ReDim z_Bytes(0 To STUB_SIZE - 1)
    ' z_Bytes(1-4) the host function address
    ' z_Bytes(6-9) packed method ordinal/method params size
    ' z_Bytes(11-14) negative relative jump from stub last instruction to default procedure
    ' ----------------- the stub is a simple 15 byte block -----------------
    ' mov   edx, 12345678h      &HBA + z_Bytes(1-4)
    ' mov   ecx, 12345678h      &HB9 + z_Bytes(6-9)
    ' jmp   _JUMP_ANCHOR        &HE9 + z_Bytes(11-14), _JUMP_ANCHOR = thunk start + 19
    ' --------------------------------------------------------------------------------
    z_Bytes(0) = &HBA: z_Bytes(5) = &HB9: z_Bytes(10) = &HE9: z_Bytes(15) = &H90
    
    ' location in thunk for 1st stub (each are 16 bytes, dword aligned)
    p = ((hThunk Xor &H80000000) + THUNK_SIZE * 4) Xor &H80000000
    j = 19 - (THUNK_SIZE * 4 + 15)  ' calculate relative jump used by 1st stub
    '^^ 19 = location of default procedure, 15 = length of the stub w/o padding
    
    ' for each of the hooked methods passed to this routine, redirect them...
    For n = 0 To nrRedirects * 3 - 1 Step 3
        ' n+0 = Ordinal, n+1 = pThis+params byte size, n+2 = host function address
        thunkCopyMemory z_Bytes(1), pMethods(n + 2), 4   ' host callback address
        ' pack Ordinal (hiWord) with params size (loWord)
        thunkCopyMemory z_Bytes(6), (pMethods(n) * &H10000 Or pMethods(n + 1)), 4
        thunkCopyMemory z_Bytes(11), j, 4               ' relative jump
        thunkCopyMemory ByVal p, z_Bytes(0), STUB_SIZE
        ' set the redirection for this method to the stub we just created
        thunkCopyMemory ByVal ((hRedirects Xor &H80000000) + pMethods(n) * 4 + 32) Xor &H80000000, p, 4
        p = ((p Xor &H80000000) + STUB_SIZE) Xor &H80000000 ' next stub location
        j = j - STUB_SIZE                               ' next relative jump
    Next
    If bWantIUnkRelease = False Then        ' create stub for primary's IUnknown::Release
        thunkCopyMemory z_Bytes(1), 0&, 4
        thunkCopyMemory z_Bytes(6), &H20004, 4
        thunkCopyMemory z_Bytes(11), j, 4
        thunkCopyMemory ByVal p, z_Bytes(0), STUB_SIZE
        ' set redirection for Release: 8 bytes after dispInterface VTable size + counter = 40
        thunkCopyMemory ByVal ((hRedirects Xor &H80000000) + 40) Xor &H80000000, p, 4
    End If
    ' copy the redirect address over existing VTable pointer
    thunkCopyMemory ByVal ptrPrimary, ((hRedirects Xor &H80000000) + 32) Xor &H80000000, 4
    
    If ptrDefault <> 0 Then             ' hook dispInterface's IUnknown::Release
        ' use the same stub for primary's IUnknown::Release, 40 bytes into hRedirects (36+counter)
        thunkCopyMemory p, ByVal ((hRedirects Xor &H80000000) + 40) Xor &H80000000, 4
        thunkCopyMemory ByVal ((hRedirects Xor &H80000000) + 12) Xor &H80000000, p, 4
        ' copy the redirect address over existing VTable pointer
        thunkCopyMemory ByVal ptrDefault, ((hRedirects Xor &H80000000) + 4) Xor &H80000000, 4
    End If
    Erase z_Bytes()
    
    thunkCopyMemory oIUnk, hData, 4
    Debug.Print "data section: "; hData
    
    ' create the adhoc IDispatch. Enables VB to recognize this as an object
    thunkCreateStdDispatch hData, hData, ObjPtr(oITypeInfo), VarPtr(oIDispatch)
    Set oITypeInfo = Nothing
    
    Debug.Assert (oIDispatch <> 0)
    If oIDispatch <> 0 Then                 ' if failure, we unwind
        ' copy the IDispatch pointer to the object data
        thunkCopyMemory ByVal ((hData Xor &H80000000) + 28) Xor &H80000000, oIDispatch, 4
        Set oHost = oIUnk
        Set CreateTasker_HookCOM = oHost
    End If
    Set oIUnk = Nothing ' will self-delete thunk & memory allocations if routine failed
    
releaseStatics:
    nrDefMethods = 0: nrRedirects = 0
    Erase pMethods()
    Set oHost = Nothing: oIDisp = Empty
    Exit Function
    
SetAPIreferences:       ' modules used by the thunk
    hMod.Add thunkGetModuleHandle("user32.dll"), "u32"
    hMod.Add thunkGetModuleHandle("kernel32.dll"), "k32"
    hMod.Add thunkGetModuleHandle("ole32.dll"), "o32"
    If lEbMode <> 0 Then
        lEbMode = 0
        On Error Resume Next
        Debug.Print 1 / 0
        If Err Then
            Err.Clear: On Error GoTo 0
            hMod.Add thunkGetModuleHandle("vba6.dll"), "vba"
            ' note: if next line errors, change vba6.dll to vba5.dll
            lEbMode = thunkGetProcAddress(hMod("vba"), ByVal "EbMode")
        End If
        On Error GoTo 0
    End If
    Return
    
AbortCleanup:
    If hData <> 0 Then thunkCoTaskMemFree hData
    If hRedirects <> 0 Then thunkCoTaskMemFree hRedirects
    If hTLBstructs <> 0 Then
        Set oITypeInfo = Nothing
        thunkCoTaskMemFree hTLBstructs
    End If

End Function


