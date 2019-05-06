Attribute VB_Name = "modCheckingOptions"
' // modCheckingOptions.bas - working with checking options
' // © Krivous Anatoly Anatolevich (The trick), 2016

Option Explicit

' // State of checking
Private Enum eModuleState
    MS_OVFLV_CHECK = 1  ' // Integer overflow
    MS_ARBND_CHECK = 2  ' // Array bounds
    MS_FLCHK_CHECK = 4  ' // Float
End Enum

Private Const SIZEOF_JMP                As Long = 5

' // Number of opcodes that causes checkings
Private Const OVFLV_OPCODES             As Long = 22
Private Const ARBND_OPCODES             As Long = 14
Private Const FLCHK_OPCODES             As Long = 9

Private pOriginalOpcodes(255, 5)            As Long         ' // Original opcodes handlers addresses
Private pNewOpcodes(255, 5)                 As Long         ' // New opcodes handlers addresses
Private pOpcodesGroupTable(5)               As Long         ' // Pointers to opcodes tables handlers
Private pAddrOfNextOpcodeProc               As Long         ' // Address of procedure that performs the next opcode
Private pAddrOfErrRaiseProc                 As Long         ' // Address of proc that raises error
Private lIntOvflOpcodes(OVFLV_OPCODES - 1)  As Long         ' // Integer overflow opcodes
Private lArrBndsOpcodes(ARBND_OPCODES - 1)  As Long         ' // Array bounds opcodes
Private lFltChkOpcodes(FLCHK_OPCODES - 1)   As Long         ' // Float opcodes
Private lBugOpcodes(0)                      As Long         ' // Not Not opcode handler
Private pNewOpcodesHandlers                 As Long         ' // Pointer to page that contains new opcodes handlers
Private eState                              As eModuleState ' // Current state
Private bIsInitialize                       As Boolean      ' // Determine if module has been initialized

' // Set/unset overflow checking
Public Property Let IntegerOverflowCheck( _
                    ByVal bValue As Boolean)
    
    If Not bIsInitialize Then

        Err.Raise 5
        Exit Property

    End If
    
    If bValue And Not CBool(eState And MS_OVFLV_CHECK) Then
    
        If Not UnpatchOpcodes(lIntOvflOpcodes) Then
        
            Err.Raise 5
            Exit Property
            
        End If
        
    ElseIf Not bValue And CBool(eState And MS_OVFLV_CHECK) Then
    
        If Not PatchOpcodes(lIntOvflOpcodes) Then
        
            Err.Raise 5
            Exit Property
            
        End If
        
    End If
    
    eState = eState And (Not MS_OVFLV_CHECK) Or (MS_OVFLV_CHECK And bValue)
    
End Property
Public Property Get IntegerOverflowCheck() As Boolean
    IntegerOverflowCheck = eState And MS_OVFLV_CHECK
End Property

' // Set/unset array bounds checking
Public Property Let ArrayBoundsCheck( _
                    ByVal bValue As Boolean)
    
    If Not bIsInitialize Then

        Err.Raise 5
        Exit Property

    End If
    
    If bValue And Not CBool(eState And MS_ARBND_CHECK) Then
    
        If Not UnpatchOpcodes(lArrBndsOpcodes) Then
        
            Err.Raise 5
            Exit Property
            
        End If
        
    ElseIf Not bValue And CBool(eState And MS_ARBND_CHECK) Then
    
        If Not PatchOpcodes(lArrBndsOpcodes) Then
        
            Err.Raise 5
            Exit Property
            
        End If
        
    End If
    
    eState = eState And (Not MS_ARBND_CHECK) Or (MS_ARBND_CHECK And bValue)
    
End Property
Public Property Get ArrayBoundsCheck() As Boolean
    ArrayBoundsCheck = eState And MS_ARBND_CHECK
End Property

' // Set/unset float operations checking
Public Property Let FloatingPointCheck( _
                    ByVal bValue As Boolean)
    
    If Not bIsInitialize Then

        Err.Raise 5
        Exit Property

    End If
    
    If bValue And Not CBool(eState And MS_FLCHK_CHECK) Then
    
        If Not UnpatchOpcodes(lFltChkOpcodes) Then
        
            Err.Raise 5
            Exit Property
            
        End If
        
    ElseIf Not bValue And CBool(eState And MS_FLCHK_CHECK) Then
    
        If Not PatchOpcodes(lFltChkOpcodes) Then
        
            Err.Raise 5
            Exit Property
            
        End If
        
    End If
    
    eState = eState And (Not MS_FLCHK_CHECK) Or (MS_FLCHK_CHECK And bValue)
    
End Property
Public Property Get FloatingPointCheck() As Boolean
    FloatingPointCheck = eState And MS_FLCHK_CHECK
End Property

' // Uninitialize module
Public Sub UnInitialize()
    
    If Not bIsInitialize Then Exit Sub
    
    IntegerOverflowCheck = True
    ArrayBoundsCheck = True
    FloatingPointCheck = True
    
    UnpatchOpcodes lBugOpcodes()
    
    VirtualFree ByVal pNewOpcodesHandlers, 0, MEM_RELEASE
    
    bIsInitialize = False
    
End Sub

' // Initialize module
Public Function Initialize() As Boolean
    Dim lIndex  As Long
    
    If bIsInitialize Then
        
        Initialize = True
        Exit Function
        
    End If

    ' // Initialize Integer overflow opcodes
    lIntOvflOpcodes(0) = &HFC0D&     ' // Byte <- Integer
    lIntOvflOpcodes(1) = &HFC0E&     ' // Byte <- Long
    lIntOvflOpcodes(2) = &HFC0F&     ' // Byte <- Single
    lIntOvflOpcodes(3) = &HFC10&     ' // Byte <- Double
    lIntOvflOpcodes(4) = &HFC11&     ' // Byte <- Currency
    lIntOvflOpcodes(5) = &HE4&       ' // Integer <- Long
    lIntOvflOpcodes(6) = &HE5&       ' // Integer <- Single, Integer <-Double
    lIntOvflOpcodes(7) = &HE6&       ' // Integer <- Currency
    lIntOvflOpcodes(8) = &HE8&       ' // Long <- Single, Long <-Double
    lIntOvflOpcodes(9) = &HE9&       ' // Long <- Currency
        
    lIntOvflOpcodes(10) = &HFB8E&    ' // Byte + Byte
    lIntOvflOpcodes(11) = &HA9&      ' // Integer + Integer
    lIntOvflOpcodes(12) = &HAA&      ' // Long + Long
    lIntOvflOpcodes(13) = &HAC&      ' // Currency + Currency
    
    lIntOvflOpcodes(14) = &HFBAE&    ' // Byte * Byte
    lIntOvflOpcodes(15) = &HB1&      ' // Integer * Integer
    lIntOvflOpcodes(16) = &HB2&      ' // Long * Long
    lIntOvflOpcodes(17) = &HB4&      ' // Currency * Currency
    
    lIntOvflOpcodes(18) = &HFB96&    ' // Byte - Byte
    lIntOvflOpcodes(19) = &HAD&      ' // Integer - Integer
    lIntOvflOpcodes(20) = &HAE&      ' // Long - Long
    lIntOvflOpcodes(21) = &HB0&      ' // Currency - Currency
    
    ' // Initilaize Array access opcodes
    lArrBndsOpcodes(0) = &HFC90&     ' // One-dim array index Byte
    lArrBndsOpcodes(1) = &H9D&       ' // One-dim array index Integer
    lArrBndsOpcodes(2) = &H9E&       ' // One-dim array index Long
    lArrBndsOpcodes(3) = &HA0&       ' // One-dim array index Single
    lArrBndsOpcodes(4) = &H40&       ' // One-dim array index based on element
    lArrBndsOpcodes(5) = &H41&       ' // One-dim array index Object
    lArrBndsOpcodes(6) = &HA1&       ' // One-dim array index Double
    lArrBndsOpcodes(7) = &HA1&       ' // One-dim array index Currency
    lArrBndsOpcodes(8) = &HFC96&     ' // One-dim array index Variant
    lArrBndsOpcodes(9) = &H9F&       ' // One-dim array index UDT
    lArrBndsOpcodes(10) = &HA7&      ' // Multi-dim array index
    lArrBndsOpcodes(11) = &HFF06&    ' // Multi-dim array index in UDT
    lArrBndsOpcodes(12) = &HFF07&    ' // Multi-dim array index in UDT by ref
    lArrBndsOpcodes(13) = &HA8&      ' // Multi-dim array index
    
    ' // Float checking opcodes
    lFltChkOpcodes(0) = &H73         ' // Store Single
    lFltChkOpcodes(1) = &H74         ' // Store Double
    lFltChkOpcodes(2) = &H37         ' // Push Single
    lFltChkOpcodes(3) = &H39         ' // Push Double
    lFltChkOpcodes(4) = &HFDC8&      ' // Push Single BYREF
    lFltChkOpcodes(5) = &HFDC9&      ' // Push Double BYREF
    lFltChkOpcodes(6) = &HF1&        ' // Push Currency
    lFltChkOpcodes(7) = &HFD6A&      ' // CVar(Single)
    lFltChkOpcodes(8) = &HFD6B&      ' // CVar(Double)
    
    ' // Remove Not Not Arr() bug
    lBugOpcodes(0) = &HED
    
    If Not FindVirtualMashineTable() Then Exit Function

    For lIndex = 0 To UBound(pOpcodesGroupTable)
    
        If pOpcodesGroupTable(lIndex) = 0 Then Exit Function
        memcpy pOriginalOpcodes(0, lIndex), ByVal pOpcodesGroupTable(lIndex), 255 * 4
        
    Next
    
    eState = MS_ARBND_CHECK Or MS_FLCHK_CHECK Or MS_OVFLV_CHECK
    
    bIsInitialize = LoadNewOpcodes()
    
    ' // Fix bug
    PatchOpcodes lBugOpcodes()
    
    Initialize = bIsInitialize
    
End Function

' // Patch opcodes handlers
Private Function PatchOpcodes( _
                 ByRef lOpcodes() As Long) As Boolean
    Dim lIndex  As Long
    Dim pOrigin As Long
    Dim pNew    As Long
    
    On Error GoTo err_skip_opcode
    
    For lIndex = 0 To UBound(lOpcodes)

        pOrigin = GetOpcodeHandlerTableEntryAddress(lOpcodes(lIndex))
        If pOrigin = 0 Then GoTo continue

        If VirtualProtect(ByVal pOrigin, 4, PAGE_EXECUTE_READWRITE, 0) = 0 Then GoTo continue
        
        pNew = GetNewOpcodeHandlerAddress(lOpcodes(lIndex))
        If pNew = 0 Then GoTo continue

        GetMem4 pNew, ByVal pOrigin
        
continue:

    Next
    
    PatchOpcodes = True
    
    Exit Function
    
err_skip_opcode:
    
    pNew = 0
    
    Resume Next
    
End Function

' // Unpatch opcodes handlers
Private Function UnpatchOpcodes( _
                 ByRef lOpcodes() As Long) As Boolean
    Dim pOrigin As Long
    Dim pEntry  As Long
    Dim lIndex  As Long
    
    For lIndex = 0 To UBound(lOpcodes)

        pEntry = GetOpcodeHandlerTableEntryAddress(lOpcodes(lIndex))
        If pEntry = 0 Then GoTo continue
        
        If VirtualProtect(ByVal pEntry, 4, PAGE_EXECUTE_READWRITE, 0) = 0 Then GoTo continue
        
        pOrigin = GetOriginalOpcodeHandlerAddress(lOpcodes(lIndex))
        If pOrigin = 0 Then GoTo continue

        GetMem4 pOrigin, ByVal pEntry
        
continue:
        
    Next
    
    UnpatchOpcodes = True
    
End Function

' // Get table entry of address of opcode handler for specified opcode
Private Function GetOpcodeHandlerTableEntryAddress( _
                ByVal lOpcode As Long) As Long
    Dim lIndex  As Long
    
    If lOpcode < &HFB Then
        
        GetOpcodeHandlerTableEntryAddress = pOpcodesGroupTable(0) + (lOpcode And &HFF) * 4
        
    Else
    
        lIndex = (lOpcode And &HFF00&) \ &H100 - &HFB + 1
        If lIndex > UBound(pOpcodesGroupTable) Then Exit Function
        
        GetOpcodeHandlerTableEntryAddress = pOpcodesGroupTable(lIndex) + (lOpcode And &HFF) * 4
        
    End If
    
End Function

' // Get original opcode handler address
Private Function GetOriginalOpcodeHandlerAddress( _
                 ByVal lOpcode As Long) As Long
    Dim lIndex  As Long
    
    If lOpcode < &HFB Then
        lIndex = 0
    Else
        lIndex = (lOpcode And &HFF00&) \ &H100 - &HFB + 1
        If lIndex > UBound(pOriginalOpcodes, 2) Then Exit Function
    End If
                    
    GetOriginalOpcodeHandlerAddress = pOriginalOpcodes((lOpcode And &HFF), lIndex)
    
End Function

' // Get new opcode handler address
Private Function GetNewOpcodeHandlerAddress( _
                 ByVal lOpcode As Long) As Long
    Dim lIndex  As Long
    
    If lOpcode < &HFB Then
        lIndex = 0
    Else
        lIndex = (lOpcode And &HFF00&) \ &H100 - &HFB + 1
        If lIndex > UBound(pNewOpcodes, 2) Then Exit Function
    End If
                    
    GetNewOpcodeHandlerAddress = pNewOpcodes((lOpcode And &HFF), lIndex)
    
End Function

' // Find opcodes tables, procedures
Private Function FindVirtualMashineTable() As Boolean
    Dim hVBA6                   As Long
    Dim lRet                    As Long
    Dim pEbRaiseExceptionCode   As Long
    
    On Error GoTo error_handler
    
    hVBA6 = GetModuleHandle("vba6")
    If hVBA6 = 0 Then Exit Function
    
    pEbRaiseExceptionCode = GetProcAddress(hVBA6, "EbRaiseExceptionCode")
    If pEbRaiseExceptionCode = 0 Then Exit Function
    
    ' // Parse PE
    Dim pDatPtr(0)  As Long, lDat(1)    As Long, cOld    As Currency
    Dim lSecCount   As Long, lSecIndex  As Long
    
    cOld = PtGet(pDatPtr(), GetDWORD(ArrPtr(lDat())))
    ' // IMAGE_NT_HEADERS
    pDatPtr(0) = hVBA6 + &H3C
    ' // IMAGE_FILE_HEADER.NumberOfSections
    pDatPtr(0) = hVBA6 + lDat(0) + 6
    lSecCount = lDat(0) And &HFFFF&
    ' // Go to first section header
    pDatPtr(0) = pDatPtr(0) + &HF2
    
    ' // Find 'ENGINE' section
    For lSecIndex = 0 To lSecCount - 1
        
        ' // Case insensitive 'ENGINE' (65 6E 67 69 6E 65)
        If (lDat(0) Or &H20202020) = &H69676E65 And _
           (lDat(1) Or &H2020) = &H656E& Then
           
            ' // Found
            Dim lSecSize    As Long:    Dim lSecEnd     As Long
            Dim bSign()     As Byte:    Dim lOpSize     As Long
            Dim lOpcode     As Long:    Dim lExtTbl     As Long
                    
            pDatPtr(0) = pDatPtr(0) + &H8
            
            lSecSize = lDat(0)
            pDatPtr(0) = lDat(1) + hVBA6
            lSecEnd = pDatPtr(0) + lSecSize
            
            ' // Find signature:
            ' // 33C0               XOR EAX,EAX
            ' // 8A06               MOV AL,BYTE PTR DS:[ESI]
            ' // 46                 INC ESI
            ' // FF2485 xxxxxxxx    JMP DWORD PTR DS:[EAX*4+xxxxxxxx]
            
            ReDim bSign(7)
            
            bSign(0) = &H33:  bSign(1) = &HC0: bSign(2) = &H8A: bSign(3) = &H6
            bSign(4) = &H46:  bSign(5) = &HFF: bSign(6) = &H24: bSign(7) = &H85
            
            pAddrOfNextOpcodeProc = FindSignature(pDatPtr(0), lSecSize, bSign())
            
            If pAddrOfNextOpcodeProc = 0 Then
                Exit For
            End If
            
            ' // Search raise exception handler
            ' //        B8 09000000        MOV EAX,9
            ' // +----- E9 xxxxxxxx        JMP pAddrOfErrRaiseProc
            ' // |      ....
            ' // +----> pAddrOfErrRaiseProc:
            ' //        50                 PUSH EAX
            ' //        E8 xxxxxxxx        CALL EbRaiseExceptionCode
            
            ReDim bSign(5)
            
            bSign(0) = &HB8:    bSign(1) = &H9: bSign(2) = &H0: bSign(3) = &H0
            bSign(4) = &H0:     bSign(5) = &HE9
            
            Do
            
                pDatPtr(0) = FindSignature(pDatPtr(0), lSecEnd - pDatPtr(0), bSign())
                 
                If pDatPtr(0) = 0 Then
                    Exit For
                End If
                
                pDatPtr(0) = pDatPtr(0) + 6
                
                ' // Relative offset to absolute
                pDatPtr(0) = (pDatPtr(0) + 4) + lDat(0)
                
                If (lDat(0) And &HFFFF&) = &HE850& Then
                    
                    ' // Check address
                    pDatPtr(0) = pDatPtr(0) + 2
                    
                    If (pDatPtr(0) + 4) + lDat(0) = pEbRaiseExceptionCode Then
                    
                        ' // Found
                        pAddrOfErrRaiseProc = pDatPtr(0) - 2
                        Exit Do
                        
                    End If
                    
                End If

            Loop
            
            ' // Get one-byte opcodes table pointer
            pDatPtr(0) = pAddrOfNextOpcodeProc + 8
            
            pOpcodesGroupTable(0) = lDat(0)
            
            ' // Find extended opcodes tables
            For lExtTbl = 0 To 4
            
                ' // Move to extended opcode table handler
                pDatPtr(0) = pOpcodesGroupTable(0) + (&HFB + lExtTbl) * 4
                pDatPtr(0) = lDat(0)
                
                Do While pDatPtr(0) < lSecEnd - 7
                    
                    ' // Get X86 opcode
                    lOpSize = GetInstructionSize(pDatPtr(0), lOpcode)
                    
                    ' // FF2485 xxxxxxxx    JMP DWORD PTR DS:[EAX*4+xxxxxxxx]
                    
                    If lOpcode = &HFF And _
                       lOpSize = 7 And _
                       lOpcode = (lDat(0) And &HFF) And _
                       (lDat(0) And &H700) = &H400 Then
                        
                        pDatPtr(0) = pDatPtr(0) + 3
                        pOpcodesGroupTable(lExtTbl + 1) = lDat(0)
                        
                        Exit Do
                            
                    End If
                       
                    pDatPtr(0) = pDatPtr(0) + lOpSize
                    
                Loop
                
            Next
            
            Exit For
            
        End If
        
        pDatPtr(0) = pDatPtr(0) + &H28
        
    Next
    
    PtRelease pDatPtr, cOld
    
    FindVirtualMashineTable = CBool(pOpcodesGroupTable(0)) And _
                              CBool(pOpcodesGroupTable(1)) And _
                              CBool(pOpcodesGroupTable(2)) And _
                              CBool(pOpcodesGroupTable(3)) And _
                              CBool(pOpcodesGroupTable(4)) And _
                              CBool(pOpcodesGroupTable(5)) And _
                              CBool(pAddrOfErrRaiseProc) And _
                              CBool(pAddrOfNextOpcodeProc)
                             
    Exit Function
    
error_handler:
    
    ErrorLog "modCheckingOptions:FindVirtualMashineTable"
    
End Function

' // Find signature in memory
Private Function FindSignature( _
                 ByVal pStartAddress As Long, _
                 ByVal lSize As Long, _
                 ByRef bSign() As Byte) As Long
                 
    Dim pDatPtr(0)      As Long:        Dim bDat()          As Byte
    Dim cOld            As Currency:    Dim pEndAddr        As Long
    Dim lIndex          As Long
    
    pEndAddr = pStartAddress + lSize - (UBound(bSign) + 1)
    
    ReDim bDat(UBound(bSign))
    
    cOld = PtGet(pDatPtr(), GetDWORD(ArrPtr(bDat())))
    
    pDatPtr(0) = pStartAddress
    
    Do While pDatPtr(0) < pEndAddr
    
        For lIndex = 0 To UBound(bSign) + 1
            
            If lIndex = UBound(bSign) + 1 Then
            
                FindSignature = pDatPtr(0)
                Exit Do
                
            End If
            
            If bSign(lIndex) <> bDat(lIndex) Then Exit For
            
        Next
        
        pDatPtr(0) = pDatPtr(0) + 1
        
    Loop
    
    PtRelease pDatPtr(), cOld
    
End Function

' // Load modified opcode table and place it in memory
Private Function LoadNewOpcodes() As Boolean
    Dim lOpcode     As Long:    Dim bData()     As Byte
    Dim lCount      As Long:    Dim pCurAddr    As Long
    Dim lJmpValue   As Long:    Dim lRelCount   As Long
    Dim lRelValue   As Long:    Dim lIndex      As Long
    Dim lSubIndex   As Long
    
    On Error GoTo error_handler
    
    pNewOpcodesHandlers = VirtualAlloc(ByVal 0&, 4096, MEM_RESERVE Or MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    If pNewOpcodesHandlers = 0 Then Exit Function

    pCurAddr = pNewOpcodesHandlers
    
    ' // Load raw data
    bData() = LoadResData("HANDLERS", "CUSTOM")
    
    Do While lIndex <= UBound(bData) - 8
        
        ' // Get opcode information
        GetMem4 bData(lIndex), lOpcode: lIndex = lIndex + 4
        GetMem4 bData(lIndex), lCount:  lIndex = lIndex + 4
        
        If lCount > 0 Then
        
            ' // Copy opcode code to memory
            memcpy ByVal pCurAddr, bData(lIndex), lCount:    lIndex = lIndex + lCount
            
            ' // Relocations
            GetMem4 bData(lIndex), lRelCount:   lIndex = lIndex + 4
            
            Do While lRelCount > 0
                
                GetMem4 bData(lIndex), lRelValue:   lIndex = lIndex + 4
                GetMem4 CLng(pAddrOfErrRaiseProc - (pCurAddr + lRelValue + 4)), ByVal pCurAddr + lRelValue
                
                lRelCount = lRelCount - 1
                
            Loop
            
            If lOpcode < &HFB Then
                lSubIndex = 0
            Else
                lSubIndex = (lOpcode And &HFF00&) \ &H100 - &HFB + 1
                If lSubIndex > UBound(pNewOpcodes, 2) Then Exit Function
            End If
            
            pNewOpcodes(lOpcode And &HFF&, lSubIndex) = pCurAddr
            
            pCurAddr = pCurAddr + lCount
            
            ' // Make JMP to next opcode handler
            GetMem4 &HE9&, ByVal pCurAddr
            
            lJmpValue = pAddrOfNextOpcodeProc - (pCurAddr + SIZEOF_JMP)
            
            GetMem4 lJmpValue, ByVal pCurAddr + 1
            
            pCurAddr = pCurAddr + SIZEOF_JMP
        
        Else
        
            If lOpcode < &HFB Then
                lSubIndex = 0
            Else
                lSubIndex = (lOpcode And &HFF00&) \ &H100 - &HFB + 1
                If lSubIndex > UBound(pNewOpcodes, 2) Then Exit Function
            End If
            
            pNewOpcodes(lOpcode And &HFF&, lSubIndex) = pAddrOfNextOpcodeProc
            
        End If
        
    Loop

    LoadNewOpcodes = True
    
    Exit Function
    
error_handler:
    
    ErrorLog "modCheckingOptions:LoadNewOpcodes"
    
End Function

' // Get DWORD by pointer
Private Function GetDWORD( _
                 ByVal pAddr As Long) As Long
                     
    If pAddr <= 0 Then Exit Function
    GetMem4 ByVal pAddr, GetDWORD
    
End Function

' // Create pointer
Private Function PtGet( _
                 ByRef Pointer() As Long, _
                 ByVal VarAddr As Long) As Currency
    Dim i As Long, i2 As Long
    i = GetDWORD(ArrPtr(Pointer)) + &HC
    i2 = VarPtr(PtGet)
    GetMem4 ByVal i, ByVal i2
    GetMem4 VarAddr + &HC, ByVal i
    GetMem4 Pointer(0), ByVal i2 + 4
End Function

' // Release pointer
Private Sub PtRelease( _
            ByRef Pointer() As Long, _
            ByRef prev As Currency)
    Dim i As Long
    i = VarPtr(prev)
    GetMem4 ByVal i + 4, Pointer(0)
    GetMem4 ByVal i, ByVal (Not Not Pointer) + &HC
End Sub


