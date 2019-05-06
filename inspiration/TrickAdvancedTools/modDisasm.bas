Attribute VB_Name = "modDisasm"
' //   Opcode Length Disassembler.
' //   Coded By Ms-Rem ( Ms-Rem@yandex.ru ) ICQ 286370715
' //   Port By The trick 2015

Option Explicit

Private Enum OpcodeData
    OP_NONE = &H0
    OP_MODRM = &H1
    OP_DATA_I8 = &H2
    OP_DATA_I16 = &H4
    OP_DATA_I32 = &H8
    OP_DATA_PRE66_67 = &H10
    OP_WORD = &H20
    OP_REL32 = &H40
End Enum

Dim OpcodeFlags(255)    As Byte
Dim OpcodeFlagsExt(255) As Byte
Dim IsInitialize        As Boolean

Public Function GetInstructionSize( _
                ByVal lpData As Long, _
                ByRef retOpcode As Long) As Long
                
    Dim cPtr        As Long:        Dim cPtrData    As Byte
    Dim PFX66       As Boolean:     Dim PFX67       As Boolean
    Dim flags       As OpcodeData:  Dim Opcode      As Byte
    Dim SibPresent  As Boolean:     Dim iMod        As Byte
    Dim iRM         As Byte:        Dim iReg        As Byte
    Dim OffsetSize  As Byte:        Dim Add         As Byte
    
    If Not IsInitialize Then Initialize
    
    cPtr = lpData
    
    GetMem1 ByVal cPtr, cPtrData
    
    Do While (cPtrData = &H2E Or cPtrData = &H3E Or cPtrData = &H36 Or _
              cPtrData = &H26 Or cPtrData = &H64 Or cPtrData = &H65 Or _
              cPtrData = &HF0 Or cPtrData = &HF2 Or cPtrData = &HF3 Or _
              cPtrData = &H66 Or cPtrData = &H67)
        
        If cPtrData = &H66 Then PFX66 = True
        If cPtrData = &H67 Then PFX67 = True
        
        cPtr = cPtr + 1
        
        If cPtr - lpData > 16 Then Exit Function
        
        GetMem1 ByVal cPtr, cPtrData
        
    Loop
    
    Opcode = cPtrData

    If Opcode = &HF Then
        ' // Two bytes
        cPtr = cPtr + 1
        GetMem1 ByVal cPtr, cPtrData
        flags = OpcodeFlagsExt(cPtrData)
        retOpcode = Opcode Or (cPtrData * &H100&)
    Else
        retOpcode = Opcode
        flags = OpcodeFlags(Opcode)
        If Opcode >= &HA0 And Opcode <= &HA3 Then PFX66 = PFX67
    End If
    
    cPtr = cPtr + 1
    If flags And OP_WORD Then cPtr = cPtr + 1
    
    GetMem1 ByVal cPtr, cPtrData
    
    If flags And OP_MODRM Then

        iMod = cPtrData \ 64
        iReg = (cPtrData And &H38) \ 8
        iRM = cPtrData And &H7
        
        cPtr = cPtr + 1: GetMem1 ByVal cPtr, cPtrData
        
        If ((Opcode = &HF6) And Not CBool(iReg)) Then flags = flags Or OP_DATA_I8
        If ((Opcode = &HF7) And Not CBool(iReg)) Then flags = flags Or OP_DATA_PRE66_67
 
        SibPresent = Not PFX67 And (iRM = 4)
        
        Select Case iMod
        Case 0
            If (PFX67 And (iRM = 6)) Then OffsetSize = 2
            If (Not PFX67 And (iRM = 5)) Then OffsetSize = 4
        Case 1
            OffsetSize = 1
        Case 2
            If PFX67 Then OffsetSize = 2 Else OffsetSize = 4
        Case 3
            SibPresent = False
        End Select
        
        If SibPresent Then
        
            If (((cPtrData And 7) = 5) And ((Not CBool(iMod)) Or (iMod = 2))) Then OffsetSize = 4
            cPtr = cPtr + 1
            
        End If
        
        cPtr = cPtr + OffsetSize
    
    End If
    
    If (flags And OP_DATA_I8) Then cPtr = cPtr + 1
    If (flags And OP_DATA_I16) Then cPtr = cPtr + 2
    If (flags And OP_DATA_I32) Then cPtr = cPtr + 4
    If (PFX66) Then Add = 2 Else Add = 4
    If (flags And OP_DATA_PRE66_67) Then cPtr = cPtr + Add
    
    GetInstructionSize = cPtr - lpData
    
End Function

Private Sub Initialize()
    Dim buf() As Byte
    Dim idx   As Long
    
    buf = LoadResData("OPCODES", "CUSTOM")
    
    For idx = 0 To 255
        OpcodeFlagsExt(idx) = buf(idx)
    Next
    
    For idx = 0 To 255
        OpcodeFlags(idx) = buf(idx + 256)
    Next
    
End Sub
