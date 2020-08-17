Attribute VB_Name = "EvaluatorTest"
Private ops(0 To 600) As Operation
Private opsPtr As Long
#Const devMode = False

Private Enum IType
    notSpecified = 0
    push = 1
    acces = 2
    binary = 3
    'fake
    Fake1 = 99
    Fake2 = 999
    Fake3 = 9999
    Fake4 = 99999
    Fake5 = 999999
    Fake6 = 9999999
    Fake7 = 99999999
    Fake8 = 999999999
End Enum
Private Enum ISubType
    'Arithmatic
    opAdd = 1
    opSub = 2
    opMul = 3
    opDiv = 4
    opPow = 5
    'Logic
    OpAnd = 6
    OpOr = 7
    OpXor = 8
    OpIIf = 9
    'string
    OpCat = 10
    OpLike = 11
    'comparison
    OpEql = 12
    OpNeq = 13
    opLt = 14
    opLte = 15
    opGt = 16
    opGte = 17
End Enum
Private Type Operation
    Type As IType
    subType As ISubType
    value As Variant
End Type

Private Sub pushV(ByRef stack() As Variant, ByRef index As Long, ByVal item As Variant)
    Dim size As Long: size = UBound(stack)
    #If Not devMode Then
        If index > size Then
            ReDim Preserve stack(0 To size * 2)
        End If
    #End If
    stack(index) = item
    index = index + 1
End Sub

Private Function popV(ByRef stack() As Variant, ByRef index As Variant) As Variant
    Dim size As Long: size = UBound(stack)
    #If Not devMode Then
    If index < size / 3 Then
        ReDim Preserve stack(0 To CLng(size / 2))
    End If
    #End If
    index = index - 1
    popV = stack(index)
    stack(index) = Empty
End Function

Private Sub pushO(ByRef stack() As Operation, ByRef index As Long, ByRef item As Operation)
    Dim size As Long: size = UBound(stack)
    #If Not devMode Then
    If index > size Then
        ReDim Preserve stack(0 To size * 2)
    End If
    #End If
    stack(index) = item
    index = index + 1
End Sub

Private Function popO(ByRef stack() As Operation, ByRef index As Variant) As Operation
    Dim size As Long: size = UBound(stack)
    #If Not devMode Then
    If index < size / 3 Then
        ReDim Preserve stack(0 To CLng(size / 2))
    End If
    #End If
    index = index - 1
    popO = stack(index)
    stack(index) = Empty
End Function

Private Function execute(params() As Operation) As Variant
    Dim stack() As Variant
    #If devMode Then
        ReDim stack(0 To 100)
    #Else
        ReDim stack(0 To 4)
    #End If
    Dim stackPtr As Long: stackPtr = 0
    For opIndex = 0 To UBound(params)
        Dim op As Operation: op = params(opIndex)
        Select Case op.Type
            Case IType.push
                Call pushV(stack, stackPtr, op.value)
            Case IType.Fake1
            Case IType.Fake2
            Case IType.Fake3
            Case IType.Fake4
            Case IType.Fake5
            Case IType.Fake6
            Case IType.Fake7
            Case IType.Fake8
            Case IType.binary
                Dim v2 As Variant: v2 = popV(stack, stackPtr)
                Dim v1 As Variant: v1 = popV(stack, stackPtr)
                Dim result As Variant
                Select Case op.subType
                    'Arithmatic
                    Case ISubType.opAdd
                        result = v1 + v2
                    Case ISubType.opSub
                        result = v1 - v2
                    Case ISubType.opMul
                        result = v1 * v2
                    Case ISubType.opDiv
                        result = v1 / v2
                    Case ISubType.opPow
                        result = v1 ^ v2
                End Select
                Call pushV(stack, stackPtr, result)
            Case IType.notSpecified
                Exit For
        End Select
    Next
    execute = stack(0)
End Function

Private Sub addOp(kType As IType, Optional subType As ISubType, Optional value As Variant)
    With ops(opsPtr)
        .Type = kType
        .subType = subType
        .value = value
    End With
    opsPtr = opsPtr + 1
End Sub

Sub test()
    '(3*(2+5)+5*8/2^(2+1))/26-1=0
    opsPtr = 0
    Call addOp(push, , 0)
    For i = 0 To 8
        Call addOp(push, , 3)
        Call addOp(push, , 2)
        Call addOp(push, , 5)
        Call addOp(binary, opAdd)
        Call addOp(binary, opMul)
        Call addOp(push, , 5)
        Call addOp(push, , 8)
        Call addOp(binary, opMul)
        Call addOp(push, , 2)
        Call addOp(push, , 2)
        Call addOp(push, , 1)
        Call addOp(binary, opAdd)
        Call addOp(binary, opPow)
        Call addOp(binary, opDiv)
        Call addOp(binary, opAdd)
        Call addOp(push, , 26)
        Call addOp(binary, opDiv)
        Call addOp(push, , 1)
        Call addOp(binary, opSub)
        Call addOp(binary, opAdd)
    Next
        
    'Test1
    x = Timer
        For i = 0 To 10 ^ 4
            Call execute(ops)
        Next
    Debug.Print Timer - x
    
'    'Test2
    expression = "0+(3*(2+5)+5*8/2^(2+1))/26-1+(3*(2+5)+5*8/2^(2+1))/26-1+(3*(2+5)+5*8/2^(2+1))/26-1+(3*(2+5)+5*8/2^(2+1))/26-1+(3*(2+5)+5*8/2^(2+1))/26-1+(3*(2+5)+5*8/2^(2+1))/26-1+(3*(2+5)+5*8/2^(2+1))/26-1+(3*(2+5)+5*8/2^(2+1))/26-1"
'    With stdLambda.Create(expression)
'        x = Timer
'            For i = 0 To 10 ^ 4
'                Call .Run
'            Next
'        Debug.Print Timer - x
'    End With
    
    'test3
    With Application
        x = Timer
            For i = 0 To 10 ^ 4
                Call .Evaluate(expression & "+" & i)
            Next
        Debug.Print Timer - x
    End With
End Sub
