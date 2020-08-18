Attribute VB_Name = "EvaluatorTest"
'Direct call convention of VBA.CallByName
#If VBA7 Then
  Private Declare PtrSafe Function rtcCallByName Lib "VBE7.dll" (ByRef vRet As Variant, ByVal cObj As Object, ByVal sMethod As LongPtr, ByVal eCallType As VbCallType, ByRef pArgs() As Variant, ByVal lcid As Long) As Long
  Private Declare PtrSafe Sub VariantCopy Lib "oleaut32.dll" (ByRef pvargDest As Variant, ByRef pvargSrc As Variant)
#ElseIf VBA6 Then
  Private Declare PtrSafe Function rtcCallByName Lib "msvbvm60" (ByRef vRet As Variant, ByVal cObj As Object, ByVal sMethod As LongPtr, ByVal eCallType As VbCallType, ByRef pArgs() As Variant, ByVal lcid As Long) As Long
  Private Declare PtrSafe Sub VariantCopy Lib "oleaut32.dll" (ByRef pvargDest As Variant, ByRef pvargSrc As Variant)
#Else
  Private Declare Function rtcCallByName Lib "msvbvm60" (ByRef vRet As Variant, ByVal cObj As Object, ByVal sMethod As LongPtr, ByVal eCallType As VbCallType, ByRef pArgs() As Variant, ByVal lcid As Long) As Long
  Private Declare Sub VariantCopy Lib "oleaut32.dll" (ByRef pvargDest As Variant, ByRef pvargSrc As Variant)
#End If


Private ops(0 To 600) As Operation
Private opsPtr As Long
Const minStackSize = 30 'note that the stack size may become smaller than this
Const devStackSize = 200
#Const devMode = True

Private Enum iType
    oPush = 1
    oPop = 2
    oMerge = 3
    oAccess = 4
    oSet = 5
    oArithmetic = 6
    oLogic = 7
    oComparison = 8
    oMisc = 9
    oJump = 10
    oReturn = 11
    oObject = 12
End Enum
Private Enum ISubType
    'Arithmetic
    oAdd = 1
    oSub = 2
    oMul = 3
    oDiv = 4
    oPow = 5
    oNeg = 6
    'Logic
    oAnd = 7
    oOr = 8
    oNot = 9
    oXor = 10
    'comparison
    oEql = 11
    oNeq = 12
    oLt = 13
    oLte = 14
    oGt = 15
    oGte = 16
    'misc operators
    oCat = 17
    oLike = 18
    'misc
    ifTrue = 19
    ifFalse = 20
    withValue = 21
    'object
    oPropGet = 22
    oPropLet = 23
    oPropSet = 24
    oMethodCall = 25
    oEquality = 26
    oIsOperator = 27
    oEnum = 28 '
End Enum
Private Type Operation
    Type As iType
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
    Call CopyVariant(stack(index), item)
    index = index + 1
End Sub

Private Function popV(ByRef stack() As Variant, ByRef index As Variant) As Variant
    Dim size As Long: size = UBound(stack)
    #If Not devMode Then
        If index < size / 3 And index < minStackSize Then
            ReDim Preserve stack(0 To CLng(size / 2))
        End If
    #End If
    index = index - 1
    Call CopyVariant(popV, stack(index))
    #If devMode Then
        stack(index) = Empty
    #End If
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
End Function

Private Function execute(ByRef operations() As Operation) As Variant
    Dim stack() As Variant
    #If devMode Then
        ReDim stack(0 To devStackSize)
    #Else
        ReDim stack(0 To 4)
    #End If
    Dim stackPtr As Long: stackPtr = 0
    
    Dim op As Operation
    Dim v1 As Variant
    Dim v2 As Variant
    Dim v3 As Variant
    Dim opIndex As Long: opIndex = 0
    Dim opCount As Long: opCount = UBound(operations)
    
    While opIndex < opCount
        op = operations(opIndex)
        opIndex = opIndex + 1
        Select Case op.Type
            Case iType.oPush
                Call pushV(stack, stackPtr, op.value)
            'Arithmetic
            Case iType.oArithmetic
                v2 = popV(stack, stackPtr)
                Select Case op.subType
                    Case ISubType.oAdd
                        v1 = popV(stack, stackPtr)
                        v3 = v1 + v2
                    Case ISubType.oSub
                        v1 = popV(stack, stackPtr)
                        v3 = v1 - v2
                    Case ISubType.oMul
                        v1 = popV(stack, stackPtr)
                        v3 = v1 * v2
                    Case ISubType.oDiv
                        v1 = popV(stack, stackPtr)
                        v3 = v1 / v2
                    Case ISubType.oPow
                        v1 = popV(stack, stackPtr)
                        v3 = v1 ^ v2
                    Case ISubType.oNeg
                        v3 = -v2
                    Case Else
                        v3 = Empty
                End Select
                Call pushV(stack, stackPtr, v3)
            'Comparison
            Case iType.oComparison
                v2 = popV(stack, stackPtr)
                v1 = popV(stack, stackPtr)
                Select Case op.subType
                    Case ISubType.oEql
                        v3 = v1 = v2
                    Case ISubType.oNeq
                        v3 = v1 <> v2
                    Case ISubType.oGt
                        v3 = v1 > v2
                    Case ISubType.oGte
                        v3 = v1 >= v2
                    Case ISubType.oLt
                        v3 = v1 < v2
                    Case ISubType.oLte
                        v3 = v1 <= v2
                    Case Else
                        v3 = Empty
                End Select
                Call pushV(stack, stackPtr, v3)
            'Logic
            Case iType.oLogic
                v2 = popV(stack, stackPtr)
                Select Case op.subType
                    Case ISubType.oAnd
                        v1 = popV(stack, stackPtr)
                        v3 = v1 And v2
                    Case ISubType.oOr
                        v1 = popV(stack, stackPtr)
                        v3 = v1 Or v2
                    Case ISubType.oNot
                        v3 = Not v2
                    Case ISubType.oXor
                        v1 = popV(stack, stackPtr)
                        v3 = v1 Xor v2
                    Case Else
                        v3 = Empty
                End Select
                Call pushV(stack, stackPtr, v3)
            'Object
            Case iType.oObject
                Call objectCaller(stack, stackPtr, op)
            'Misc
            Case iType.oMisc
                v2 = popV(stack, stackPtr)
                v1 = popV(stack, stackPtr)
                Select Case op.subType
                    Case ISubType.oCat
                        v3 = v1 & v2
                    Case ISubType.oLike
                        v3 = v1 Like v2
                    Case Else
                        v3 = Empty
                End Select
                Call pushV(stack, stackPtr, v3)
            'Variable
            Case iType.oAccess
                Call pushV(stack, stackPtr, stack(stackPtr - op.value))
            Case iType.oSet
                v1 = popV(stack, stackPtr)
                stack(stackPtr - op.value) = v1
            'Flow
            Case iType.oJump
                Select Case op.subType
                    Case ISubType.ifTrue
                        v1 = popV(stack, stackPtr)
                        If v1 Then
                            opIndex = op.value
                        End If
                    Case ISubType.ifFalse
                        v1 = popV(stack, stackPtr)
                        If Not v1 Then
                            opIndex = op.value
                        End If
                    Case Else
                        opIndex = op.value
                End Select
            Case iType.oReturn
                Select Case op.subType
                    Case ISubType.withValue
                        v1 = popV(stack, stackPtr)
                        opIndex = stack(stackPtr - 1)
                        stack(stackPtr - 1) = v1
                    Case Else
                        opIndex = popV(stack, stackPtr)
                End Select
            'Data
            Case iType.oMerge
                Call CopyVariant(v1, popV(stack, stackPtr))
                Call CopyVariant(stack(stackPtr - 1), v1)
            Case iType.oPop
                Call popV(stack, stackPtr)
            Case Else
                opIndex = 10 ^ 6 'TODO: replace by infinity or something
        End Select
    Wend
    execute = stack(0)
End Function


'Calls an object method/setter/getter/letter
'@param {ByRef Variant()} stack     The stack to get the data from and add the result to
'@param {ByRef Long} stackPtr       The pointer that indicates the position of the top of the stack
'@param {ByRef Operation} op        The operation to execute
'@returns {void}
Private Sub objectCaller(ByRef stack() As Variant, ByRef stackPtr As Long, ByRef op As Operation)
    'Get the arguments
    Dim funcName As Variant: funcName = popV(stack, stackPtr)
    Dim argCount As Variant
    Dim args() As Variant
    If VarType(funcName) = vbString Then
        'If no argument count is specified, there are no arguments
        argCount = 0
        args = Array()
    Else
        'If an argument count is provided, extract all arguments into an array
        argCount = funcName
        ReDim args(1 To argCount)
        For i = 1 To argCount
            Call CopyVariant(args(i), popV(stack, stackPtr))
        Next
        funcName = popV(stack, stackPtr)
    End If
    
    'Get caller type
    Dim callerType As VbCallType
    Select Case op.subType
        Case ISubType.oPropGet:     callerType = VbGet
        Case ISubType.oMethodCall:  callerType = VbMethod
        Case ISubType.oPropLet:     callerType = VbLet
        Case ISubType.oPropSet:     callerType = VbSet
    End Select
                
    'Call rtcCallByName
    Dim hr As Long, res As Variant, obj As Object
    Set obj = popV(stack, stackPtr)
    hr = rtcCallByName(res, obj, StrPtr(funcName), callerType, args, &H409)
    
    'If error then raise, otherwise push stack
    If hr < 0 Then
        Call Throw("Error in calling " & sFuncName & " property of " & TypeName(value) & " object.")
    Else
        If op.subType = ISubType.oPropGet Or op.subType = ISubType.oMethodCall Then
            Call pushV(stack, stackPtr, res)
        End If
    End If
End Sub



' =============================================
'
' Shit below only for testing, not needed later
'
' =============================================


Private Sub addOp(kType As iType, Optional subType As ISubType, Optional value As Variant)
    With ops(opsPtr)
        .Type = kType
        .subType = subType
        Call CopyVariant(.value, value)
    End With
    opsPtr = opsPtr + 1
End Sub

'Throws an error
'@param {string} The error message to be thrown
'@returns {void}
Private Sub Throw(ByVal sMessage As String)
    MsgBox sMessage, vbCritical
    End
End Sub

'Copies one variant to a destination
'@param {ByRef Variant} dest Destination to copy variant to
'@param {Variant} value Source to copy variant from.
Private Sub CopyVariant(ByRef dest As Variant, ByVal value As Variant)
  If IsObject(value) Then
    Set dest = value
  Else
    dest = value
  End If
End Sub

Sub objectTest()
    'Set stdLambda.Create("1+" & "1").oFunctExt = new Collection
    opsPtr = 0
    Call addOp(oPush, , stdLambda)
    Call addOp(oPush, , "Create")
    Call addOp(oPush, , "1+")
    Call addOp(oPush, , "1")
    Call addOp(oMisc, oCat)
    Call addOp(oPush, , 1)
    Call addOp(oObject, oMethodCall)
    Call addOp(oPush, , "oFunctExt")
    Call addOp(oPush, , New Collection)
    Call addOp(oPush, , 1)
    Call addOp(oObject, oPropSet)
        
    Debug.Print execute(ops)
    Debug.Print TypeName(stdLambda.Create("1+1").oFunctExt)
End Sub

Sub functionTest()
    'fib(v) {
    '   if (v<=1) return v
    '   return fib(v-1)+fib(v-2)
    '}
    'fib(6)
    opsPtr = 0
    Call addOp(oJump, , 19)             'Should skip function declaration
    Call addOp(oAccess, , 1)            'Get the argument (argument was pushed to the stack)
    Call addOp(oPush, , 1)              'Push comparison value
    Call addOp(oComparison, oLte, 1)    'Compare argument with comparison value
    Call addOp(oJump, ifFalse, 6)       'Skip following line if comparison yields false
    Call addOp(oReturn, withValue)      'Return
    Call addOp(oPush, , 11)             'Push the return address to the stack
    Call addOp(oAccess, , 2)            'Retrieve the argument of this call
    Call addOp(oPush, , 1)              'Push constant to subtract onto the stack
    Call addOp(oArithmetic, oSub)       'Subtract the contsntant
    Call addOp(oJump, , 1)              'Make recurse function call
    Call addOp(oPush, , 16)             'Push the return address to the stack
    Call addOp(oAccess, , 3)            'Retrieve the argument of this call
    Call addOp(oPush, , 2)              'Push constant to subtract onto the stack
    Call addOp(oArithmetic, oSub)       'Subtract the contsntant
    Call addOp(oJump, , 1)              'Make recurse function call
    Call addOp(oArithmetic, oAdd)       'Add the values up
    Call addOp(oMerge)                  'Remove argument
    Call addOp(oReturn, withValue)      'Return
    Call addOp(oPush, , 22)             'Push return address
    Call addOp(oPush, , 22)             'Add argument for initial call
    Call addOp(oJump, , 1)              'Perform initial call
        
    Debug.Print execute(ops)
End Sub


Sub performanceTest()
    '(3*(2+5)+5*8/2^(2+1))/26-1=0
    opsPtr = 0
    Call addOp(oPush, , 0)
    For i = 0 To 8
        Call addOp(oPush, , 3)
        Call addOp(oPush, , 2)
        Call addOp(oPush, , 5)
        Call addOp(oArithmetic, oAdd)
        Call addOp(oArithmetic, oMul)
        Call addOp(oPush, , 5)
        Call addOp(oPush, , 8)
        Call addOp(oArithmetic, oMul)
        Call addOp(oPush, , 2)
        Call addOp(oPush, , 2)
        Call addOp(oPush, , 1)
        Call addOp(oArithmetic, oAdd)
        Call addOp(oArithmetic, oPow)
        Call addOp(oArithmetic, oDiv)
        Call addOp(oArithmetic, oAdd)
        Call addOp(oPush, , 26)
        Call addOp(oArithmetic, oDiv)
        Call addOp(oPush, , 1)
        Call addOp(oArithmetic, oSub)
        Call addOp(oArithmetic, oAdd)
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
