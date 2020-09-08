VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdLambda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Implement stdICallable interface
Implements stdICallable

'Direct call convention of VBA.CallByName
#If VBA7 Then
  'VBE7 is interchangable with msvbvm60.dll    however VBE7.dll appears to always be present where as msvbvm60 is only occasionally present.
  Private Declare PtrSafe Function rtcCallByName Lib "VBE7.dll" (ByVal cObj As Object, ByVal sMethod As LongPtr, ByVal eCallType As VbCallType, ByRef pArgs() As Variant, ByVal lcid As Long) As Variant
#Else
  Private Declare Function rtcCallByName Lib "msvbvm60" (ByVal cObj As Object, ByVal sMethod As LongPtr, ByVal eCallType As VbCallType, ByRef pArgs() As Variant, ByVal lcid As Long) As Variant
#End If
#Const devMode = True 'Used for enabling some debugging features

'Tokens, token definitions and operations
Private Type TokenDefinition
    name As String
    Regex As String
    RegexObj As Object
End Type
Private Type token
    Type As TokenDefinition
    value As Variant
    BracketDepth As Long
End Type
Private Type Operation
    Type As iType
    subType As ISubType
    value As Variant
End Type

'Evaluation operation types
Private Enum iType
    oPush = 1
    oPop = 2
    oMerge = 3
    oAccess = 4
    oSet = 5
    oArithmetic = 6
    oLogic = 7
    oFunc = 8
    oComparison = 9
    oMisc = 10
    oJump = 11
    oReturn = 12
    oObject = 13
End Enum
Private Enum ISubType
    'Arithmetic
    oAdd = 1
    oSub = 2
    oMul = 3
    oDiv = 4
    oPow = 5
    oNeg = 6
    oMod = 7
    'Logic
    oAnd = 8
    oOr = 9
    oNot = 10
    oXor = 11
    'comparison
    oEql = 12
    oNeq = 13
    oLt = 14
    oLte = 15
    oGt = 16
    oGte = 17
    oIs = 18
    'misc operators
    oCat = 19
    oLike = 20
    'misc
    ifTrue = 21
    ifFalse = 22
    withValue = 23
    argument = 24
    'object
    oPropGet = 25
    oPropLet = 26
    oPropSet = 27
    oMethodCall = 28
    oEquality = 29    'Yet to be implemented
    oIsOperator = 30  'Yet to be implemented
    oEnum = 31        'Yet to be implemented
End Enum

'Special constant used in parsing:
Const UniqueConst As String = "3207af79-30df-4890-ade1-640f9f28f309"

Private tokens() As token
Private iTokenIndex As Long
Private operations() As Operation
Private iOperationIndex As Long
Private stackSize As Long
Private scopes() As Variant
Private scopesArgCount() As Variant
Private scopeCount As Long
Private funcScope As Long


Const minStackSize = 30 'note that the stack size may become smaller than this
Public oFunctExt As Object 'Dictionary<string => stdCallback>
Private vLastArgs As Variant

Private oCache As Object


''Usage:
'Debug.Print stdLambda.Create("1+3*8/2*(2+2+3)").Execute()
'With stdLambda.Create("$1+1+3*8/2*(2+2+3)")
'    Debug.Print .Execute(10)
'    Debug.Print .Execute(15)
'    Debug.Print .Execute(20)
'End With
'Debug.Print stdLambda.Create("$1.Range(""A1"")").Execute(Sheets(1)).Address(True, True, xlA1, True)
'Debug.Print stdLambda.Create("$1#join("","")").Execute(stdArray.Create(1,2))
Public Function Create(ByVal sEquation As String, Optional ByVal bSandboxExtras As Boolean = False) As stdLambda
    'Cache Lambda created
    If oCache Is Nothing Then Set oCache = CreateObject("Scripting.Dictionary")
    Dim sID As String: sID = bSandboxExtras & ")" & sEquation
    If Not oCache.exists(sID) Then
        Set oCache(sID) = New stdLambda
        Call oCache(sID).Init(sEquation, bSandboxExtras)
    End If
    
    'Return cached lambda
    Set Create = oCache(sID)
End Function

Public Function CreateMultiline(ByRef sEquation As Variant, Optional ByVal bSandboxExtras As Boolean = False) As stdLambda
    Set CreateMultiline = Create(Join(sEquation, " "), bSandboxExtras)
End Function

Public Sub Init(ByVal sEquation As String, ByVal bSandboxExtras As Boolean)
    If Not bSandboxExtras Then
        Set Me.oFunctExt = stdLambda.oFunctExt
    End If
    tokens = Tokenise(sEquation)
    iTokenIndex = 1
    iOperationIndex = 0
    stackSize = 0
    scopeCount = 0
    funcScope = 0
    Call parseBlock("eof")
    Call finishOperations
End Sub

Private Function stdICallable_Run(ParamArray params() As Variant) As Variant
    'Bind args
    vLastArgs = params

    'Execute top-down parser
    Call CopyVariant(stdICallable_Run, evaluate(operations))
End Function
Private Function stdICallable_RunEx(params() As Variant) As Variant
    'Bind args
    vLastArgs = params

    'Execute top-down parser
    Call CopyVariant(stdICallable_RunEx, evaluate(operations))
End Function

Function Run(ParamArray params() As Variant) As Variant
    'Bind args
    vLastArgs = params

    'Execute top-down parser
    Call CopyVariant(Run, evaluate(operations))
End Function

Function RunEx(params() As Variant) As Variant
    'Bind args
    vLastArgs = params
    
    'Execute top-down parser
    Call CopyVariant(RunEx, evaluate(operations))
End Function



'================
'
'  TOKENISATION
'
'================


'Tokeniser helpers
Private Function getTokenDefinitions() As TokenDefinition()
    Dim arr() As TokenDefinition
    ReDim arr(1 To 99)
    
    Dim i As Long: i = 0
    'Whitespace
    i = i + 1: arr(i) = getTokenDefinition("space", "\s+") 'String

    'Literal
    i = i + 1: arr(i) = getTokenDefinition("literalString", """(?:""""|[^""])*""") 'String
    i = i + 1: arr(i) = getTokenDefinition("literalNumber", "\d+(?:\.\d+)?")   'Number
    i = i + 1: arr(i) = getTokenDefinition("literalBoolean", "True|False")
    
    'Named operators
    i = i + 1: arr(i) = getTokenDefinition("is", "is")
    i = i + 1: arr(i) = getTokenDefinition("mod", "mod")
    i = i + 1: arr(i) = getTokenDefinition("and", "and")
    i = i + 1: arr(i) = getTokenDefinition("or", "or")
    i = i + 1: arr(i) = getTokenDefinition("xor", "xor")
    i = i + 1: arr(i) = getTokenDefinition("not", "not")

    'Structural
    ' Inline if
    i = i + 1: arr(i) = getTokenDefinition("if", "if")
    i = i + 1: arr(i) = getTokenDefinition("then", "then")
    i = i + 1: arr(i) = getTokenDefinition("else", "else")
    i = i + 1: arr(i) = getTokenDefinition("end", "end")
    ' Brackets
    i = i + 1: arr(i) = getTokenDefinition("lBracket", "\(")
    i = i + 1: arr(i) = getTokenDefinition("rBracket", "\)")
    ' Functions
    i = i + 1: arr(i) = getTokenDefinition("fun", "fun")
    i = i + 1: arr(i) = getTokenDefinition("comma", ",") 'params
    ' Lines
    i = i + 1: arr(i) = getTokenDefinition("colon", ":")
    
    'VarName
    i = i + 1: arr(i) = getTokenDefinition("arg", "\$\d+")
    i = i + 1: arr(i) = getTokenDefinition("var", "[a-zA-Z][a-zA-Z0-9_]*")
    
    'Operators
    i = i + 1: arr(i) = getTokenDefinition("propertyAccess", "\.")
    i = i + 1: arr(i) = getTokenDefinition("methodAccess", "\#")
    i = i + 1: arr(i) = getTokenDefinition("multiply", "\*")
    i = i + 1: arr(i) = getTokenDefinition("divide", "\/")
    i = i + 1: arr(i) = getTokenDefinition("power", "\^")
    i = i + 1: arr(i) = getTokenDefinition("add", "\+")
    i = i + 1: arr(i) = getTokenDefinition("subtract", "\-")
    i = i + 1: arr(i) = getTokenDefinition("equal", "\=")
    i = i + 1: arr(i) = getTokenDefinition("notEqual", "\<\>")
    i = i + 1: arr(i) = getTokenDefinition("greaterThanEqual", "\>\=")
    i = i + 1: arr(i) = getTokenDefinition("greaterThan", "\>")
    i = i + 1: arr(i) = getTokenDefinition("lessThanEqual", "\<\=")
    i = i + 1: arr(i) = getTokenDefinition("lessThan", "\<")
    i = i + 1: arr(i) = getTokenDefinition("concatenate", "\&")
    
    ReDim Preserve arr(1 To i)

    getTokenDefinitions = arr
End Function

'===========
'
'  PARSING
'
'===========

Private Sub parseBlock(ParamArray endToken() As Variant)
    Call addScope
    Dim size As Integer: size = stackSize + 1
    
    ' Consume multiple lines
    Dim bLoop As Boolean: bLoop = True
    Do
        While optConsume("colon"): Wend
        Call parseStatement
        
        For i = LBound(endToken) To UBound(endToken)
            If peek(endToken(i)) Then
                bLoop = False
            End If
        Next
    Loop While bLoop
    
    ' Get rid of all extra expression results and declarations
    While stackSize > size
        Call addOperation(oMerge, , , -1)
    Wend
    scopeCount = scopeCount - 1
End Sub

Private Sub addScope()
    scopeCount = scopeCount + 1
    Dim scope As Long: scope = scopeCount
    ReDim Preserve scopes(1 To scope)
    ReDim Preserve scopesArgCount(1 To scope)
    Set scopes(scope) = CreateObject("Scripting.Dictionary")
    Set scopesArgCount(scope) = CreateObject("Scripting.Dictionary")
End Sub

Private Sub parseStatement()
    If peek("var") And peek("equal", 2) Then
        Call parseAssignment
    ElseIf peek("fun") Then
        Call parseFunctionDeclaration
    Else
        Call parseExpression
    End If
End Sub

Private Sub parseExpression()
    Call parseLogicPriority1
End Sub

Private Sub parseLogicPriority1() 'xor
    Call parseLogicPriority2
    Dim bLoop As Boolean: bLoop = True
    Do
        If optConsume("xor") Then
            Call parseLogicPriority2
            Call addOperation(oLogic, oXor, , -1)
        Else
            bLoop = False
        End If
    Loop While bLoop
End Sub

Private Sub parseLogicPriority2() 'or
    Call parseLogicPriority3
    Dim bLoop As Boolean: bLoop = True
    Do
        If optConsume("or") Then
            Call parseLogicPriority3
            Call addOperation(oLogic, oOr, , -1)
        Else
            bLoop = False
        End If
    Loop While bLoop
End Sub

Private Sub parseLogicPriority3() 'and
    Call parseLogicPriority4
    Dim bLoop As Boolean: bLoop = True
    Do
        If optConsume("and") Then
            Call parseLogicPriority4
            Call addOperation(oLogic, oOr, , -1)
        Else
            bLoop = False
        End If
    Loop While bLoop
End Sub

Private Sub parseLogicPriority4() 'not
    Dim invert As Variant: invert = vbNull
    While optConsume("not")
        If invert = vbNull Then invert = False
        invert = Not invert
    Wend
    
    Call parseComparisonPriority1
    
    If invert <> vbNull Then
        Call addOperation(oLogic, oNot)
        If invert = False Then
            Call addOperation(oLogic, oNot)
        End If
    End If
End Sub

Private Sub parseComparisonPriority1() '=, <>, <, <=, >, >=, is, Like
    Call parseArithmeticPriority1
    Dim bLoop As Boolean: bLoop = True
    Do
        If optConsume("lessThan") Then
            Call parseArithmeticPriority1
            Call addOperation(oComparison, oLt, , -1)
        ElseIf optConsume("lessThanEqual") Then
            Call parseArithmeticPriority1
            Call addOperation(oComparison, oLte, , -1)
        ElseIf optConsume("greaterThan") Then
            Call parseArithmeticPriority1
            Call addOperation(oComparison, oGt, , -1)
        ElseIf optConsume("greaterThanEqual") Then
            Call parseArithmeticPriority1
            Call addOperation(oComparison, oGte, , -1)
        ElseIf optConsume("equal") Then
            Call parseArithmeticPriority1
            Call addOperation(oComparison, oEql, , -1)
        ElseIf optConsume("notEqual") Then
            Call parseArithmeticPriority1
            Call addOperation(oComparison, oNeq, , -1)
        ElseIf optConsume("is") Then
            Call parseArithmeticPriority1
            Call addOperation(oComparison, oIs, , -1)
        ElseIf optConsume("like") Then
            Call parseArithmeticPriority1
            Call addOperation(oComparison, oLike, , -1)
        Else
            bLoop = False
        End If
    Loop While bLoop
End Sub

Private Sub parseArithmeticPriority1() '&
    Call parseArithmeticPriority2
    Dim bLoop As Boolean: bLoop = True
    Do
        If optConsume("concatenate") Then
            Call parseArithmeticPriority2
            Call addOperation(oMisc, oCat, , -1)
        Else
            bLoop = False
        End If
    Loop While bLoop
End Sub

Private Sub parseArithmeticPriority2() '+, -
    Call parseArithmeticPriority3
    Dim bLoop As Boolean: bLoop = True
    Do
        If optConsume("add") Then
            Call parseArithmeticPriority3
            Call addOperation(oArithmetic, oAdd, , -1)
        ElseIf optConsume("subtract") Then
            Call parseArithmeticPriority3
            Call addOperation(oArithmetic, oSub, , -1)
        Else
            bLoop = False
        End If
    Loop While bLoop
End Sub


Private Sub parseArithmeticPriority3() 'mod
    Call parseArithmeticPriority4
    Dim bLoop As Boolean: bLoop = True
    Do
        If optConsume("modulus") Then
            Call parseArithmeticPriority4
            Call addOperation(oArithmetic, oMod, , -1)
        Else
            bLoop = False
        End If
    Loop While bLoop
End Sub

Private Sub parseArithmeticPriority4() '*, /
    Call parseArithmeticPriority5
    Dim bLoop As Boolean: bLoop = True
    Do
        If optConsume("multiply") Then
            Call parseArithmeticPriority4
            Call addOperation(oArithmetic, oMul, , -1)
        ElseIf optConsume("divide") Then
            Call parseArithmeticPriority4
            Call addOperation(oArithmetic, oDiv, , -1)
        Else
            bLoop = False
        End If
    Loop While bLoop
End Sub

Private Sub parseArithmeticPriority5() '+, - (unary)
    If optConsume("subtract") Then
        Call parseArithmeticPriority5 'recurse
        Call addOperation(oArithmetic, oNeg)
    ElseIf optConsume("add") Then
        Call parseArithmeticPriority5 'recurse
    Else
        Call parseArithmeticPriority6
    End If
End Sub

Private Sub parseArithmeticPriority6() '^
    Call parseFlowPriority1
    Do
        If optConsume("power") Then
            Call parseArithmeticPriority6andahalf '- and + are still identity operators
            Call addOperation(oArithmetic, oPow, , -1)
        Else
            bLoop = False
        End If
    Loop While bLoop
End Sub

Private Sub parseArithmeticPriority6andahalf() '+, - (unary)
    If optConsume("subtract") Then
        Call parseArithmeticPriority6andahalf 'recurse
        Call addOperation(oArithmetic, oNeg)
    ElseIf optConsume("add") Then
        Call parseArithmeticPriority6andahalf 'recurse
    Else
        Call parseFlowPriority1
    End If
End Sub

Private Sub parseFlowPriority1() 'if then else
    If optConsume("if") Then
        Call parseExpression
        Dim skipThenJumpIndex As Integer: skipThenJumpIndex = addOperation(oJump, ifFalse, , -1)
        
        Dim size As Integer: size = stackSize
        Call consume("then")
        Call parseBlock("else", "end")
        Dim skipElseJumpIndex As Integer: skipElseJumpIndex = addOperation(oJump)
        operations(skipThenJumpIndex).value = iOperationIndex
        stackSize = size
        
        If optConsume("end") Then
            Call addOperation(oPush, , 0, 1) 'Expressions should always return a value
            operations(skipElseJumpIndex).value = iOperationIndex
        Else
            Call consume("else")
            Call parseBlock("eof", "rBracket", "end")
            operations(skipElseJumpIndex).value = iOperationIndex
        
            Call optConsume("end")
        End If
    Else
        Call parseValuePriority1
    End If
End Sub

Private Sub parseValuePriority1() 'numbers, $vars, strings, booleans, (expressions)

    If peek("literalNumber") Then
        Call addOperation(oPush, , CDbl(consume("literalNumber")), 1)
    ElseIf peek("arg") Then
        Call addOperation(oAccess, argument, consume("arg"), 1)
        Call parseManyAccessors
    ElseIf peek("literalString") Then
        Call parseString
    ElseIf peek("literalBoolean") Then
        Call addOperation(oPush, , consume("literalBoolean") = "true", 1)
    ElseIf peek("var") Then
        If Not parseScopeAccess Then
            Call parseFunction
        End If
        Call parseManyAccessors
    Else
        Call consume("lBracket")
        Call parseExpression
        Call consume("rBracket")
        Call parseManyAccessors
    End If
End Sub

Private Function parseFunction() As Variant
    Call addOperation(oPush, , consume("var"), 1)
    Dim size As Integer: size = stackSize
    Call parseOptParameters
    Call addOperation(oFunc)
    stackSize = size
End Function

Private Sub parseManyAccessors()
    Dim bLoop As Boolean: bLoop = True
    Do
        bLoop = False
        If parseOptObjectProperty() Then bLoop = True
        If parseOptObjectMethod() Then bLoop = True
    Loop While bLoop
End Sub

Private Function parseOptObjectProperty() As Boolean
    parseOptObjectProperty = False
    If optConsume("propertyAccess") Then
        Dim size As Integer: size = stackSize
        Call addOperation(oPush, , consume("var"), 1)
        Call parseOptParameters
        Call addOperation(oObject, oPropGet)
        stackSize = size
        parseOptObjectProperty = True
    End If
End Function

Private Function parseOptObjectMethod() As Boolean
    parseOptObjectMethod = False
    If optConsume("methodAccess") Then
        Dim size As Integer: size = stackSize
        Call addOperation(oPush, , consume("var"), 1)
        Call parseOptParameters
        Call addOperation(oObject, oMethodCall)
        stackSize = size
        parseOptObjectMethod = True
    End If
End Function

Private Function parseOptParameters() As Boolean
    parseOptParameters = False
    If optConsume("lBracket") Then
        Dim iArgCount As Integer
        While Not peek("rBracket")
            If iArgCount > 0 Then
                Call consume("comma")
            End If
            Call parseExpression
            iArgCount = iArgCount + 1
        Wend
        Call consume("rBracket")
        If iArgCount > 0 Then
            Call addOperation(oPush, , iArgCount, 1)
        End If
        parseOptParameters = True
    End If
End Function

Private Sub parseString()
    Dim sRes As String: sRes = consume("literalString")
    sRes = Mid(sRes, 2, Len(sRes) - 2)
    sRes = Replace(sRes, """""", """")
    Call addOperation(oPush, , sRes, 1)
End Sub

Private Function parseScopeAccess() As Boolean
    If peek("lBracket", 2) Then
        parseScopeAccess = parseFunctionAccess()
    Else
        parseScopeAccess = parseVariableAccess()
    End If
End Function

Private Function parseVariableAccess() As Boolean
    parseVariableAccess = False
    Dim varName As String: varName = consume("var")
    Dim offset As Long: offset = findVariable(varName)
    If offset >= 0 Then
        parseVariableAccess = True
        Call addOperation(oAccess, , 1 + offset, 1)
    Else
        iTokenIndex = iTokenIndex - 1 ' Revert token consumption
    End If
End Function

Private Sub parseAssignment()
    Dim varName As String: varName = consume("var")
    Call consume("equal")
    Call parseExpression
    Dim offset As Long: offset = findVariable(varName)
    If offset >= 0 Then
        ' If the variable already existed, move the data to that pos on the stack
        Call addOperation(oSet, , offset, -1)
        Call addOperation(oAccess, , offset, 1) ' To keep a return value
    Else
        ' If the variable didn't exist yet, treat this stack pos as its source
        Call scopes(scopeCount).Add(varName, stackSize)
    End If
End Sub

Private Function findVariable(varName As String) As Long
    Dim scope As Long: scope = scopeCount
    findVariable = -1
    While scope > 0
        If scopes(scope).exists(varName) Then
            If scope < funcScope Then
                Throw ("Can't access """ & varName & """, functions can unfortunately not access data outside their block")
            ElseIf scopesArgCount(scope).exists(varName) Then
                Throw ("Expected a variable, but found a function for name " & varName)
            Else
                findVariable = stackSize - scopes(scope).item(varName)
                scope = 0
            End If
        End If
        scope = scope - 1
    Wend
End Function
'todo:
Private Function parseFunctionAccess() As Boolean
    parseFunctionAccess = False
    Dim funcName As String: funcName = consume("var")
    Dim argCount As Long
    Dim funcPos As Long: funcPos = findFunction(funcName, argCount)
    If funcPos <> -1 Then
        parseFunctionAccess = True
        Dim returnPosIndex As Integer: returnPosIndex = addOperation(oPush, , , 1)
        
        ' Consume the arguments
        consume ("lBracket")
        Dim iArgCount As Integer
        While Not peek("rBracket")
            If iArgCount > 0 Then Call consume("comma")
            Call parseExpression
            iArgCount = iArgCount + 1
        Wend
        Call consume("rBracket")
        If iArgCount <> argCount Then
            Throw (argCount & " arguments should have been provided to " & funcName & " but only " & iArgCount & " were received")
        End If
        
        ' Add call and return data
        Call addOperation(oJump, , funcPos, -iArgCount) 'only -argCount since pushing Result and popping return pos cancel out
        operations(returnPosIndex).value = iOperationIndex
    Else
        iTokenIndex = iTokenIndex - 1 ' Revert token consumption
    End If
End Function
'todo:
Private Sub parseFunctionDeclaration()
    ' Create a dedicated scope for this funcion
    Call addScope
    Dim prevFuncScope As Long: prevFuncScope = funcScope
    funcScope = scopeCount
    
    ' Add operation to skip this code in normal operation flow
    Dim skipToIndex As Integer: skipToIndex = addOperation(oJump)
    
    ' Obtain the signature
    Call consume("fun")
    Dim funcName As String: funcName = consume("var")
    Call consume("lBracket")
    Dim iArgCount As Integer
    While Not peek("rBracket")
        If iArgCount > 0 Then Call consume("comma")
        Call parseParameterDeclaration
        iArgCount = iArgCount + 1
    Wend
    Call consume("rBracket")
    
    ' Register the function
    Call scopes(scopeCount - 1).Add(funcName, iOperationIndex)
    Call scopesArgCount(scopeCount - 1).Add(funcName, iArgCount)
    
    ' Obtain the body
    Call parseBlock("end")
    Call consume("end")
    While iArgCount > 0
        Call addOperation(oMerge, , , -1)
        iArgCount = iArgCount - 1
    Wend
    Call addOperation(oReturn, withValue, , -1)
    operations(skipToIndex).value = iOperationIndex
    
    ' Reset the scope
    scopeCount = scopeCount - 1
    funcScope = prevFuncScope
End Sub
'todo:
Private Sub parseParameterDeclaration()
    Dim varName As String: varName = consume("var")
    Dim offset As Long: offset = findVariable(varName)
    If offset >= 0 Then
        Throw ("You can't declare multiple parameters with the same name")
    Else
        ' Reserve a spot for this parameter, it will be pushed by the caller
        stackSize = stackSize + 1
        Call scopes(scopeCount).Add(varName, stackSize)
    End If
End Sub

Private Function findFunction(varName As String, Optional ByRef argCount As Long) As Long
    Dim scope As Long: scope = scopeCount
    findFunction = -1
    While scope > 0
        If scopes(scope).exists(varName) Then
            If Not scopesArgCount(scope).exists(varName) Then
                Throw ("Expected a function, but found a variable for name " & varName)
            Else
                findFunction = scopes(scope).item(varName)
                argCount = scopesArgCount(scope).item(varName)
                scope = 0
            End If
        End If
        scope = scope - 1
    Wend
End Function


'==============
'
'  EVALUATION
'
'==============

'Evaluates the given list of operations
'@param {Operation()} operations  The operations to evaluate
'@returns {Variant} The result of the operations
Private Function evaluate(ByRef ops() As Operation) As Variant
    Dim stack() As Variant
    ReDim stack(0 To 5)
    Dim stackPtr As Long: stackPtr = 0
    
    Dim op As Operation
    Dim v1 As Variant
    Dim v2 As Variant
    Dim v3 As Variant
    Dim opIndex As Long: opIndex = 0
    Dim opCount As Long: opCount = UBound(ops)
    
    While opIndex <= opCount
        op = ops(opIndex)
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
                    Case ISubType.oMod
                        v1 = popV(stack, stackPtr)
                        v3 = v1 Mod v2
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
            'Func
            Case iType.oFunc
                Dim args() As Variant
                args = getArgs(stack, stackPtr)
                v1 = popV(stack, stackPtr)
                v3 = evaluateFunc(v1, args)
                Call pushV(stack, stackPtr, v3)
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
                Select Case op.subType
                    Case ISubType.argument
                        Dim iArgIndex As Long: iArgIndex = val(Mid(op.value, 2)) + LBound(vLastArgs) - 1
                        If iArgIndex <= UBound(vLastArgs) Then
                            Call pushV(stack, stackPtr, vLastArgs(iArgIndex))
                        Else
                            Call Throw("Argument " & iArgIndex & " not supplied to Lambda.")
                        End If
                    Case Else
                        Call pushV(stack, stackPtr, stack(stackPtr - op.value))
                End Select
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
    Call CopyVariant(evaluate, stack(0))
End Function

'Retrieves the arguments from the stack
'@param {ByRef Variant()} stack     The stack to get the data from and add the result to
'@param {ByRef Long} stackPtr       The pointer that indicates the position of the top of the stack
'@returns {Variant()} The args list
Private Function getArgs(ByRef stack() As Variant, ByRef stackPtr As Long) As Variant
    Dim argCount As Variant: argCount = stack(stackPtr - 1)
    Dim args() As Variant
    If VarType(argCount) = vbString Then
        'If no argument count is specified, there are no arguments
        argCount = 0
        args = Array()
    Else
        'If an argument count is provided, extract all arguments into an array
        Call popV(stack, stackPtr)
        ReDim args(1 To argCount)
        For i = 1 To argCount
            Call CopyVariant(args(i), popV(stack, stackPtr))
        Next
    End If
    
    getArgs = args
End Function

'Calls an object method/setter/getter/letter
'@param {ByRef Variant()} stack     The stack to get the data from and add the result to
'@param {ByRef Long} stackPtr       The pointer that indicates the position of the top of the stack
'@param {ByRef Operation} op        The operation to execute
Private Sub objectCaller(ByRef stack() As Variant, ByRef stackPtr As Long, ByRef op As Operation)
    'Get the name and arguments
    Dim args() As Variant: args = getArgs(stack, stackPtr)
    Dim funcName As Variant: funcName = popV(stack, stackPtr)
    
    'Get caller type
    Dim callerType As VbCallType
    Select Case op.subType
        Case ISubType.oPropGet:     callerType = VbGet
        Case ISubType.oMethodCall:  callerType = VbMethod
        Case ISubType.oPropLet:     callerType = VbLet
        Case ISubType.oPropSet:     callerType = VbSet
    End Select
                
    'Call rtcCallByName
    Dim obj As Object
    Set obj = popV(stack, stackPtr)
    Call pushV(stack, stackPtr, rtcCallByName(obj, StrPtr(funcName), callerType, args, &H409))
End Sub

'Evaluates the built in standard functions
'@param {String} sFuncName The name of the function to invoke
'@param {Variant} args() The arguments
'@returns The result
Private Function evaluateFunc(ByVal sFuncName As String, ByVal args As Variant) As Variant
    Dim iArgStart As Long: iArgStart = LBound(args)
    If TypeName(stdLambda.oFunctExt) = "Dictionary" Then
        If Me.oFunctExt.exists(sFuncName) Then
            evaluateFunc = Me.oFunctExt(sFuncName).RunEx(args)
            Exit Function
        End If
    End If
    
    Select Case LCase(sFuncName)
        Case "eval": evaluateFunc = stdLambda.Create(args(iArgStart)).Run()

        'Useful OOP constants
        Case "thisworkbook": Set evaluateFunc = ThisWorkbook
        Case "application":  Set evaluateFunc = Application

        'MATH:
        '-----
        Case "abs": evaluateFunc = VBA.Math.Abs(args(iArgStart))
        Case "int": evaluateFunc = VBA.Int(args(iArgStart))
        Case "fix": evaluateFunc = VBA.Fix(args(iArgStart))
        Case "exp": evaluateFunc = VBA.Math.Exp(args(iArgStart))
        Case "log": evaluateFunc = VBA.Math.Log(args(iArgStart))
        Case "sqr": evaluateFunc = VBA.Math.Sqr(args(iArgStart))
        Case "sgn": evaluateFunc = VBA.Math.Sgn(args(iArgStart))
        Case "rnd": evaluateFunc = VBA.Math.Rnd(args(iArgStart))

        'Trigonometry
        Case "cos":  evaluateFunc = VBA.Math.Cos(args(iArgStart))
        Case "sin":  evaluateFunc = VBA.Math.Sin(args(iArgStart))
        Case "tan":  evaluateFunc = VBA.Math.Tan(args(iArgStart))
        Case "atn":  evaluateFunc = VBA.Math.Atn(args(iArgStart))
        Case "asin": evaluateFunc = VBA.Math.Atn(args(iArgStart) / VBA.Math.Sqr(-1 * args(iArgStart) * args(iArgStart) + 1))
        Case "acos": evaluateFunc = VBA.Math.Atn(-1 * args(iArgStart) / VBA.Math.Sqr(-1 * args(iArgStart) * args(iArgStart) + 1)) + 2 * Atn(1)

        'VBA Constants:
        Case "vbcrlf":          evaluateFunc = vbCrLf
        Case "vbcr":            evaluateFunc = vbCr
        Case "vblf":            evaluateFunc = vbLf
        Case "vbnewline":       evaluateFunc = vbNewLine
        Case "vbnullchar":      evaluateFunc = vbNullChar
        Case "vbnullstring":    evaluateFunc = vbNullString
        Case "vbobjecterror":   evaluateFunc = vbObjectError
        Case "vbtab":           evaluateFunc = vbTab
        Case "vbback":          evaluateFunc = vbBack
        Case "vbformfeed":      evaluateFunc = vbFormFeed
        Case "vbverticaltab":   evaluateFunc = vbVerticalTab


        'VBA Structure
        Case "array": evaluateFunc = args
        'TODO: Case "CallByName": evaluateFunc = CallByName(args(iArgStart))
        Case "createobject"
            Select Case UBound(args)
                Case iArgStart
                    Set evaluateFunc = CreateObject(args(iArgStart))
                Case iArgStart + 1
                    Set evaluateFunc = CreateObject(args(iArgStart), args(iArgStart + 1))
            End Select
        Case "getobject"
            Select Case UBound(args)
                Case iArgStart
                    Set evaluateFunc = GetObject(args(iArgStart))
                Case iArgStart + 1
                    Set evaluateFunc = GetObject(args(iArgStart), args(iArgStart + 1))
            End Select
        Case "iff"
            If CBool(args(iArgStart)) Then
                evaluateFunc = args(iArgStart + 1)
            Else
                evaluateFunc = args(iArgStart + 2)
            End If

        'VBA Casting
        Case "cbool": evaluateFunc = VBA.Conversion.CBool(args(iArgStart))
        Case "cbyte": evaluateFunc = VBA.Conversion.CByte(args(iArgStart))
        Case "ccur":  evaluateFunc = VBA.Conversion.CCur(args(iArgStart))
        Case "cdate": evaluateFunc = VBA.Conversion.CDate(args(iArgStart))
        Case "csng":  evaluateFunc = VBA.Conversion.CSng(args(iArgStart))
        Case "cdbl":  evaluateFunc = VBA.Conversion.CDbl(args(iArgStart))
        Case "cint":  evaluateFunc = VBA.Conversion.CInt(args(iArgStart))
        Case "clng":  evaluateFunc = VBA.Conversion.CLng(args(iArgStart))
        Case "cstr":  evaluateFunc = VBA.Conversion.CStr(args(iArgStart))
        Case "cvar":  evaluateFunc = VBA.Conversion.CVar(args(iArgStart))
        Case "cverr": evaluateFunc = VBA.Conversion.CVErr(args(iArgStart))
        
        'Conversion
        Case "asc":     evaluateFunc = VBA.Asc(args(iArgStart))
        Case "chr":     evaluateFunc = VBA.Chr(args(iArgStart))
        
        Case "format"
            Select Case UBound(args)
                Case iArgStart
                    evaluateFunc = Format(args(iArgStart))
                Case iArgStart + 1
                    evaluateFunc = Format(args(iArgStart), args(iArgStart + 1))
                Case iArgStart + 2
                    evaluateFunc = Format(args(iArgStart), args(iArgStart + 1), args(iArgStart + 2))
                Case iArgStart + 3
                    evaluateFunc = Format(args(iArgStart), args(iArgStart + 1), args(iArgStart + 2), args(iArgStart + 3))
            End Select
        Case "hex":     evaluateFunc = VBA.Conversion.Hex(args(iArgStart))
        Case "oct":     evaluateFunc = VBA.Conversion.Oct(args(iArgStart))
        Case "str":     evaluateFunc = VBA.Conversion.Str(args(iArgStart))
        Case "val":     evaluateFunc = VBA.Conversion.val(args(iArgStart))
        
        'String functions
        Case "trim":  evaluateFunc = VBA.Trim(args(iArgStart))
        Case "lcase": evaluateFunc = VBA.LCase(args(iArgStart))
        Case "ucase": evaluateFunc = VBA.UCase(args(iArgStart))
        Case "right": evaluateFunc = VBA.Right(args(iArgStart), args(iArgStart + 1))
        Case "left":  evaluateFunc = VBA.Left(args(iArgStart), args(iArgStart + 1))
        Case "len":   evaluateFunc = VBA.Len(args(iArgStart))

        Case "mid"
            Select Case UBound(args)
                Case iArgStart + 1
                    evaluateFunc = VBA.Mid(args(iArgStart), args(iArgStart + 1))
                Case iArgStart + 2
                    evaluateFunc = VBA.Mid(args(iArgStart), args(iArgStart + 1), args(iArgStart + 2))
            End Select
        'Misc
        Case "now": evaluateFunc = VBA.DateTime.Now()

        Case Else
            Throw ("No such function: " & sFuncName)
    End Select
End Function

'==========================
'
'  General helper methods
'
'==========================

'------------
'tokenisation
'------------

'Tokenise the input string
'@param {string} sInput String to tokenise
'@return {token[]} A list of Token structs
Private Function Tokenise(ByVal sInput As String) As token()
    Dim defs() As TokenDefinition
    defs = getTokenDefinitions()
    
    Dim tokens() As token, iTokenDef As Long
    ReDim tokens(1 To 1)
    
    Dim sInputOld As String
    sInputOld = sInput
    
    Dim iNumTokens As Long
    iNumTokens = 0
    While Len(sInput) > 0
        Dim bMatched As Boolean
        bMatched = False
        
        For iTokenDef = 1 To UBound(defs)
            'Test match, if matched then add token
            If defs(iTokenDef).RegexObj.test(sInput) Then
                'Get match details
                Dim oMatch As Object: Set oMatch = defs(iTokenDef).RegexObj.execute(sInput)
                
                'Create new token
                iNumTokens = iNumTokens + 1
                ReDim Preserve tokens(1 To iNumTokens)
                
                'Tokenise
                tokens(iNumTokens).Type = defs(iTokenDef)
                tokens(iNumTokens).value = oMatch(0)
                
                'Trim string to unmatched range
                sInput = Mid(sInput, Len(oMatch(0)) + 1)
                
                'Flag that a match was made
                bMatched = True
                Exit For
            End If
        Next
        
        'If no match made then syntax error
        If Not bMatched Then
            Call Throw("Syntax Error unexpected character """ & Mid(sInput, 1, 1) & """")
        End If
    Wend
    
    'Add eof token
    ReDim Preserve tokens(1 To iNumTokens + 1)
    tokens(iNumTokens + 1).Type.name = "eof"
    
    Tokenise = removeTokens(tokens, "space")
End Function

Private Function getTokenDefinition(ByVal sName As String, ByVal sRegex As String, Optional ByVal ignoreCase As Boolean = True) As TokenDefinition
    getTokenDefinition.name = sName
    getTokenDefinition.Regex = sRegex
    Set getTokenDefinition.RegexObj = CreateObject("VBScript.Regexp")
    getTokenDefinition.RegexObj.Pattern = "^(?:" & sRegex & ")"
    getTokenDefinition.RegexObj.ignoreCase = ignoreCase
End Function

'Copies one variant to a destination
'@param {ByRef Token()} tokens Tokens to remove the specified type from
'@param {string} sRemoveType   Token type to remove.
'@returns {Token()} The modified token array.
Private Function removeTokens(ByRef tokens() As token, ByVal sRemoveType As String) As token()
    Dim iCountRemoved As Long: iCountRemoved = 0
    Dim iToken As Long
    For iToken = LBound(tokens) To UBound(tokens)
        If tokens(iToken).Type.name <> sRemoveType Then
            tokens(iToken - iCountRemoved) = tokens(iToken)
        Else
            iCountRemoved = iCountRemoved + 1
        End If
    Next
    ReDim Preserve tokens(LBound(tokens) To (UBound(tokens) - iCountRemoved))
    removeTokens = tokens
End Function

'-------
'parsing
'-------

'Shifts the Tokens array (uses an index)
'@returns {token} The token at the tokenIndex
Private Function ShiftTokens() As token
    If iTokenIndex = 0 Then iTokenIndex = 1
    
    'Get next token
    ShiftTokens = tokens(iTokenIndex)
    
    'Increment token index
    iTokenIndex = iTokenIndex + 1
End Function

' Consumes a token
' @param {string} token The token type name to consume
' @throws If the expected token wasn't found
' @returns {string} The value of the token
Private Function consume(ByVal sType As String) As String
    Dim firstToken As token
    firstToken = ShiftTokens()
    If firstToken.Type.name <> sType Then
        Call Throw("Unexpected token, found: " & firstToken.Type.name & " but expected: " & sType)
    Else
        consume = firstToken.value
    End If
End Function

'Checks whether the token at iTokenIndex is of the given type
'@param {string} token  The token that is expected
'@param {long} offset   The number of tokens to look into the future, defaults to 1
'@returns {boolean} Whether the expected token was found
Private Function peek(ByVal sTokenType As String, Optional offset As Long = 1) As Boolean
    If iTokenIndex = 0 Then iTokenIndex = 1
    If iTokenIndex + offset - 1 <= UBound(tokens) Then
        peek = tokens(iTokenIndex + offset - 1).Type.name = sTokenType
    Else
        peek = False
    End If
End Function

' Combines peek and consume, consuming a token only if matched, without throwing an error if not
' @param {string} token The token that is expected
' @returns {vbNullString|string} Whether the expected token was found
Private Function optConsume(ByVal sTokenType As String) As Boolean
    Dim matched As Boolean: matched = peek(sTokenType)
    If matched Then
        Call consume(sTokenType)
    End If
    optConsume = matched
End Function

'Checks the value of the passed parameter, to check if it is the unique constant
'@param {Variant} test The value to test. May be an object or literal value
'@returns {Boolean} True if the value is the unique constant, otherwise false
Private Function isUniqueConst(ByRef test As Variant) As Boolean
    If Not IsObject(test) Then
        If VarType(test) = vbString Then
            If test = UniqueConst Then
                isUniqueConst = True
                Exit Function
            End If
        End If
    End If
    isUniqueConst = False
End Function
 
'Adds an operation to the instance operations list
'@param {IType} kType         The main type of the operation
'@param {ISubType} subType    The sub type of the operation
'@param {Variant} value       The value associated with the operation
'@param {Integer} stackDelta  The effect this has on the stack size (increasing or decreasing it)
'@returns {Integer} The index of the created operation
Private Function addOperation(kType As iType, Optional subType As ISubType, Optional value As Variant, Optional stackDelta As Integer) As Integer
    If iOperationIndex = 0 Then
        ReDim Preserve operations(0 To 1)
    Else
        Dim size As Long: size = UBound(operations)
        If iOperationIndex > size Then
            ReDim Preserve operations(0 To size * 2)
        End If
    End If
    
    With operations(iOperationIndex)
        .Type = kType
        .subType = subType
        Call CopyVariant(.value, value)
    End With
    addOperation = iOperationIndex
    stackSize = stackSize + stackDelta
    
    iOperationIndex = iOperationIndex + 1
End Function

'Resizes the operations list so there are no more empty operations
Private Sub finishOperations()
    ReDim Preserve operations(0 To iOperationIndex)
End Sub

'----------
'evaluation
'----------

Private Sub pushV(ByRef stack() As Variant, ByRef index As Long, ByVal item As Variant)
    Dim size As Long: size = UBound(stack)
    If index > size Then
        ReDim Preserve stack(0 To size * 2)
    End If
    If IsObject(item) Then
        Set stack(index) = item
    Else
        stack(index) = item
    End If
    index = index + 1
End Sub

Private Function popV(ByRef stack() As Variant, ByRef index As Variant) As Variant
    Dim size As Long: size = UBound(stack)
    If index < size / 3 And index > minStackSize Then
        ReDim Preserve stack(0 To CLng(size / 2))
    End If
    index = index - 1
    If IsObject(stack(index)) Then
        Set popV = stack(index)
    Else
        popV = stack(index)
    End If
    #If devMode Then
        stack(index) = Empty
    #End If
End Function

'-------
'general
'-------

'Copies one variant to a destination
'@param {ByRef Variant} dest Destination to copy variant to
'@param {Variant} value Source to copy variant from.
'@perf This appears to be a faster variant of "oleaut32.dll\VariantCopy" + it's multi-platform
Private Sub CopyVariant(ByRef dest As Variant, ByVal value As Variant)
  If IsObject(value) Then
    Set dest = value
  Else
    dest = value
  End If
End Sub


'Throws an error
'@param {string} The error message to be thrown
'@returns {void}
Private Sub Throw(ByVal sMessage As String)
    MsgBox sMessage, vbCritical
    End
End Sub




