VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdLambda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Direct call convention of VBA.CallByName
#If VBA6 Then
  Private Declare PtrSafe Function rtcCallByName Lib "msvbvm60" (ByRef vRet As Variant, ByVal cObj As Object, ByVal sMethod As LongPtr, ByVal eCallType As VbCallType, ByRef pArgs() As Variant, ByVal lcid As Long) As Long
  Private Declare PtrSafe Sub VariantCopy Lib "oleaut32.dll" (ByRef pvargDest As Variant, ByRef pvargSrc As Variant)
#Else
  Private Declare Function rtcCallByName Lib "msvbvm60" (ByRef vRet As Variant, ByVal cObj As Object, ByVal sMethod As LongPtr, ByVal eCallType As VbCallType, ByRef pArgs() As Variant, ByVal lcid As Long) As Long
  Private Declare Sub VariantCopy Lib "oleaut32.dll" (ByRef pvargDest As Variant, ByRef pvargSrc As Variant)
#End If

'Tokens and definitions
Private Type TokenDefinition
    Name As String
    Regex As String
    RegexObj As Object
End Type
Private Type token
    Type As TokenDefinition
    value As Variant
    BracketDepth As Long
End Type

'Special constant used in parsing:
Const UniqueConst As String = "3207af79-30df-4890-ade1-640f9f28f309"

Private tokens() As token
Private iTokenIndex As Long
Private vLastArgs As Variant





''Usage:
'Debug.Print stdLambda.Create("1+3*8/2*(2+2+3)").Execute()
'With stdLambda.Create("$1+1+3*8/2*(2+2+3)")
'    Debug.Print .Execute(10)
'    Debug.Print .Execute(15)
'    Debug.Print .Execute(20)
'End With
'Debug.Print stdLambda.Create("$1.Range(""A1"")").Execute(Sheets(1)).Address(True, True, xlA1, True)
'Debug.Print stdLambda.Create("$1#join("","")").Execute(stdArray.Create(1,2))
Public Function Create(ByVal sEquation As String) As stdLambda
    Set Create = New stdLambda
    Call Create.Init(sEquation)
End Function
Public Sub Init(ByVal sEquation As String)
    tokens = Tokenise(sEquation)
End Sub

Function Execute(ParamArray args() As Variant) As Variant
    'Convert arguments to array (vLastArgs)
    Dim i As Long
    If UBound(args) - LBound(args) + 1 = 0 Then
        vLastArgs = Array()
    Else
        ReDim vLastArgs(1 To (UBound(args) - LBound(args) + 1))
        For i = LBound(args) To UBound(args)
            Call CopyVariant(vLastArgs(i + 1), args(i))
        Next
    End If
    
    'Ensure iTokenIndex = 1 for evaluation
    iTokenIndex = 1
    
    'Execute top-down parser
    Call CopyVariant(Execute, expression())
End Function










Public Sub test()
    'Evaluate expresssion
    Debug.Print stdLambda.Create("1+3*8/2*(2+2+3)").Execute()

    'Evaluate a expression with variables
    With stdLambda.Create("$1+1+3*8/2*(2+2+3)")
        Debug.Print .Execute(10)
        Debug.Print .Execute(15)
        Debug.Print .Execute(20)
    End With
    
    'Evaluate property access
    Debug.Print stdLambda.Create("$1.Range(""A1"")").Execute(Sheets(1)).Address(True, True, xlA1, True)
    
    'Evaluate method access
    Range("A1").value = 1
    Range("A2").value = 2
    Range("A3").value = 3
    Range("A4").value = 4
    Debug.Print stdLambda.Create("$1#Find(4)").Execute(Range("A:A")).Address(True, True, xlA1, True)

    'TODO:
    'Evaluate pure function
    'Debug.Print stdLambda.Create("cos($1)").Execute(PI)
    Debug.Assert False
End Sub






'Evaluate an expression
Function expression() As Variant
    Dim res As Variant: Call CopyVariant(res, term())
    Dim bLoop As Boolean: bLoop = True
    Do
        If optConsume("add") Then
            res = res + term()
        ElseIf optConsume("sub") Then
            res = res - term()
        Else
            bLoop = False
        End If
    Loop While bLoop
    Call CopyVariant(expression, res)
End Function

Function term() As Variant
    Dim res As Variant: Call CopyVariant(res, factor())
    Dim bLoop As Boolean: bLoop = True
    Do
        If optConsume("mul") Then
            res = res * factor()
        ElseIf optConsume("div") Then
            res = res / factor()
        Else
            bLoop = False
        End If
    Loop While bLoop
    Call CopyVariant(term, res)
End Function

Function factor() As Variant
    Dim res As Variant
    
    'Deal with not unary operator
    Dim invert As Variant: invert = vbNull
    While optConsume("not")
        If invert = vbNull Then invert = False
        invert = Not invert
    Wend

    'Deal with numbers, vars, strings, booleans and brackets
    If peek("literalNumber") Then
        res = CDbl(consume("literalNumber"))
    ElseIf peek("var") Then
        Call CopyVariant(res, EvaluateVarName(consume("var")))
    ElseIf peek("literalString") Then
        res = consumeString()
    ElseIf peek("literalBoolean") Then
        res = consume("literalBoolean") = "true"
    Else
        Call consume("lBracket")
        res = expression()
        Call consume("rBracket")
    End If
    
    'Execute invert
    If invert <> vbNull Then
        If invert = True Then
            res = Not CBool(res)
        Else
            res = CBool(res)
        End If
    End If
    
    Call CopyVariant(factor, manyAccessors(res))
End Function

Function manyAccessors(value As Variant) As Variant
    'Copy value into res
    Dim res As Variant
    Call CopyVariant(res, value)

    Dim bLoop As Boolean: bLoop = True
    Do
        Dim newRes As Variant: newRes = UniqueConst
        If isUniqueConst(newRes) Then Call CopyVariant(newRes, optObjectProperty(res))
        If isUniqueConst(newRes) Then Call CopyVariant(newRes, optObjectMethod(res))

        If Not isUniqueConst(newRes) Then
            Call CopyVariant(res, newRes)
        Else
            bLoop = False
        End If
    Loop While bLoop

    Call CopyVariant(manyAccessors, res)
End Function

Function optObjectProperty(ByRef value As Variant)
    If IsObject(value) Then
        If optConsume("propertyAccess") Then
            Dim sFuncName As String: sFuncName = consume("funcName")
            Dim args() As Variant: args = optConsumeParameters()
            
            'Call rtcCallByName
            Dim hr As Long, res As Variant
            hr = rtcCallByName(res, value, StrPtr(sFuncName), VbCallType.VbGet, args, &H409)
            
            'If error then raise
            If hr < 0 Then
                Call Throw("Error in calling " & sFuncName & " property of " & TypeName(value) & " object.")
            Else
                Call CopyVariant(optObjectProperty, res)
            End If
            Exit Function
        End If
    End If
    
    optObjectProperty = UniqueConst
End Function

Function optObjectMethod(ByRef value As Variant) As Variant
    If IsObject(value) Then
        If optConsume("methodAccess") Then
            Dim sFuncName As String: sFuncName = consume("funcName")
            Dim args() As Variant: args = optConsumeParameters()
            
            'Call rtcCallByName
            Dim hr As Long, res As Variant
            hr = rtcCallByName(res, value, StrPtr(sFuncName), VbCallType.VbMethod, args, &H409)
            
            'If error then raise
            If hr < 0 Then
                Call Throw("Error in calling " & sFuncName & " method of " & TypeName(value) & " object.")
            Else
                Call CopyVariant(optObjectMethod, res)
            End If
            Exit Function
        End If
    End If
    
    optObjectMethod = UniqueConst
End Function

Function optConsumeParameters() As Variant
    If optConsume("lBracket") Then
        Dim args() As Variant
        args = Array()

        While Not peek("rBracket")
            If UBound(args) - LBound(args) + 1 > 0 Then Call consume("comma")
            ReDim Preserve args(0 To UBound(args) + 1)
            args(UBound(args)) = expression()
        Wend

        Call consume("rBracket")
        optConsumeParameters = args
    Else
        optConsumeParameters = Array()
    End If
End Function

Function consumeString() As String
    Dim sRes As String: sRes = consume("literalString")
    sRes = Mid(sRes, 2, Len(sRes) - 2)
    sRes = Replace(sRes, """""", """")
    consumeString = sRes
End Function


'Evaluates Variable name
Private Function EvaluateVarName(ByVal sValueName As String) As Variant
    Dim iArgIndex As Long: iArgIndex = Val(Mid(sValueName, 2))
    
    If iArgIndex <= UBound(vLastArgs) Then
        'Evaluate varname, allow for object values...
        Call CopyVariant(EvaluateVarName, vLastArgs(iArgIndex))
    Else
        Call Throw("Argument " & iArgIndex & " not supplied to Lambda.")
    End If
End Function


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
                Dim oMatch As Object: Set oMatch = defs(iTokenDef).RegexObj.Execute(sInput)
                
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
            Call Throw("Syntax Error - Lexer Error")
        End If
    Wend
    
    Tokenise = tokens
End Function

'Tokeniser helpers
Private Function getTokenDefinitions() As TokenDefinition()
    Dim arr() As TokenDefinition
    ReDim arr(1 To 99)
    
    Dim i As Long: i = 0
    'Literal
    i = i + 1: arr(i) = getTokenDefinition("literalString", """(?:""""|[^""])*""") 'String
    i = i + 1: arr(i) = getTokenDefinition("literalNumber", "\d+(?:\.\d+)?")   'Number
    i = i + 1: arr(i) = getTokenDefinition("literalBoolean", "True|False")
    
    'Structural
    i = i + 1: arr(i) = getTokenDefinition("lBracket", "\(")
    i = i + 1: arr(i) = getTokenDefinition("rBracket", "\)")
    i = i + 1: arr(i) = getTokenDefinition("comma", ",")
    i = i + 1: arr(i) = getTokenDefinition("funcName", "[a-zA-Z][a-zA-Z0-9_]+")
    
    'VarName
    i = i + 1: arr(i) = getTokenDefinition("var", "\$\d+")
    
    'Operators
    i = i + 1: arr(i) = getTokenDefinition("propertyAccess", "\.")
    i = i + 1: arr(i) = getTokenDefinition("methodAccess", "\#")
    i = i + 1: arr(i) = getTokenDefinition("mul", "\*")
    i = i + 1: arr(i) = getTokenDefinition("div", "\/")
    i = i + 1: arr(i) = getTokenDefinition("add", "\+")
    i = i + 1: arr(i) = getTokenDefinition("sub", "\-")
    i = i + 1: arr(i) = getTokenDefinition("eq", "\=")
    i = i + 1: arr(i) = getTokenDefinition("neq", "\<\>")
    i = i + 1: arr(i) = getTokenDefinition("gt", "\>")
    i = i + 1: arr(i) = getTokenDefinition("gte", "\>\=")
    i = i + 1: arr(i) = getTokenDefinition("lt", "\<")
    i = i + 1: arr(i) = getTokenDefinition("lte", "\<\=")
    i = i + 1: arr(i) = getTokenDefinition("concat", "\&")
    
    ReDim Preserve arr(1 To i)

    getTokenDefinitions = arr
End Function
Private Function getTokenDefinition(ByVal sName As String, ByVal sRegex As String, Optional ByVal ignoreCase As Boolean = True) As TokenDefinition
    getTokenDefinition.Name = sName
    getTokenDefinition.Regex = sRegex
    Set getTokenDefinition.RegexObj = CreateObject("VBScript.Regexp")
    getTokenDefinition.RegexObj.Pattern = "^(?:" & sRegex & ")"
    getTokenDefinition.RegexObj.ignoreCase = ignoreCase
End Function


'==============================================================================================================================
'
'Helper Methods:
'
'==============================================================================================================================

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

'Shifts the Tokens array (uses an index)
'@returns {token} The token at the tokenIndex
Private Function ShiftTokens() As token
    If iTokenIndex = 0 Then iTokenIndex = 1
    
    'Get next token
    ShiftTokens = tokens(iTokenIndex)
    
    'Increment token index
    iTokenIndex = iTokenIndex + 1
End Function

'Throws an error
'@param {string} The error message to be thrown
'@returns {void}
Private Sub Throw(ByVal sMessage As String)
    MsgBox sMessage, vbCritical
    End
End Sub


' Consumes a token
' @param {string} token The token type name to consume
' @throws If the expected token wasn't found
' @returns {string} The value of the token
Private Function consume(ByVal sType As String) As String
    Dim firstToken As token
    firstToken = ShiftTokens()
    If firstToken.Type.Name <> sType Then
        Call Throw("Unexpected token, found: " & firstToken.Type.Name & " but expected: " & sType)
    Else
        consume = firstToken.value
    End If
End Function

'Checks whether the token at iTokenIndex is of the given type
'@param {string} token The token that is expected
'@returns {boolean} Whether the expected token was found
Private Function peek(ByVal sTokenType As String) As Boolean
    If iTokenIndex = 0 Then iTokenIndex = 1
    If iTokenIndex <= UBound(tokens) Then
        peek = tokens(iTokenIndex).Type.Name = sTokenType
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
        consume(sTokenType)
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


'==============================================================================================================================
'
'Old Deprecated functions for reference:
'
'==============================================================================================================================
Private Function zzEvaluateBinaryOperator(ByRef tokens() As token, ByRef args As Variant, ByVal iToken As Long)
    Dim result As Variant
    Select Case tokens(iToken).Type.Name
        Case "add"
            result = tokens(iToken - 1).value + tokens(iToken + 1).value
        Case "sub"
            result = tokens(iToken - 1).value - tokens(iToken + 1).value
        Case "mul"
            result = tokens(iToken - 1).value * tokens(iToken + 1).value
        Case "div"
            result = tokens(iToken - 1).value / tokens(iToken + 1).value
        Case "BooleanOp"
            Select Case tokens(iToken).value
                Case "="
                    result = tokens(iToken - 1).value = tokens(iToken + 1).value
                Case ">"
                    result = tokens(iToken - 1).value > tokens(iToken + 1).value
                Case ">="
                    result = tokens(iToken - 1).value >= tokens(iToken + 1).value
                Case "<"
                    result = tokens(iToken - 1).value < tokens(iToken + 1).value
                Case "<="
                    result = tokens(iToken - 1).value <= tokens(iToken + 1).value
                Case "<>"
                    result = tokens(iToken - 1).value <> tokens(iToken + 1).value
                Case Else
                    Debug.Print "Unexpected evaluation of Binary Operator """ & tokens(iToken).value & """"
                    End
            End Select
        Case "Concatenate"
            result = tokens(iToken - 1).value & tokens(iToken + 1).value
        Case Else
            Debug.Print "Unexpected evaluation of Binary Operator """ & tokens(iToken).value & """"
            End
    End Select
    
    
    RemoveToken tokens, iToken + 1
    tokens(iToken).Type.Name = "RESULT"
    tokens(iToken).value = result
    RemoveToken tokens, iToken - 1
    
End Function




