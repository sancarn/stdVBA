Private Type TokenDefinition
    Name As String
    Regex As String
    RegexObj As Object
End Type
Private Type token
    Type As TokenDefinition
    Value As Variant
    BracketDepth As Long
End Type

Private tokens() As token
Private iTokenIndex As Long

Sub Test
  Debug.Print EvaluateEx("1+3*8/2*(2+2+3)")
End Sub




'Shifts the Tokens array (uses an index)
'@returns {token} The token at the tokenIndex
Function ShiftTokens() As token
    If iTokenIndex = 0 Then iTokenIndex = 1
    
    'Get next token
    ShiftTokens = tokens(iTokenIndex)
    
    'Increment token index
    iTokenIndex = iTokenIndex + 1
End Function

Sub Throw(ByVal sMessage As String)
    Debug.Print "Unexpected token, found: " & firstToken.Type.Name & " but expected: " & sType
    End
End Sub


' Consumes a token
' @param {string} token The token type name to consume
' @throws If the expected token wasn't found
' @returns {string} The value of the token
Function consume(ByVal sType As String) As String
    Dim firstToken As token
    firstToken = ShiftTokens()
    If firstToken.Type.Name <> sType Then
        Debug.Print "Unexpected token, found: " & firstToken.Type.Name & " but expected: " & sType
        End
    Else
        consume = firstToken.Value
    End If
End Function

'Checks whether the token at iTokenIndex is of the given type
'@param {string} token The token that is expected
'@returns {boolean} Whether the expected token was found
Function peek(ByVal sTokenType As String) As Boolean
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
Function optConsume(ByVal sTokenType As String) As String
    Dim matched As Boolean: matched = peek(sTokenType)
    If matched Then
        optConsume = consume(sTokenType)
    Else
        optConsume = vbNullString
    End If
End Function

'------------------------------------------------------


Function EvaluateEx(ByVal sExpression As String)
    tokens = Tokenise(sExpression)
    iTokenIndex = 1
    EvaluateEx = expression()
End Function

'Evaluate an expression
Function expression() As Variant
    Dim res As Variant: res = term()
    Dim bLoop As Boolean: bLoop = True
    Do
        If optConsume("add") <> vbNullString Then
            res = res + term()
        ElseIf optConsume("sub") <> vbNullString Then
            res = res - term()
        Else
            bLoop = False
        End If
    Loop While bLoop
    expression = res
End Function

Function term() As Variant
    Dim res As Variant: res = factor()
    Dim bLoop As Boolean: bLoop = True
    Do
        If optConsume("mul") <> vbNullString Then
            res = res * factor()
        ElseIf optConsume("div") <> vbNullString Then
            res = res / factor()
        Else
            bLoop = False
        End If
    Loop While bLoop
    term = res
End Function

Function factor() As Variant
    Dim res As Variant
    If peek("literalNumber") Then
        res = CDbl(consume("literalNumber"))
    Else
        Call consume("lBracket")
        res = expression()
        Call consume("rBracket")
    End If
    factor = res
End Function




Function Tokenise(ByVal sInput As String) As token()
    Dim defs() As TokenDefinition
    defs = getTokenDefinitions()
    
    Dim tokens() As token, iTokenDef As Long
    ReDim tokens(1 To 1)
    
    Dim sInputOld As String
    sInputOld = sInput
    
    Dim iBracketDepth As Long
    iBracketDepth = 0
    
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
                tokens(iNumTokens).Value = oMatch(0)
                tokens(iNumTokens).BracketDepth = iBracketDepth
                
                'Trim string to unmatched range
                sInput = Mid(sInput, Len(oMatch(0)) + 1)
                
                'Mark bracket depth as we tokenise
                Select Case defs(iTokenDef).Name
                    Case "LParen"
                        iBracketDepth = iBracketDepth + 1
                    Case "RParen"
                        iBracketDepth = iBracketDepth - 1
                        
                        'Overwrite bracket depth
                        tokens(iNumTokens).BracketDepth = iBracketDepth
                    Case Else
                        'No change to bracket depth
                End Select
                
                'Flag that a match was made
                bMatched = True
                Exit For
            End If
        Next
        
        'If no match made then syntax error
        If Not bMatched Then
            Debug.Print "Syntax Error - Lexer Error"
            End
        End If
    Wend
    
    Tokenise = tokens
End Function

'Tokeniser helpers
Private Function getTokenDefinitions() As TokenDefinition()
    Dim arr(1 To 17) As TokenDefinition
    'Literal
    arr(1) = getTokenDefinition("literalString", """(?:""""|[^""])*""") 'String
    arr(2) = getTokenDefinition("literalNumber", "\d+(?:\.\d+)?")       'Number
    arr(3) = getTokenDefinition("literalBoolean", "True|False")
    
    'Structural
    arr(4) = getTokenDefinition("lBracket", "\(")
    arr(5) = getTokenDefinition("rBracket", "\)")
    arr(6) = getTokenDefinition("zzFuncDelim", ",")
    arr(7) = getTokenDefinition("zzIfStatement", "if")
    arr(8) = getTokenDefinition("zzFuncName", "[a-zA-Z][a-zA-Z0-9_]+")
    
    'VarName
    arr(9) = getTokenDefinition("zzVarName", "\$\d+")
    
    'Operators
    arr(10) = getTokenDefinition("zzPropertyAccess", "\.")
    arr(11) = getTokenDefinition("zzMethodAccess", "\.")
    arr(12) = getTokenDefinition("mul", "\*")
    arr(13) = getTokenDefinition("div", "\/")
    arr(14) = getTokenDefinition("add", "\+")
    arr(15) = getTokenDefinition("sub", "\-")
    arr(16) = getTokenDefinition("zzBooleanOp", "(?:\=|\>\=|\>|\<|\<\=|\<\>)")
    arr(17) = getTokenDefinition("zzConcatenate", "\&")
    
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
'Old Deprecated functions for reference:
'
'==============================================================================================================================
Function zzEvaluateBinaryOperator(ByRef tokens() As token, ByRef args As Variant, ByVal iToken As Long)
    Dim result As Variant
    Select Case tokens(iToken).Type.Name
        Case "add"
            result = tokens(iToken - 1).Value + tokens(iToken + 1).Value
        Case "sub"
            result = tokens(iToken - 1).Value - tokens(iToken + 1).Value
        Case "mul"
            result = tokens(iToken - 1).Value * tokens(iToken + 1).Value
        Case "div"
            result = tokens(iToken - 1).Value / tokens(iToken + 1).Value
        Case "BooleanOp"
            Select Case tokens(iToken).Value
                Case "="
                    result = tokens(iToken - 1).Value = tokens(iToken + 1).Value
                Case ">"
                    result = tokens(iToken - 1).Value > tokens(iToken + 1).Value
                Case ">="
                    result = tokens(iToken - 1).Value >= tokens(iToken + 1).Value
                Case "<"
                    result = tokens(iToken - 1).Value < tokens(iToken + 1).Value
                Case "<="
                    result = tokens(iToken - 1).Value <= tokens(iToken + 1).Value
                Case "<>"
                    result = tokens(iToken - 1).Value <> tokens(iToken + 1).Value
                Case Else
                    Debug.Print "Unexpected evaluation of Binary Operator """ & tokens(iToken).Value & """"
                    End
            End Select
        Case "Concatenate"
            result = tokens(iToken - 1).Value & tokens(iToken + 1).Value
        Case Else
            Debug.Print "Unexpected evaluation of Binary Operator """ & tokens(iToken).Value & """"
            End
    End Select
    
    
    RemoveToken tokens, iToken + 1
    tokens(iToken).Type.Name = "RESULT"
    tokens(iToken).Value = result
    RemoveToken tokens, iToken - 1
    
End Function


Function zzEvaluateLiteral(ByRef tok As token) As token
    Dim tRet As token
    tRet.Type.Name = "RESULT"
    If Left(tok.Value, 1) = """" Then
        tRet.Value = DeSerialize(tok.Value)
    Else
        tRet.Value = CDbl(tok.Value)
    End If
    
    zzEvaluateLiteral = tRet
End Function
Function zzEvaluateVarName(ByRef tok As token, ByRef args As Variant) As token
    Dim tRet As token
    tRet.Type.Name = "RESULT"
    
    Dim iArgIndex As Long: iArgIndex = Val(Mid(tok.Value, 2))
    
    'Evaluate varname, allow for object values...
    If VarType(args(iArgIndex)) = vbObject Then
        Set tRet.Value = args(iArgIndex)
    Else
        tRet.Value = args(iArgIndex)
    End If
    
    zzEvaluateVarName = tRet
End Function
Function zzDeSerialize(ByVal sData As String) As String
    sData = Mid(sData, 2, Len(sData) - 2)
    While InStr(1, sData, """""") > 0
        sData = Replace(sData, """""", """")
    Wend
    zzDeSerialize = sData
End Function
