VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdLambda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements stdICallable

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

Public oFunctExt As Object 'Dictionary<string => stdCallback>
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
Public Sub Init(ByVal sEquation As String, ByVal bSandboxExtras As Boolean)
    If Not bSandboxExtras Then
        Set Me.oFunctExt = stdLambda.oFunctExt
    End If
    tokens = Tokenise(sEquation)
End Sub

Private Function stdICallable_Run(ParamArray params() As Variant) As Variant
    'Bind args
    vLastArgs = params

    'Ensure iTokenIndex = 1 for evaluation
    iTokenIndex = 1

    'Execute top-down parser
    Call CopyVariant(stdICallable_Run, expression())
End Function
Private Function stdICallable_RunEx(params() As Variant) As Variant
    'Bind args
    vLastArgs = params

    'Ensure iTokenIndex = 1 for evaluation
    iTokenIndex = 1

    'Execute top-down parser
    Call CopyVariant(stdICallable_RunEx, expression())
End Function

Function Run(ParamArray params() As Variant) As Variant
    'Bind args
    vLastArgs = params

    'Ensure iTokenIndex = 1 for evaluation
    iTokenIndex = 1
    
    'Execute top-down parser
    Call CopyVariant(Run, expression())
End Function

Function RunEx(params() As Variant) As Variant
    'Bind args
    vLastArgs = params
    
    'Ensure iTokenIndex = 1 for evaluation
    iTokenIndex = 1
    
    'Execute top-down parser
    Call CopyVariant(Run, expression())
End Function










Public Sub zTest()
    'Evaluate expresssion
    Debug.Print stdLambda.Create("1+3*8/2*(2+2+3)").Run()

    'Evaluate a expression with variables
    With stdLambda.Create("$1+1+3*8/2*(2+2+3)")
        Debug.Print .Run(10)
        Debug.Print .Run(15)
        Debug.Print .Run(20)
    End With
    
    'Evaluate property access
    Debug.Print stdLambda.Create("$1.Range(""A1"")").Run(Sheets(1)).Address(True, True, xlA1, True)
    
    'Evaluate method access
    Range("A1").value = 1
    Range("A2").value = 2
    Range("A3").value = 3
    Range("A4").value = 4
    Debug.Print stdLambda.Create("$1#Find(4)").Run(Range("A:A")).Address(True, True, xlA1, True)

    'TODO:
    'Evaluate pure function
    Debug.Print round(stdLambda.Create("cos($1)").Run(3.14159))
    
    With stdLambda.Create("if $1=1 and $2=""hello world"" then ""bart"" else ""lisa""")
        Debug.Print .Run(1, "hello world")
        Debug.Print .Run(1, 0)
    End With
    
    Debug.Assert False
End Sub
Public Sub zTestPerf()
    Dim i As Long, iStart As Double
    
    Dim lambda As stdLambda: Set lambda = stdLambda.Create("1+3*8/2*(2+2+3)+1+3*8/2*(2+2+3)+1+3*8/2*(2+2+3)+1+3*8/2*(2+2+3)+1+3*8/2*(2+2+3)+1+3*8/2*(2+2+3)")
    
    iStart = Timer
    For i = 1 To 10 ^ 2
        Call lambda.Run
    Next
    Debug.Print "Lambda: " & (Timer - iStart)
    
    iStart = Timer
    For i = 1 To 10 ^ 2
        Call Application.Evaluate("1+3*8/2*(2+2+3)+1+3*8/2*(2+2+3)+1+3*8/2*(2+2+3)+1+3*8/2*(2+2+3)+1+3*8/2*(2+2+3)+1+3*8/2*(2+2+3)")
    Next
    Debug.Print "Application.Evaluate: " & (Timer - iStart)
    
End Sub



Private Function expression() As Variant
    Dim res As Variant
    If optConsume("if") Then
        Dim condition As Variant: Call CopyVariant(condition, logic())
        consume ("then")
        Dim ifTrue As Variant: Call CopyVariant(ifTrue, expression())
        consume ("else")
        Dim ifFalse As Variant: Call CopyVariant(ifFalse, expression())
        If condition Then
            Call CopyVariant(res, ifTrue)
        Else
            Call CopyVariant(res, ifFalse)
        End If
    Else
        Call CopyVariant(res, logic())
    End If
    
    Call CopyVariant(expression, res)
End Function

Private Function logic()
    Dim res As Variant: Call CopyVariant(res, clause())
    Dim bLoop As Boolean: bLoop = True
    Do
        If optConsume("or") Then
            res = res And subClause() 'Can't break out if value is truthy unless we return an ast and execute from there...
        Else
            bLoop = False
        End If
    Loop While bLoop
    Call CopyVariant(logic, res)
End Function

Private Function clause()
    Dim res As Variant: Call CopyVariant(res, subClause())
    Dim bLoop As Boolean: bLoop = True
    Do
        If optConsume("and") Then
            res = res And subClause() 'Can't break out if value is falsy unless we return an ast and execute from there...
        Else
            bLoop = False
        End If
    Loop While bLoop
    Call CopyVariant(clause, res)
End Function

Private Function subClause()
    Dim res As Variant: Call CopyVariant(res, comparee())
    Dim bLoop As Boolean: bLoop = True
    Do
        If optConsume("lessThan") Then
            res = res < comparee()
        ElseIf optConsume("lessThanEqual") Then
            res = res <= comparee()
        ElseIf optConsume("greaterThan") Then
            res = res > comparee()
        ElseIf optConsume("greaterThanEqual") Then
            res = res >= comparee()
        ElseIf optConsume("equal") Then
            res = res = comparee()
        ElseIf optConsume("notEqual") Then
            res = res <> comparee()
        ElseIf optConsume("is") Then
            res = res Is comparee()
        ElseIf optConsume("like") Then
            res = res Like comparee()
        Else
            bLoop = False
        End If
    Loop While bLoop
    Call CopyVariant(subClause, res)
End Function

Private Function comparee() As Variant
    Dim res As Variant: Call CopyVariant(res, term())
    Dim bLoop As Boolean: bLoop = True
    Do
        If optConsume("add") Then
            res = res + term()
        ElseIf optConsume("subtract") Then
            res = res - term()
        Else
            bLoop = False
        End If
    Loop While bLoop
    Call CopyVariant(comparee, res)
End Function

Private Function term() As Variant
    Dim res As Variant: Call CopyVariant(res, factor())
    Dim bLoop As Boolean: bLoop = True
    Do
        If optConsume("power") Then
            res = res ^ factor()
        ElseIf optConsume("multiply") Then
            res = res * factor()
        ElseIf optConsume("divide") Then
            res = res / factor()
        ElseIf optConsume("concatenate") Then
            res = res & factor()
        ElseIf optConsume("modulus") Then
            res = res Mod factor()
        Else
            bLoop = False
        End If
    Loop While bLoop
    Call CopyVariant(term, res)
End Function

Private Function factor() As Variant
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
    ElseIf peek("funcName") Then
        Call CopyVariant(res , consumeFunction())
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

Private Function consumeFunction() As Variant
    Dim sFuncName As String: sFuncName = consume("funcName")
    Dim args() As Variant: args = optConsumeParameters()
    Call CopyVariant(consumeFunction, evaluateFunc(sFuncName, args))
End Function

Private Function manyAccessors(value As Variant) As Variant
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

Private Function optObjectProperty(ByRef value As Variant)
    If IsObject(value) Then
        If optConsume("propertyAccess") Then
            Dim sFuncName As String: sFuncName = consume("funcName")
            Dim args() As Variant: args = optConsumeParameters()
            
            'Call rtcCallByName
            Dim hr As Long, res As Variant
            hr = rtcCallByName(res, value, StrPtr(sFuncName), VbCallType.VbGet, args, &H409)
            
            'If error then raise
            If hr < 0 Then
                Call Throw("Error in calling " & sFuncName & " property of " & typename(value) & " object.")
            Else
                Call CopyVariant(optObjectProperty, res)
            End If
            Exit Function
        End If
    End If
    
    optObjectProperty = UniqueConst
End Function

Private Function optObjectMethod(ByRef value As Variant) As Variant
    If IsObject(value) Then
        If optConsume("methodAccess") Then
            Dim sFuncName As String: sFuncName = consume("funcName")
            Dim args() As Variant: args = optConsumeParameters()
            
            'Call rtcCallByName
            Dim hr As Long, res As Variant
            hr = rtcCallByName(res, value, StrPtr(sFuncName), VbCallType.VbMethod, args, &H409)
            
            'If error then raise
            If hr < 0 Then
                Call Throw("Error in calling " & sFuncName & " method of " & typename(value) & " object.")
            Else
                Call CopyVariant(optObjectMethod, res)
            End If
            Exit Function
        End If
    End If
    
    optObjectMethod = UniqueConst
End Function

Private Function optConsumeParameters() As Variant
    If optConsume("lBracket") Then
        Dim arguments() As Variant
        arguments = Array()

        While Not peek("rBracket")
            If UBound(arguments) - LBound(arguments) + 1 > 0 Then
                Call consume("comma")
                ReDim Preserve arguments(0 To UBound(arguments) + 1)
            Else
                ReDim arguments(0 To 0)
            End If
            arguments(UBound(arguments)) = expression()
        Wend

        Call consume("rBracket")
        optConsumeParameters = arguments
    Else
        optConsumeParameters = Array()
    End If
End Function

Private Function consumeString() As String
    Dim sRes As String: sRes = consume("literalString")
    sRes = Mid(sRes, 2, Len(sRes) - 2)
    sRes = Replace(sRes, """""", """")
    consumeString = sRes
End Function


'Evaluates Variable name
'@param {string} sValueName Value like $\d+ to be interpreted as argument
Private Function EvaluateVarName(ByVal sValueName As String) As Variant
    Dim iArgIndex As Long: iArgIndex = Val(Mid(sValueName, 2)) + LBound(vLastArgs) - 1
    
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
    
    Tokenise = removeTokens(tokens, "space")
End Function

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
    i = i + 1: arr(i) = getTokenDefinition("like", "like")

    'Structural
    ' Inline if
    i = i + 1: arr(i) = getTokenDefinition("if", "if")
    i = i + 1: arr(i) = getTokenDefinition("then", "then")
    i = i + 1: arr(i) = getTokenDefinition("else", "else")
    ' Brackets
    i = i + 1: arr(i) = getTokenDefinition("lBracket", "\(")
    i = i + 1: arr(i) = getTokenDefinition("rBracket", "\)")
    ' Functions
    i = i + 1: arr(i) = getTokenDefinition("funcName", "[a-zA-Z][a-zA-Z0-9_]+")
    i = i + 1: arr(i) = getTokenDefinition("comma", ",") 'params
    
    'VarName
    i = i + 1: arr(i) = getTokenDefinition("var", "\$\d+")
    
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
    i = i + 1: arr(i) = getTokenDefinition("greaterThan", "\>")
    i = i + 1: arr(i) = getTokenDefinition("greaterThanEqual", "\>\=")
    i = i + 1: arr(i) = getTokenDefinition("lessThan", "\<")
    i = i + 1: arr(i) = getTokenDefinition("lessThanEqual", "\<\=")
    i = i + 1: arr(i) = getTokenDefinition("concatenate", "\&")
    
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
Private Function evaluateFunc(ByVal sFuncName As String, ByVal args As Variant) As Variant
    Dim iArgStart As Long: iArgStart = LBound(args)
    If typename(stdLambda.oFunctExt) = "Dictionary" Then
        If Me.oFunctExt.exists(sFuncName) Then
            evaluateFunc = Me.oFunctExt(sFuncName).RunEx(args)
            Exit Function
        End If
    End If
    
    Select Case LCase(sFuncName)
        Case "eval": evaluateFunc = stdLambda.Create(args(iArgStart)).Run()

        'Useful OOP constants
        Case "thisworkbook": set evaluateFunc = ThisWorkbook
        Case "application":  set evaluateFunc = Application

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
        Case "clng":  evaluateFunc = VBA.Conversion.Clng(args(iArgStart))
        Case "cstr":  evaluateFunc = VBA.Conversion.CStr(args(iArgStart))
        Case "cvar":  evaluateFunc = VBA.Conversion.CVar(args(iArgStart))
        Case "cverr":   evaluateFunc = VBA.Conversion.CVErr(args(iArgStart))
        
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
        Case "val":     evaluateFunc = VBA.Conversion.Val(args(iArgStart))
        
        'String functions
        Case "trim": evaluateFunc = VBA.Trim(args(iArgStart))
        Case "lcase": evaluateFunc = VBA.LCase(args(iArgStart))
        Case "ucase": evaluateFunc = VBA.UCase(args(iArgStart))
        Case "right": evaluateFunc = VBA.Right(args(iArgStart), args(iArgStart + 1))
        Case "left": evaluateFunc = VBA.Left(args(iArgStart), args(iArgStart + 1))
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

'==============================================================================================================================
'
'Helper Methods:
'
'==============================================================================================================================

'Copies one variant to a destination
'@param {ByRef Token()} tokens Tokens to remove the specified type from
'@param {string} sRemoveType   Token type to remove.
'@returns {Token()} The modified token array.
Private Function removeTokens(ByRef tokens() As token, ByVal sRemoveType As String) As token()
    Dim iCountRemoved As Long: iCountRemoved = 0
    Dim iToken As Long
    For iToken = LBound(tokens) To UBound(tokens)
        If tokens(iToken).Type.Name <> sRemoveType Then
            tokens(iToken - iCountRemoved) = tokens(iToken)
        Else
            iCountRemoved = iCountRemoved + 1
        End If
    Next
    ReDim Preserve tokens(LBound(tokens) To (UBound(tokens) - iCountRemoved))
    removeTokens = tokens
End Function


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



