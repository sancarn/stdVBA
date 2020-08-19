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
#If VBA7 Then
  'VBE7 is interchangable with msvbvm60.dll    however VBE7.dll appears to always be present where as msvbvm60 is only occasionally present.
  Private Declare PtrSafe Function rtcCallByName Lib "VBE7.dll" (ByRef vRet As Variant, ByVal cObj As Object, ByVal sMethod As LongPtr, ByVal eCallType As VbCallType, ByRef pArgs() As Variant, ByVal lcid As Long) As Long
  Private Declare PtrSafe Sub VariantCopy Lib "oleaut32.dll" (ByRef pvargDest As Variant, ByRef pvargSrc As Variant)
#Else
  Private Declare Function rtcCallByName Lib "msvbvm60" (ByRef vRet As Variant, ByVal cObj As Object, ByVal sMethod As LongPtr, ByVal eCallType As VbCallType, ByRef pArgs() As Variant, ByVal lcid As Long) As Long
  Private Declare Sub VariantCopy Lib "oleaut32.dll" (ByRef pvargDest As Variant, ByRef pvargSrc As Variant)
#End If
#Const devMode = True 'Used for enabling some debugging features

'Tokens, token definitions and operations
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
Public Sub Init(ByVal sEquation As String, ByVal bSandboxExtras As Boolean)
    If Not bSandboxExtras Then
        Set Me.oFunctExt = stdLambda.oFunctExt
    End If
    tokens = Tokenise(sEquation)
    iTokenIndex = 1
    iOperationIndex = 0
    Call finishOperations
    Call parseExpressionPriority1
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

Private Sub parseExpressionPriority1()
    Call parseLogicPriority1
    'todo:
'    If optConsume("if") Then
'        Dim condition As Variant: Call CopyVariant(condition, expression())
'        consume ("then")
'        Dim ifTrue As Variant: Call CopyVariant(ifTrue, expression())
'        consume ("else")
'        Dim ifFalse As Variant: Call CopyVariant(ifFalse, expression())
'        If condition Then
'            Call CopyVariant(res, ifTrue)
'        Else
'            Call CopyVariant(res, ifFalse)
'        End If
'    Else
'        Call CopyVariant(res, logic())
'    End If
End Sub

Private Sub parseLogicPriority1() 'and
    Call parseLogicPriority2
    Dim bLoop As Boolean: bLoop = True
    Do
        If optConsume("xor") Then
            Call parseLogicPriority2
            Call addOperation(oLogic, oXor)
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
            Call addOperation(oLogic, oOr)
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
            Call addOperation(oLogic, oOr)
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
            Call addOperation(oComparison, oLt)
        ElseIf optConsume("lessThanEqual") Then
            Call parseArithmeticPriority1
            Call addOperation(oComparison, oLte)
        ElseIf optConsume("greaterThan") Then
            Call parseArithmeticPriority1
            Call addOperation(oComparison, oGt)
        ElseIf optConsume("greaterThanEqual") Then
            Call parseArithmeticPriority1
            Call addOperation(oComparison, oGte)
        ElseIf optConsume("equal") Then
            Call parseArithmeticPriority1
            Call addOperation(oComparison, oEql)
        ElseIf optConsume("notEqual") Then
            Call parseArithmeticPriority1
            Call addOperation(oComparison, oNeq)
        ElseIf optConsume("is") Then
            Call parseArithmeticPriority1
            Call addOperation(oComparison, oIs)
        ElseIf optConsume("like") Then
            Call parseArithmeticPriority1
            Call addOperation(oComparison, oLike)
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
            Call addOperation(oMisc, oCat)
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
            Call addOperation(oArithmetic, oAdd)
        ElseIf optConsume("subtract") Then
            Call parseArithmeticPriority3
            Call addOperation(oArithmetic, oSub)
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
            Call addOperation(oArithmetic, oMod)
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
            Call addOperation(oArithmetic, oMul)
        ElseIf optConsume("divide") Then
            Call parseArithmeticPriority4
            Call addOperation(oArithmetic, oDiv)
        Else
            bLoop = False
        End If
    Loop While bLoop
End Sub

Private Sub parseArithmeticPriority5() '^
    Call parseValuePriority1
    Do
        If optConsume("power") Then
            Call parseValuePriority1
            Call addOperation(oArithmetic, oPow)
        Else
            bLoop = False
        End If
    Loop While bLoop
End Sub

Private Sub parseValuePriority1() 'numbers, $vars, strings, booleans, (expressions)
    If peek("literalNumber") Then
        Call addOperation(oPush, , CDbl(consume("literalNumber")))
    ElseIf peek("var") Then
        Call addOperation(oAccess, argument, consume("var"))
        Call parseManyAccessors
    ElseIf peek("literalString") Then
        Call parseString
    ElseIf peek("literalBoolean") Then
        Call addOperation(oPush, , consume("literalBoolean") = "true")
    ElseIf peek("funcName") Then
        'todo:
        'Call CopyVariant(res, consumeFunction())
        Call parseManyAccessors
    Else
        Call consume("lBracket")
        Call parseExpressionPriority1
        Call consume("rBracket")
        Call parseManyAccessors
    End If
End Sub

Private Function consumeFunction() As Variant
    Dim sFuncName As String: sFuncName = consume("funcName")
    Dim args() As Variant: args = optConsumeParameters()
    Call CopyVariant(consumeFunction, evaluateFunc(sFuncName, args))
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
        Call addOperation(oPush, , consume("funcName"))
        Call parseOptParameters
        Call addOperation(oObject, oPropGet)
        parseOptObjectProperty = True
    End If
End Function

Private Function parseOptObjectMethod() As Boolean
    parseOptObjectMethod = False
    If optConsume("methodAccess") Then
        Call addOperation(oPush, , consume("funcName"))
        Call parseOptParameters
        Call addOperation(oObject, oMethodCall)
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
            Call parseExpressionPriority1
            iArgCount = iArgCount + 1
        Wend
        Call consume("rBracket")
        Call addOperation(oPush, , iArgCount)
        parseOptParameters = True
    End If
End Function

Private Sub parseString()
    Dim sRes As String: sRes = consume("literalString")
    sRes = Mid(sRes, 2, Len(sRes) - 2)
    sRes = Replace(sRes, """""", """")
    Call addOperation(oPush, , sRes)
End Sub

'todo: add to evaluation method
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
        Case "val":     evaluateFunc = VBA.Conversion.val(args(iArgStart))
        
        'String functions
        Case "Trim": evaluateFunc = VBA.Trim(args(iArgStart))
        Case "LCase": evaluateFunc = VBA.LCase(args(iArgStart))
        Case "UCase": evaluateFunc = VBA.UCase(args(iArgStart))
        Case "Right": evaluateFunc = VBA.Right(args(iArgStart), args(iArgStart + 1))
        Case "Left": evaluateFunc = VBA.Left(args(iArgStart), args(iArgStart + 1))
        Case "Mid"
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
        Call Throw("Error in calling " & sFuncName & " property of " & typename(value) & " object.")
    Else
        If op.subType = ISubType.oPropGet Or op.subType = ISubType.oMethodCall Then
            Call pushV(stack, stackPtr, res)
        End If
    End If
End Sub

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
            Call Throw("Syntax Error - Lexer Error")
        End If
    Wend
    
    Tokenise = removeTokens(tokens, "space")
End Function

Private Function getTokenDefinition(ByVal sName As String, ByVal sRegex As String, Optional ByVal ignoreCase As Boolean = True) As TokenDefinition
    getTokenDefinition.Name = sName
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
        If tokens(iToken).Type.Name <> sRemoveType Then
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
 
'Adds an operation to the instance operations list
'@param {IType} kType       The main type of the operation
'@param {ISubType} subType  The sub type of the operation
'@param {Variant} value     The value associated with the operation
'@returns {Operation} Returns the created operation
Private Function addOperation(kType As iType, Optional subType As ISubType, Optional value As Variant) As Operation
    Dim size As Long: size = UBound(operations)
    If size = 0 Then
        ReDim Preserve operations(0 To 1)
    ElseIf iOperationIndex > size Then
        ReDim Preserve operations(0 To size * 2)
    End If
    
    With operations(iOperationIndex)
        .Type = kType
        .subType = subType
        Call CopyVariant(.value, value)
    End With
    addOperation = operations(iOperationIndex)
    
    iOperationIndex = iOperationIndex + 1
End Function

'Resizes the operations list so there are no more empty operations
Private Sub finishOperations()
    ReDim Preserve operations(0 To iOperationsIndex)
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




