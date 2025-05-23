VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdLambda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Ensure Option-Explicit is disabled!! 
'For VB6 compatibility we rely on the auto-definition of Application and ThisWorkbook
'Option Explicit

'Used for enabling some debugging features
#Const devMode = True 

'For Mac testing purposes only
'#const Mac = true



'Implement stdICallable interface
Implements stdICallable

'Direct call convention of VBA.CallByName
#If Not Mac Then
    #If VBA7 Then
        'VBE7 is interchangable with msvbvm60.dll    however VBE7.dll appears to always be present where as msvbvm60 is only occasionally present.
        Private Declare PtrSafe Function rtcCallByName Lib "VBE7.dll" (ByVal cObj As Object, ByVal sMethod As LongPtr, ByVal eCallType As VbCallType, ByRef pArgs() As Variant, ByVal lcid As Long) As Variant
    #Else
        Private Declare Function rtcCallByName Lib "msvbvm60" (ByVal cObj As Object, ByVal sMethod As Long, ByVal eCallType As VbCallType, ByRef pArgs() As Variant, ByVal lcid As Long) As Variant
    #End If
#End If


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
    oFieldCall = 29
    oEquality = 30    'Yet to be implemented
    oIsOperator = 31  'Yet to be implemented
    oEnum = 32        'Yet to be implemented
End Enum

Private Enum LambdaType
    iStandardLambda = 1
    iBoundLambda = 2
End Enum

'Special constant used in parsing:
Private Const UniqueConst As String = "3207af79-30df-4890-ade1-640f9f28f309"
Private Const vbGetOrMethod as Long = VbGet Or VbMethod
Private Const minStackSize as Long = 30 'note that the stack size may become smaller than this

'@TODO: Convert to TThis
Private Type TSingleton
    Cache As Object
End Type
Private Type TBoundLambda
    Lambda As stdLambda  'was oBound
    Args As Variant      'was vBound
End Type
Private Type TThis
    Singleton As TSingleton

    tokens() As token
    iTokenIndex As Long
    operations() As Operation
    iOperationIndex As Long
    stackSize As Long
    scopes() As Variant
    scopesArgCount() As Variant
    scopeCount As Long
    funcScope As Long
    Equation as string

    'Instance var on both class and object
    FunctionExtensions As Object
    isBoundLambda as Boolean
    Bound As TBoundLambda
    
    UsePerformanceCache as Boolean
    PerformanceCache as Object

    ExitStackName() As String
    ExitStackIndex() As Long
    ExitStackSize As Long
End Type
Private This as TThis








'Create a stdLambda object from a string equation
'@constructor
'@param sEquation - The equation to parse
'@param bUsePerformanceCache - Whether to cache the parsed equation for performance
'@param bSandboxExtras - Whether to allow extra functions to be called
'@returns - A first class object representing the equation. Can be called with `.Run(param1, param2, ...)` or `.RunEx(params)`
'@example ```vb
'Debug.Print stdLambda.Create("1+3*8/2*(2+2+3)").Run()
'With stdLambda.Create("$1+1+3*8/2*(2+2+3)")
'    Debug.Print .Run(10)
'    Debug.Print .Run(15)
'    Debug.Print .Run(20)
'End With
'Debug.Print stdLambda.Create("$1.Range(""A1"")").Run(Sheets(1)).Address(True, True, xlA1, True)
'Debug.Print stdLambda.Create("$1.join("","")").Run(stdArray.Create(1,2))
'```
Public Function Create(ByVal sEquation As String, Optional ByVal bUsePerformanceCache As Boolean = False, Optional ByVal bSandboxExtras As Boolean = False) As stdLambda
    'Cache Lambda created
    If this.Singleton.Cache Is Nothing Then Set this.Singleton.Cache = CreateObject("Scripting.Dictionary")
    Dim sID As String: sID = bUsePerformanceCache & "-" & bSandboxExtras & ")" & sEquation
    If Not this.Singleton.Cache.exists(sID) Then
        Set this.Singleton.Cache(sID) = New stdLambda
        Call this.Singleton.Cache(sID).protInit(LambdaType.iStandardLambda, sEquation, bUsePerformanceCache, bSandboxExtras)
    End If
    
    'Return cached lambda
    Set Create = this.Singleton.Cache(sID)
End Function

'Create a stdLambda object from an Array of strings
'@constructor
'@param sEquation as Variant<Array<String>> - The equation to parse
'@param bUsePerformanceCache - Whether to cache the parsed equation for performance
'@param bSandboxExtras - Whether to allow extra functions to be called
'@returns - A first class object representing the equation. Can be called with `.Run(param1, param2, ...)` or `.RunEx(params)`
'@example ```vb
'Debug.Print stdLambda.CreateMultiline(Array( _
'  "let x = 8/2*(2+2+3)", _
'  "1+3*x"
')).Run()
'```
Public Function CreateMultiline(ByRef sEquation As Variant, Optional ByVal bUsePerformanceCache As Boolean = False, Optional ByVal bSandboxExtras As Boolean = False) As stdLambda
    Set CreateMultiline = Create(Join(sEquation, " "), bUsePerformanceCache, bSandboxExtras)
End Function

'Initialise the lambda
'@protected
'@param iLambdaType as LambdaType - Type of lambda to initialise
'@param params - Parameters to initialise the lambda with
Public Sub protInit(ByVal iLambdaType As Long, ParamArray params() As Variant)
    Select Case iLambdaType
        Case LambdaType.iStandardLambda
            Dim sEquation As String: sEquation = params(0)
            this.Equation = sEquation
            this.UsePerformanceCache = params(1)
            Dim bSandboxExtras As Boolean: bSandboxExtras = params(2)
            
            'Performance cache
            if this.UsePerformanceCache then set this.PerformanceCache = CreateObject("Scripting.Dictionary")

            'Function extensions
            Set this.FunctionExtensions = stdLambda.oFunctExt
            If bSandboxExtras OR this.FunctionExtensions is nothing Then set this.FunctionExtensions = CreateObject("Scripting.Dictionary")

            this.isBoundLambda = false
            this.tokens = Tokenise(sEquation)
            this.iTokenIndex = 1
            this.iOperationIndex = 0
            this.stackSize = 0
            this.scopeCount = 0
            this.funcScope = 0
            Call parseBlock("eof")
            Call finishOperations

        Case LambdaType.iBoundLambda
            this.isBoundLambda = True
            Set this.Bound.Lambda = params(0)
            this.Bound.Args = params(1)
            this.Equation = "BOUND..."
            
            'Function extensions
            Set this.FunctionExtensions = stdLambda.oFunctExt
            If bSandboxExtras OR this.FunctionExtensions is nothing Then set this.FunctionExtensions = CreateObject("Scripting.Dictionary")
        Case Else
            Err.Raise 1, "stdLambda::Init", "No lambda with that type."
    End Select
End Sub

'Run the lambda from passed parameters
'@param params as Array<Variant> - Array of parameters to run the lambda with
'@returns - The result of the lambda
Public Function Run(ParamArray params() As Variant) As Variant
Attribute Run.VB_UserMemId = 0
    If Not this.isBoundLambda Then
        'Execute top-down parser
        Call CopyVariant(Run, evaluate(this.operations, params))
    Else
        Call CopyVariant(Run, this.Bound.Lambda.RunEx(ConcatArrays(this.Bound.Args, params)))
    End If
End Function

'Run the lambda from an array of parameters
'@param params as Variant<Array<Variant>> - Array of parameters to run the lambda with
'@returns - The result of the lambda
Public Function RunEx(ByVal params As Variant) As Variant
    If Not this.isBoundLambda Then
        If Not isArray(params) Then
            Err.Raise 1, "params to be supplied as array of arguments", ""
        End If
        
        'Execute top-down parser
        Call CopyVariant(RunEx, evaluate(this.operations, params))
    Else
        Call CopyVariant(RunEx, this.Bound.Lambda.RunEx(ConcatArrays(this.Bound.Args, params)))
    End If
End Function

'Bind parameters to the function.  Arguments will be passed in the order they are supplied, before any arguments supplied to the function.
'@param params as Array<Variant> - Parameters to bind to the lambda.
'@returns - The lambda 
Public Function Bind(ParamArray params() As Variant) As stdLambda
    Set Bind = BindEx(params)
End Function

'Bind an array of parameters to the function. Arguments will be passed in the order they are supplied, before any arguments supplied to the function.
'@param params as Variant<Array<Variant>> - Array of parameters to bind
'@returns - The lambda existing lambda
Public Function BindEx(ByVal params As Variant) As stdLambda
    Set BindEx = New stdLambda
    Dim callable As stdICallable: Set callable = Me
    Call BindEx.protInit(LambdaType.iBoundLambda, callable, params)
End Function

'Bind a named global variable to the function
'@param sGlobalName - New global name
'@param variable - Data to store in global variable
'@returns - The lambda existing lambda
Public Function BindGlobal(ByVal sGlobalName as string, ByVal variable as Variant) as stdLambda
    set BindGlobal = Me
    If this.isBoundLambda Then
        Call this.Bound.Lambda.BindGlobal(sGlobalName, variable)
    Else
        If this.FunctionExtensions is nothing then Set this.FunctionExtensions = CreateObject("Scripting.Dictionary")
        If IsObject(variable) Then
            Set this.FunctionExtensions(sGlobalName) = variable
        Else
            Let this.FunctionExtensions(sGlobalName) = variable
        End If
    End If
End Function

'Extend the lambda with new functions and named global variables
'@returns Object<Dictionary<string,stdICallable> | Dictionary<string,variant>> - Dictionary of functions and named global variables
Public Property Get oFunctExt() as Object
  set oFunctExt = this.FunctionExtensions
End Property

'Implementation of stdICallable::Run
'@param params as Array<Variant> - Parameters to run the lambda with
'@returns - The result of the lambda
Private Function stdICallable_Run(ParamArray params() As Variant) As Variant
    If Not this.isBoundLambda Then
        'Execute top-down parser
        Call CopyVariant(stdICallable_Run, evaluate(this.operations, params))
    Else
        Call CopyVariant(stdICallable_Run, this.Bound.Lambda.RunEx(ConcatArrays(this.Bound.Args, params)))
    End If
End Function

'Implementation of stdICallable::RunEx
'@param params as Variant<Array<Variant>> - Array of parameters to run the lambda with
'@returns - The result of the lambda
Private Function stdICallable_RunEx(ByVal params As Variant) As Variant
    If Not isArray(params) Then
        Err.Raise 1, "params to be supplied as array of arguments", ""
    End If
    
    If Not this.isBoundLambda Then
        'Execute top-down parser
        Call CopyVariant(stdICallable_RunEx, evaluate(this.operations, params))
    Else
        Call CopyVariant(stdICallable_RunEx, this.Bound.Lambda.RunEx(ConcatArrays(this.Bound.Args, params)))
    End If
End Function

'Implementation of stdICallable::Bind
'@param params as Array<Variant> - Parameters to bind to the lambda.
'@returns - The bound lambda
Private Function stdICallable_Bind(ParamArray params() As Variant) As stdICallable
    Set stdICallable_Bind = BindEx(params)
End Function

'Late-bound, no dependency, function calling
'@param sMessage as "obj"|"className"|"bindGlobal" - Message to send. "obj" returns the object, "className" returns the class name, "bindGlobal" binds a global variable to the lambda.
'@param success  - Success of message. If message wasn't processed return false.
'@param params   - Parameters to pass along with message
'@returns - Anything returned by the function
Private Function stdICallable_SendMessage(ByVal sMessage as string, ByRef success as boolean, ByVal params as variant) as Variant
    select case sMessage
        case "obj"
            set stdICallable_SendMessage = Me
            success = true
        case "className"
            stdICallable_SendMessage = "stdLambda"
            success = true
        case "bindGlobal"
            'Bind global based whether this is a bound lambda or not
            Call BindGlobal(params(0), params(1))
            success = true
        case else 
            success = false
    end select
End Function

'================
'
'  TOKENISATION
'
'================


'Get token definitions, a mapping of token names to regexes. This is used in the tokenisation process.
'@private
'@returns - Array of token definitions
Private Function getTokenDefinitions() As TokenDefinition()
    Dim arr() As TokenDefinition
    ReDim arr(1 To 99)
    
    Dim i As Long: i = 0
    'Whitespace
    i = i + 1: arr(i) = getTokenDefinition("space", "\s+") 'String

    'Literal
    i = i + 1: arr(i) = getTokenDefinition("literalString", """(?:""""|[^""])*""") 'String
    i = i + 1: arr(i) = getTokenDefinition("literalNumber", "\d+(?:\.\d+)?") 'Number
    i = i + 1: arr(i) = getTokenDefinition("literalBoolean", "True|False", isKeyword:=True)
    
    'Named operators
    i = i + 1: arr(i) = getTokenDefinition("is", "is", isKeyword:=True)
    i = i + 1: arr(i) = getTokenDefinition("mod", "mod", isKeyword:=True)
    i = i + 1: arr(i) = getTokenDefinition("and", "and", isKeyword:=True)
    i = i + 1: arr(i) = getTokenDefinition("or", "or", isKeyword:=True)
    i = i + 1: arr(i) = getTokenDefinition("xor", "xor", isKeyword:=True)
    i = i + 1: arr(i) = getTokenDefinition("not", "not", isKeyword:=True)
    i = i + 1: arr(i) = getTokenDefinition("like", "like", isKeyword:=True)
    i = i + 1: arr(i) = getTokenDefinition("let", "let", isKeyword:=True)
    i = i + 1: arr(i) = getTokenDefinition("set", "set", isKeyword:=True)

    'Structural
    ' Inline if
    i = i + 1: arr(i) = getTokenDefinition("if", "if", isKeyword:=True)
    i = i + 1: arr(i) = getTokenDefinition("then", "then", isKeyword:=True)
    i = i + 1: arr(i) = getTokenDefinition("else", "else", isKeyword:=True)
    i = i + 1: arr(i) = getTokenDefinition("end", "end", isKeyword:=True)

    ' Loops
    i = i + 1: arr(i) = getTokenDefinition("for", "for", isKeyword:=True)
    i = i + 1: arr(i) = getTokenDefinition("to", "to", isKeyword:=True)
    i = i + 1: arr(i) = getTokenDefinition("step", "step", isKeyword:=True)
    i = i + 1: arr(i) = getTokenDefinition("next", "next", isKeyword:=True)

    i = i + 1: arr(i) = getTokenDefinition("do", "do", isKeyword:=True)
    i = i + 1: arr(i) = getTokenDefinition("until", "until", isKeyword:=True)
    i = i + 1: arr(i) = getTokenDefinition("while", "while", isKeyword:=True)
    i = i + 1: arr(i) = getTokenDefinition("loop", "loop", isKeyword:=True)

    i = i + 1: arr(i) = getTokenDefinition("exit", "exit", isKeyword:=True)


    ' Brackets
    i = i + 1: arr(i) = getTokenDefinition("lBracket", "\(")
    i = i + 1: arr(i) = getTokenDefinition("rBracket", "\)")
    ' Functions
    i = i + 1: arr(i) = getTokenDefinition("fun", "fun", isKeyword:=True)
    i = i + 1: arr(i) = getTokenDefinition("comma", ",") 'params
    ' Lines
    i = i + 1: arr(i) = getTokenDefinition("colon", ":")
    
    'VarName
    i = i + 1: arr(i) = getTokenDefinition("arg", "\$\d+")
    i = i + 1: arr(i) = getTokenDefinition("var", "[a-zA-Z][a-zA-Z0-9_]*")
    
    'Operators
    i = i + 1: arr(i) = getTokenDefinition("propertyAccess", "\.\$")
    i = i + 1: arr(i) = getTokenDefinition("methodAccess", "(\.\#)")
    i = i + 1: arr(i) = getTokenDefinition("fieldAccess", "\.")
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

'Continually parse statements until some endToken is reached.
'@private
'@param endToken as Array<String> - Names of tokens which should be checked to stop parsing statements
'@remark - Entry point for parsing
Private Sub parseBlock(ParamArray endToken() As Variant)
    Call addScope
    Dim size As Integer: size = this.stackSize + 1
    
    ' Consume multiple lines
    Dim bLoop As Boolean: bLoop = True
    Do
        While optConsume("colon"): Wend
        Call parseStatement
        While optConsume("colon"): Wend
        
        For i = LBound(endToken) To UBound(endToken)
            If peek(endToken(i)) Then
                bLoop = False
            End If
        Next
    Loop While bLoop
    
    ' Get rid of all extra expression results and declarations
    While this.stackSize > size
        Call addOperation(oMerge, , , -1)
    Wend
    this.scopeCount = this.scopeCount - 1
End Sub

'Increment the number of scopes and initialise them. Scopes are used to store variables, functions and function arg counts.
Private Sub addScope()
    this.scopeCount = this.scopeCount + 1
    Dim scope As Long: scope = this.scopeCount
    ReDim Preserve this.scopes(1 To scope)
    ReDim Preserve this.scopesArgCount(1 To scope)
    Set this.scopes(scope) = CreateObject("Scripting.Dictionary")
    Set this.scopesArgCount(scope) = CreateObject("Scripting.Dictionary")
End Sub

'Parse a statement of code.
'@remark - A statement consists of either a variable assignment, a function declaration or an expression
'(typically this wouldn't classify as a statement, but for the purpose of stdLambda and simplifying parsing it does).
Private Sub parseStatement()
    If (peek("set") or peek("let")) and peek("var",2) And peek("equal", 3) Then
        Call parseAssignment
    ElseIf peek("fun") Then
        Call parseFunctionDeclaration
    Else
        Call parseExpression
    End If
End Sub

'Parse expression.
'@remark - Parsing an expression is split into a number of stages, each with a different priority.
'This is to ensure that the correct order of operations is followed. E.G. think about the expression `1+2*3`. We want to ensure
'that the multiplication is performed before the addition. This is done by parsing the expression in a specific order. The order
'is defined by the priority of the operation. The order of operations is as follows:
'1. Logical XOR
'2. Logical OR
'3. Logical AND
'4. Logical NOT
'5. Comparison (=, <>, <, <=, >, >=, is, Like)
'6. Concatenation (&)
'7. Arithmetic (+, -)
'8. Arithmetic (mod)
'9. Arithmetic (*, /)
'10. Arithmetic (Unary +, -) e.g. -1
'11. Arithmetic (^)
'12. Arithmetic (Unary +, -) (for RHS of power operator) e.g. 2^-1
'13. Flow (if then else)
'14. Do (do until while loop)
'15. For (for do step next)
'16. Value (numbers, $vars, strings, booleans, brackets)
'@remark - The order of priority is opposite to the order of evaluation. I.E. Comparrison is evaluated before Logical AND allowing 
'expressions such as `1<2 and 2<3` to be evaluated correctly without requiring bracketing. It's important to note however that all
'comparrisons have the same priority. This means that `1<2<3` will be evaluated as `(1<2)<3` which is not the same as `1<(2<3)`.
Private Sub parseExpression()
    Call parseLogicPriority1
End Sub

'Parse Logical XOR
Private Sub parseLogicPriority1()
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

'Parse Logical OR
Private Sub parseLogicPriority2()
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

'Parse Logical AND
Private Sub parseLogicPriority3()
    Call parseLogicPriority4
    Dim bLoop As Boolean: bLoop = True
    Do
        If optConsume("and") Then
            Call parseLogicPriority4
            Call addOperation(oLogic, oAnd, , -1)
        Else
            bLoop = False
        End If
    Loop While bLoop
End Sub

'Parse Logical NOT
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

'Parse comparison operators (=, <>, <, <=, >, >=, is, Like, =)
Private Sub parseComparisonPriority1()
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

'Parse concatenation operator (&)
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

'Parse + and - operators
Private Sub parseArithmeticPriority2()
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

'Parse mod operator
Private Sub parseArithmeticPriority3() 'mod
    Call parseArithmeticPriority4
    Dim bLoop As Boolean: bLoop = True
    Do
        If optConsume("mod") Then
            Call parseArithmeticPriority4
            Call addOperation(oArithmetic, oMod, , -1)
        Else
            bLoop = False
        End If
    Loop While bLoop
End Sub

'Parse * and / operators
Private Sub parseArithmeticPriority4()
    Call parseArithmeticPriority5
    Dim bLoop As Boolean: bLoop = True
    Do
        If optConsume("multiply") Then
            Call parseArithmeticPriority5
            Call addOperation(oArithmetic, oMul, , -1)
        ElseIf optConsume("divide") Then
            Call parseArithmeticPriority5
            Call addOperation(oArithmetic, oDiv, , -1)
        Else
            bLoop = False
        End If
    Loop While bLoop
End Sub

'Parse unary + and - operators (i.e. -1) 
Private Sub parseArithmeticPriority5()
    If optConsume("subtract") Then
        Call parseArithmeticPriority5 'recurse
        Call addOperation(oArithmetic, oNeg)
    ElseIf optConsume("add") Then
        Call parseArithmeticPriority5 'recurse
    Else
        Call parseArithmeticPriority6
    End If
End Sub

'Parse power (^) operator
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

'Parse unary + and - operators on the RHS of a power operator (i.e. 2^-1)
Private Sub parseArithmeticPriority6andahalf()
    If optConsume("subtract") Then
        Call parseArithmeticPriority6andahalf 'recurse
        Call addOperation(oArithmetic, oNeg)
    ElseIf optConsume("add") Then
        Call parseArithmeticPriority6andahalf 'recurse
    Else
        Call parseFlowPriority1
    End If
End Sub

'Parse flow control (if ... then ... else ... end)
Private Sub parseFlowPriority1()
    If optConsume("if") Then
        Call parseExpression
        Dim skipThenJumpIndex As Integer: skipThenJumpIndex = addOperation(oJump, ifFalse, , -1)
        
        Dim size As Integer: size = this.stackSize
        Call consume("then")
        Call parseBlock("else", "end")
        Dim skipElseJumpIndex As Integer: skipElseJumpIndex = addOperation(oJump)
        this.operations(skipThenJumpIndex).value = this.iOperationIndex
        this.stackSize = size
        
        If optConsume("end") Then
            ' Why should it always return? I disabled it for now because in a loop it keeps pushing onto the stack Call addOperation(oPush, , 0, 1) 'Expressions should always return a value
            this.operations(skipElseJumpIndex).value = this.iOperationIndex
        Else
            Call consume("else")
            Call parseBlock("eof", "rBracket", "end")
            this.operations(skipElseJumpIndex).value = this.iOperationIndex
        
            Call optConsume("end")
        End If
        Call consumeExit("if", This.iOperationIndex)
    Else
        Call parseFlowPriority2
    End If
End Sub

'Parse loops (do ... until|while ... loop)
Private Sub parseFlowPriority2()
    If optConsume("do") Then
        ' Do Until|While {Expr}
        Dim RepeatLoop As Integer: RepeatLoop = This.iOperationIndex
        Dim SkipLoop   As Integer: SkipLoop   = -1
        If optConsume("until") Then
            Call parseExpression
            SkipLoop = addOperation(oJump, ifTrue, , -1)
        ElseIf optConsume("while") Then
            Call parseExpression
            SkipLoop = addOperation(oJump, ifFalse, , -1)
        Else
            ' Infinite loop as long as no exit do is used
        End If

        Call parseBlock("loop")
        Call consume("loop")

        ' xxx
        ' Loop Until|While {expr}
        If optConsume("until") Then
            Call parseExpression
            SkipLoop = addOperation(oJump, ifTrue, , -1)
        ElseIf optConsume("while") Then
            Call parseExpression
            SkipLoop = addOperation(oJump, ifFalse, , -1)
        Else
            ' Nothing happens
        End If

        ' Loop back to start
        Call addOperation(oJump, , RepeatLoop, -1)
        Call consumeExit("do", This.iOperationIndex)
        If SkipLoop > -1 Then This.operations(SkipLoop).value = This.iOperationIndex
    Else
        Call parseFlowPriority3
    End If
End Sub

Private Sub parseFlowPriority3()
    If optConsume("for") Then

        ' 1. let index = LowerBound
        Dim varName As String: varName = consume("var")
        Call consume("equal")
        Call parseExpression
        Dim IncrementIndex As Long: IncrementIndex = findVariable(varName)
        If IncrementIndex >= 0 Then
            ' If the variable already existed, move the data to that pos on the stack
            Call addOperation(oSet, , IncrementIndex, -1)
        Else
            ' If the variable didn't exist yet, treat this stack pos as its source
            Call this.scopes(this.scopeCount).add(varName, this.stackSize)
            IncrementIndex = this.stackSize - 1
            If IncrementIndex = 0 Then IncrementIndex = 1 ' Avoid referencing the current operation. | Used when the last added variable was declared in for loop
        End If

        ' Index > UpperBound
        ' Jump if Condition true
        Call addOperation(oAccess, , IncrementIndex, 1)
        Dim RepeatLoop As Long: RepeatLoop = This.iOperationIndex - 1' beginning of loop (condition)
        If peek("to") Then
            This.iTokenIndex = This.iTokenIndex + 1
        Else
            Call parseBlock("to")
            Call consume("to")
        End If
        Call parseStatement
        Dim ConditionIndex As Long: ConditionIndex = addOperation(oComparison, oGt, , -1)
        Dim SkipLoop As Integer : SkipLoop = addOperation(oJump, ifTrue, , -1)  ' Skip loop

        ' Setup increment Values
        Dim Increment As Long, Operator As ISubType, Condition As ISubType
        If optConsume("step") Then
            If peek("literalNumber") Then
                Operator = oadd
                Condition = oGt
            Else
                Select Case True
                    Case optConsume("add")      : Operator = oadd      : Condition = oGt
                    Case optConsume("subtract") : Operator = osubtract : Condition = oLt
                    Case optConsume("multiply") : Operator = omultiply : Condition = oGt
                    Case optConsume("divide")   : Operator = odivide   : Condition = oLt
                    Case optConsume("power")    : Operator = opower    : Condition = oGt
                End Select
            End If
            Increment = CLng(consume("literalNumber")) ' TODO add more functionality than just numbers
        Else
            Increment = 1
            Operator  = oadd
            Condition = oGt
        End If

        This.Operations(ConditionIndex).subType = Condition ' Changes compare-operator of condition according to "step"

        ' Get rest of loop
        ' Setup index-incrementation: let index = index Operator Increment
        ' Setup loop-repeat 
        Call parseBlock("next")
        Call consume("next")
        Call addOperation(oAccess    ,          , IncrementIndex  , -1)
        Call addOperation(oPush      ,          , Increment       , -1)
        Call addOperation(oArithmetic, Operator ,                 , -1)
        Call addOperation(oSet       ,          , IncrementIndex  , -1)
        Call addOperation(oJump      ,          , RepeatLoop      , -1)
        Call consumeExit("for", This.iOperationIndex)
        This.Operations(SkipLoop).Value = This.iOperationIndex
    Else
        Call parseValuePriority1
    End If
End Sub

'Parse evaluation of numbers, arguments, strings, booleans, variable names and brackets. Will also parse accessors on these values.
'i.e. `varName.someMethod(1,2,3).someProp`
Private Sub parseValuePriority1()
    'Prefix unary operators for set/let
    Dim iOperationType as ISubType
    iOperationType = iif(optConsume("let"), oPropLet, oFieldCall)
    iOperationType = iif(optConsume("set"), oPropSet, iOperationType)

    If peek("literalNumber") Then
        Call addOperation(oPush, , CDbl(consume("literalNumber")), 1)
    ElseIf peek("arg") Then
        Call addOperation(oAccess, argument, val(mid(consume("arg"), 2)), 1)
        Call parseManyAccessors(iOperationType)
    ElseIf peek("literalString") Then
        Call parseString
    ElseIf peek("literalBoolean") Then
        Call addOperation(oPush, , consume("literalBoolean") = "true", 1)
    ElseIf peek("var") Then
        If Not parseScopeAccess Then
            Call parseFunction
        End If
        Call parseManyAccessors(iOperationType)
    ElseIf peek("step") Then ' does nothing, so that parseflowPriority3 does its thing
    ElseIf optConsume("exit") Then
        Call addExit(this.tokens(this.iTokenIndex).Type.name)
        this.iTokenIndex = this.iTokenIndex + 1
    Else
        Call consume("lBracket")
        Call parseExpression
        Call consume("rBracket")
        Call parseManyAccessors(iOperationType)
    End If
End Sub

'Parse a call to a standard in-built function and add it to the stack
'@remark - This is a special case of a function call. It is parsed differently because it is a built-in function and not 
'an in-code defined function.
Private Function parseFunction() As Variant
    Call addOperation(oPush, , consume("var"), 1)
    Dim size As Integer: size = this.stackSize
    Call parseOptParameters
    Call addOperation(oFunc)
    this.stackSize = size
End Function

'Parse object field, property and method accessors. Specifically allows for more than one accessor to be chained together
'i.e. `obj.someMethod(1,"hi").someProp`
Private Sub parseManyAccessors(Optional ByVal ISubType as ISubType = oFieldCall)
    Dim bLoop As Boolean: bLoop = True
    Do
        bLoop = False
        if parseOptObjectField(ISubType) then bLoop = True
        If parseOptObjectProperty(ISubType) Then bLoop = True
        If parseOptObjectMethod() Then bLoop = True
    Loop While bLoop
End Sub

'Parse an object's method or property call (i.e. `obj.someMethod` or `obj.someProp`) and add it to the operations stack
'@param ISubType - Whether this is a property get or property let/set
'@returns - Whether a method or property was found
Private Function parseOptObjectField(Optional ByVal ISubType as ISubType = oFieldCall) as Boolean
    parseOptObjectField = false
    if optConsume("fieldAccess") then
        Dim size As Integer: size = this.stackSize
        Call addOperation(oPush, , consume("var"), 1)
        Call parseOptParameters
        'Parse Let/Set ... = ... as a special case
        if peek("equal") and ISubType  <> oFieldCall then
            Call consume("equal")
            Call parseExpression
            Call addOperation(oObject, ISubType)
        else
            Call addOperation(oObject, oFieldCall)
        end if
        this.stackSize = size
        parseOptObjectField = True
    end if
End Function

'Parse an object's property access (i.e. Obj.someProp) and add it to the operations stack
'@param ISubType - Whether this is a property get or property let/set
'@returns - Whether a property was found
Private Function parseOptObjectProperty(Optional ByVal ISubType as ISubType = oPropGet) As Boolean
    parseOptObjectProperty = False
    If optConsume("propertyAccess") Then
        Dim size As Integer: size = this.stackSize
        Call addOperation(oPush, , consume("var"), 1)
        Call parseOptParameters
        'Parse Let/Set ... = ... as a special case
        if peek("equal") and ISubType  <> oPropGet then
            Call consume("equal")
            Call parseExpression
            Call addOperation(oObject, ISubType)
        else
            Call addOperation(oObject, oPropGet)
        end if
        this.stackSize = size
        parseOptObjectProperty = True
    End If
End Function

'Parse an object's method call (i.e. Obj.someMethod) and add it to the operations stack
'@returns - Whether a method was found
Private Function parseOptObjectMethod() As Boolean
    parseOptObjectMethod = False
    If optConsume("methodAccess") Then
        Dim size As Integer: size = this.stackSize
        Call addOperation(oPush, , consume("var"), 1)
        Call parseOptParameters
        Call addOperation(oObject, oMethodCall)
        this.stackSize = size
        parseOptObjectMethod = True
    End If
End Function

'Parse a function call's parameters and add them to the operations stack
'@returns - Whether parameters were found
'@remark Parsing parameters is "optional" in the sense that if no parameters are found, the function call will remain parsed.
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

'Parse a string literal and add it to the operations stack
Private Sub parseString()
    Dim sRes As String: sRes = consume("literalString")
    sRes = Mid(sRes, 2, Len(sRes) - 2)
    sRes = Replace(sRes, """""", """")
    Call addOperation(oPush, , sRes, 1)
End Sub

'Parse whether a function call or variable access is being made and add it to the operations stack
'@returns - Whether a named variable or function was found in the scope
Private Function parseScopeAccess() As Boolean
    If peek("lBracket", 2) Then
        parseScopeAccess = parseFunctionAccess()
    Else
        parseScopeAccess = parseVariableAccess()
    End If
End Function

'Parse the access of a named variable's value and add it to the operations stack
'@returns - Whether a variable was found in the scope
Private Function parseVariableAccess() As Boolean
    parseVariableAccess = False
    Dim varName As String: varName = consume("var")
    Dim offset As Long: offset = findVariable(varName)
    If offset >= 0 Then
        parseVariableAccess = True
        Call addOperation(oAccess, , 1 + offset, 1)
    Else
        this.iTokenIndex = this.iTokenIndex - 1 ' Revert token consumption
    End If
End Function

'Parse an assignment and add it to the current scope and operations stack
Private Sub parseAssignment()
    'Consume set/let keyword
    if not optConsume("let") then Call consume("set")
    
    Dim varName As String: varName = consume("var")
    Call consume("equal")
    Call parseExpression
    Dim offset As Long: offset = findVariable(varName)
    If offset >= 0 Then
        ' If the variable already existed, move the data to that pos on the stack
        Call addOperation(oSet, , offset, -1)
        'Call addOperation(oAccess, , offset, 1) ' To keep a return value
    Else
        ' If the variable didn't exist yet, treat this stack pos as its source
        Call this.scopes(this.scopeCount).add(varName, this.stackSize)
    End If
End Sub

'Find the position on the Operations stack of a variable by name
'@param varName - Name of variable to find
'@returns - The position of the variable on the Operations stack
Private Function findVariable(varName As String) As Long
    Dim scope As Long: scope = this.scopeCount
    findVariable = -1
    While scope > 0
        If this.scopes(scope).exists(varName) Then
            If scope < this.funcScope Then
                Call Throw("Can't access """ & varName & """, functions can unfortunately not access data outside their block")
            ElseIf this.scopesArgCount(scope).exists(varName) Then
                Call Throw("Expected a variable, but found a function for name " & varName)
            Else
                findVariable = this.stackSize - this.scopes(scope).item(varName)
                scope = 0
            End If
        End If
        scope = scope - 1
    Wend
End Function

'Parse a named scoped in-code defined function call and add it to the current scope and operations stack
'@returns - Whether a function was found
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
            Call Throw(argCount & " arguments should have been provided to " & funcName & " but only " & iArgCount & " were received")
        End If
        
        ' Add call and return data
        Call addOperation(oJump, , funcPos, -iArgCount) 'only -argCount since pushing Result and popping return pos cancel out
        this.operations(returnPosIndex).value = this.iOperationIndex
        Call consumeExit("fun", This.iOperationIndex)
    Else
        this.iTokenIndex = this.iTokenIndex - 1 ' Revert token consumption
    End If
End Function

'Parse a function declaration and add it to the current scope and operations stack
Private Sub parseFunctionDeclaration()
    ' Create a dedicated scope for this funcion
    Call addScope
    Dim prevFuncScope As Long: prevFuncScope = this.funcScope
    this.funcScope = this.scopeCount
    
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
    Call this.scopes(this.scopeCount - 1).add(funcName, this.iOperationIndex)
    Call this.scopesArgCount(this.scopeCount - 1).add(funcName, iArgCount)
    
    ' Obtain the body
    Call parseBlock("end")
    Call consume("end")
    While iArgCount > 0
        Call addOperation(oMerge, , , -1)
        iArgCount = iArgCount - 1
    Wend
    Call addOperation(oReturn, withValue, , -1)
    this.operations(skipToIndex).value = this.iOperationIndex
    
    ' Reset the scope
    this.scopeCount = this.scopeCount - 1
    this.funcScope = prevFuncScope
End Sub

'Parse a parameter declaration and add it to the current scope
Private Sub parseParameterDeclaration()
    Dim varName As String: varName = consume("var")
    Dim offset As Long: offset = findVariable(varName)
    If offset >= 0 Then
        Call Throw("You can't declare multiple parameters with the same name")
    Else
        ' Reserve a spot for this parameter, it will be pushed by the caller
        this.stackSize = this.stackSize + 1
        Call this.scopes(this.scopeCount).add(varName, this.stackSize)
    End If
End Sub

'Find function by name
'@param varName as string - Name of function to find
'@param argCount as long - Number of arguments the function takes (out)
'@returns - The position of the function in the Operations array
Private Function findFunction(varName As String, Optional ByRef argCount As Long) As Long
    Dim scope As Long: scope = this.scopeCount
    findFunction = -1
    While scope > 0
        If this.scopes(scope).exists(varName) Then
            If Not this.scopesArgCount(scope).exists(varName) Then
                Call Throw("Expected a function, but found a variable for name " & varName)
            Else
                findFunction = this.scopes(scope).item(varName)
                argCount = this.scopesArgCount(scope).item(varName)
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
'@param ops as Array<Operation> - operations  The operations to evaluate
'@param vLastArgs as Array<Variant> - The last arguments passed to the function
'@returns - The result of the operations
Private Function evaluate(ByRef ops() As Operation, ByVal vLastArgs As Variant) As Variant
    Dim stack() As Variant
    ReDim stack(0 To 5)
    Dim stackPtr As Long: stackPtr = 0
    
    Dim op As Operation
    Dim v1 As Variant
    Dim v2 As Variant
    Dim v3 As Variant
    Dim opIndex As Long: opIndex = 0
    Dim opCount As Long: opCount = UBound(ops)
    
    'If result is in performance cache then return it immediately
    if this.UsePerformanceCache then
        Dim sPerformanceCacheID as string: sPerformanceCacheID = getPerformanceCacheID(vLastArgs)     
        if this.PerformanceCache.exists(sPerformanceCacheID) then
            Call CopyVariant(evaluate, this.PerformanceCache(sPerformanceCacheID))
            Exit Function
        end if
    end if

    'Evaluate operations to identify result
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
                Call CopyVariant(v2, popV(stack, stackPtr))
                Call CopyVariant(v1, popV(stack, stackPtr))
                if vartype(v1) = vbError or vartype(v2) = vbError then
                    v3 = false
                else
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
                        Case ISubType.oLike
                            v3 = v1 Like v2
                        Case ISubType.oIs
                            v3 = v1 is v2
                        Case Else
                            v3 = Empty
                    End Select
                end if
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
                Call pushV(stack, stackPtr, evaluateFunc(v1, args))
            'Misc
            Case iType.oMisc
                v2 = popV(stack, stackPtr)
                v1 = popV(stack, stackPtr)
                Select Case op.subType
                    Case ISubType.oCat
                        v3 = v1 & v2
                    Case Else
                        v3 = Empty
                End Select
                Call pushV(stack, stackPtr, v3)
            'Variable
            Case iType.oAccess
                Select Case op.subType
                    Case ISubType.argument
                        Dim iArgIndex As Long: iArgIndex = op.value + LBound(vLastArgs) - 1
                        If iArgIndex <= UBound(vLastArgs) Then
                            Call pushV(stack, stackPtr, vLastArgs(iArgIndex))
                        Else
                            Call Throw("Argument " & iArgIndex & " not supplied to Lambda.")
                        End If
                    Case Else
                        'HACK: bug where if stack(stackPtr-op.value) was an object, then array will become locked. Array locking occurs by compiler to try to protect	
                        'instances when re-allocation would move the array, and thus corrupt the pointers. By copying the variant we redivert the compiler's efforts,	
                        'but we might actually open ourselves to errors... @issue	
                        Dim vAccessVar: Call CopyVariant(vAccessVar, stack(stackPtr - op.value))	
                        Call pushV(stack, stackPtr, vAccessVar)
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
                        Call CopyVariant(v1, popV(stack, stackPtr))
                        opIndex = stack(stackPtr - 1)
                        Call CopyVariant(stack(stackPtr - 1), v1)
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
                'End loop - This occurs when opIndex > opCount 
                opIndex = opCount+1
        End Select
    Wend

    'Add result to performance cache
    if this.UsePerformanceCache then
        if isObject(stack(0)) then
            set this.PerformanceCache(sPerformanceCacheID) = stack(0)
        else
            let this.PerformanceCache(sPerformanceCacheID) = stack(0)
        end if
    end if

    Call CopyVariant(evaluate, stack(0))
End Function

'Retrieves the arguments from the stack
'@param stack    - The stack to get the data from and add the result to
'@param stackPtr - The pointer that indicates the position of the top of the stack
'@returns Variant<Array<Variant>> - The args list
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
        
        'Arguments are held on the stack in order, which means that we need to fill the array in reverse order.
        For i = argCount To 1 Step -1
            Call CopyVariant(args(i), popV(stack, stackPtr))
        Next
    End If
    
    getArgs = args
End Function

'Calls an object method/setter/getter/letter
'@param stack    - The stack to get the data from and add the result to
'@param stackPtr - The pointer that indicates the position of the top of the stack
'@param op       - The operation to execute
Private Sub objectCaller(ByRef stack() As Variant, ByRef stackPtr As Long, ByRef op As Operation)
    'Get the name and arguments
    Dim bIsSetter as boolean: bIsSetter = op.subType = ISubType.oPropSet or op.subType = ISubType.oPropLet
    Dim iOriginalPos as long: iOriginalPos = stackPtr 
    stackPtr = stackPtr - iif(bIsSetter, 1, 0)
    Dim args() As Variant: args = getArgs(stack, stackPtr)
    if bIsSetter then
        Redim Preserve args(0 To UBound(args) + 1)
        Call CopyVariant(args(UBound(args)), stack(iOriginalPos - 1))
        #if devMode then
            stack(iOriginalPos - 1) = empty
        #end if
    end if

    Dim funcName As Variant: funcName = popV(stack, stackPtr)
    
    'Get caller type
    Dim callerType As VbCallType
    Select Case op.subType
        Case ISubType.oFieldCall:   callerType = vbGetOrMethod
        Case ISubType.oPropGet:     callerType = VbGet
        Case ISubType.oMethodCall:  callerType = VbMethod
        Case ISubType.oPropLet:     callerType = VbLet
        Case ISubType.oPropSet:     callerType = VbSet
    End Select
                
    'Call rtcCallByName
    Dim obj As Object
    Set obj = popV(stack, stackPtr)
    if bIsSetter then
        Call stdCallByName(obj, funcName, callerType, args)
        Call pushV(stack, stackPtr, args(UBound(args))) 'Return the value that was set
    else
        Call pushV(stack, stackPtr, stdCallByName(obj, funcName, callerType, args))
    end if
End Sub

'Calls an object method/setter/getter/letter. Treats dictionary properties as direct object properties, I.E. `A.B` ==> `A.item("B")`
'@param obj - The object to call
'@param funcName - The method name to call
'@param callerType - The property/method call type
'@param args  - An array of arguments. This function supports up to 30 arguments, akin to Application.Run
'@returns  - The return value of the called function
Private Function stdCallByName(ByRef obj As Object, ByVal funcName As String, ByVal callerType As VbCallType, ByRef args() As Variant) As Variant
    'If Dictionary and
    If TypeName(obj) = "Dictionary" Then
        Select Case funcName
            Case "add", "exists", "items", "keys", "remove", "removeall", "comparemode", "count", "item"
                'These methods already exist on dictionary, do not override
            Case Else
                Select Case callerType
                  Case vbGetOrMethod
                    'Call DictionaryInstance.Item(funcName) only if funcName exists on the item
                    If obj.exists(funcName) Then Call CopyVariant(stdCallByName, obj.item(funcName))
                  Case VbSet
                    Set obj(funcName) = args(0)
                  Case VbLet
                    Let obj(funcName) = args(0)
                End Select
                Exit Function
        End Select
    End If

    'Call CallByName from DLL or
    #If Mac Then
        Call CopyVariant(stdCallByName, macCallByName(obj, funcName, callerType, args))
    #Else
        'TODO: Better error handling (property or method <funcName> doesn't exist on object with type <typename(obj)>)
        On Error GoTo ErrorInRTCCallByName
        Call CopyVariant(stdCallByName, rtcCallByName(obj, StrPtr(funcName), callerType, args, &H409))
    #End If
    Exit Function
ErrorInRTCCallByName:
    'HACK: Rarely objects which are poorly implemented might throw missing member errors. This can for instance happen with RecordSet. This try loop is an attempt to catch this
    Dim iTry as Long: iTry = iTry + 1
    if iTry < 5 then
        Dim sCallerTypeName as string
        select case callerType
            case vbGetOrMethod: sCallerTypeName = "Property or Method "
            case VbGet: sCallerTypeName = "Property "
            Case VbMethod: sCallerTypeName = "Method "
        end select
        Err.Raise Err.Number, "", sCallerTypeName & funcName & " doesn't exist on object with type " & typename(obj)
        Resume
    else
        Resume
    end if
End Function

'Evaluates the built in standard functions
'@param {String} sFuncName The name of the function to invoke
'@param {Variant} args() The arguments
'@returns The result
Private Function evaluateFunc(ByVal sFuncName As String, ByVal args As Variant) As Variant
    Dim iArgStart As Long: iArgStart = LBound(args)
    If TypeName(this.FunctionExtensions) = "Dictionary" Then
        If this.FunctionExtensions.exists(sFuncName) Then
            Dim vInjectedVar As Variant
            Call CopyVariant(vInjectedVar, this.FunctionExtensions(sFuncName))
            If TypeOf vInjectedVar Is stdICallable Then
                Call CopyVariant(evaluateFunc, this.FunctionExtensions(sFuncName).RunEx(args))
            Else
                Call CopyVariant(evaluateFunc, vInjectedVar)
            End If
            Exit Function
        End If
    End If
    
    Select Case LCase(sFuncName)
        'Useful OOP constants
        Case "thisworkbook": if isObject(ThisWorkbook) then Set evaluateFunc = ThisWorkbook
        Case "application":  if isObject(Application)  then  Set evaluateFunc = Application

        'Data structures
        Case "dict":
            Dim oRetDict as object: set oRetDict = CreateObject("Scripting.Dictionary")
            For i = iArgStart To UBound(args) Step 2
                Call oRetDict.add(args(i), args(i+1))
            next
            set evaluateFunc = oRetDict

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
        Case "null":            evaluateFunc = Null
        Case "nothing":         set evaluateFunc = Nothing
        Case "empty":           evaluateFunc = Empty
        Case "missing":         evaluateFunc = getMissing()

        'VBA Structure
        Case "array": evaluateFunc = args
        'TODO: Case "callbyname": evaluateFunc = CallByName(args(iArgStart))
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
        Case "typename"
            evaluateFunc = TypeName(args(iArgStart))
            
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
        Case "right": evaluateFunc = VBA.right(args(iArgStart), args(iArgStart + 1))
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
        Case "switch"
            'TODO: Switch caching and use of dictionary would be good here
            For i = iArgStart + 1 To UBound(args) Step 2
              If i + 1 > UBound(args) Then
                Call CopyVariant(evaluateFunc, args(i))
                Exit For
              Else
                If IsObject(args(iArgStart)) And IsObject(args(i)) Then
                  If args(iArgStart) Is args(i) Then
                    Set evaluateFunc = args(i + 1)
                    Exit For
                  End If
                ElseIf (Not IsObject(args(iArgStart))) And (Not IsObject(args(i))) Then
                  If args(iArgStart) = args(i) Then
                    evaluateFunc = args(i + 1)
                    Exit For
                  End If
                End If
              End If
            Next
        Case "any"
            evaluateFunc = False
            'Detect if comparee is an object or a value
            If IsObject(args(iArgStart)) Then
                For i = iArgStart + 1 To UBound(args)
                  If IsObject(args(i)) Then
                    If args(iArgStart) Is args(i) Then
                      evaluateFunc = True
                      Exit For
                    End If
                  End If
                Next
            Else
                For i = iArgStart + 1 To UBound(args)
                  If Not IsObject(args(i)) Then
                    If args(iArgStart) = args(i) Then
                      evaluateFunc = True
                      Exit For
                    End If
                  End If
                Next
            End If
        Case "eval": evaluateFunc = stdLambda.Create(args(iArgStart)).Run()
        Case "lambda": 
            set evaluateFunc = stdLambda.Create(args(iArgStart))
            set evaluateFunc.oFuncExt = oFuncExt
        Case "isnumeric": evaluateFunc = isNumeric(args(iArgStart))
        Case "isobject": evaluateFunc = IsObject(args(iArgStart))
        Case "string": evaluateFunc = String(args(iArgStart), args(iArgStart + 1))
        Case "space": evaluateFunc = String(args(iArgStart), " ")
        Case Else
            Call Throw("No such function: " & sFuncName)
    End Select
End Function

'==========================
'
'  General helper methods
'
'==========================

'Class initialisation
Private Sub Class_Initialize()
    'If this is stdLambda predeclared class, ensure that oFuncExt is defined.
    If Me Is stdLambda Then Set oFuncExt = CreateObject("Scripting.Dictionary")
End Sub

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

'Obtains a TokenDefinition from input params
'@param {ByVal String} The name of the token
'@param {ByVal String} The regex pattern to match durin tokenisation
'@param {ByVal Boolean?=True} Should this token ignoreCase?
'@param {ByVal Boolean?=False} Is this token a keyword?
'@returns {TokenDefinition} The definition of the token
Private Function getTokenDefinition(ByVal sName As String, ByVal sRegex As String, Optional ByVal ignoreCase As Boolean = True, Optional ByVal isKeyword As Boolean = False) As TokenDefinition
    getTokenDefinition.name = sName
    getTokenDefinition.Regex = sRegex & IIf(isKeyword, "\b", "")
    Set getTokenDefinition.RegexObj = CreateObject("VBScript.Regexp")
    getTokenDefinition.RegexObj.pattern = "^(?:" & sRegex & IIf(isKeyword, "\b", "") & ")"
    getTokenDefinition.RegexObj.ignoreCase = ignoreCase
End Function

'Used to get the definition of a token by name for inline-insertion of tokens
'@param {ByVal TokenDefinition()} The name of the token
'@param {ByVal String} The name of the token
'@returns {TokenDefinition} The definition of the token
Private Function getTokenDefinitionByName(ByRef tokenDefs() As TokenDefinition, ByVal tokenName As String) As TokenDefinition
    Dim i As Long
    For i = LBound(tokenDefs) To Ubound(tokenDefs)
        If tokenDefs(i).Name = tokenName Then
            getTokenDefinitionByName = tokenDefs(i)
            Exit Function
        End If
    Next i
End Function

'Copies one variant to a destination
'@param {ByRef token()} tokens Tokens to remove the specified type from
'@param {string} sRemoveType   Token type to remove.
'@returns {token()} The modified token array.
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
    If this.iTokenIndex = 0 Then this.iTokenIndex = 1
    
    'Get next token
    ShiftTokens = this.tokens(this.iTokenIndex)
    
    'Increment token index
    this.iTokenIndex = this.iTokenIndex + 1
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
    If this.iTokenIndex = 0 Then this.iTokenIndex = 1
    If this.iTokenIndex + offset - 1 <= UBound(this.tokens) Then
        peek = this.tokens(this.iTokenIndex + offset - 1).Type.name = sTokenType
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
    If this.iOperationIndex = 0 Then
        ReDim Preserve this.operations(0 To 1)
    Else
        Dim size As Long: size = UBound(this.operations)
        If this.iOperationIndex > size Then
            ReDim Preserve this.operations(0 To size * 2)
        End If
    End If
    
    With this.operations(this.iOperationIndex)
        .Type = kType
        .subType = subType
        Call CopyVariant(.value, value)
    End With
    addOperation = this.iOperationIndex
    this.stackSize = this.stackSize + stackDelta
    
    this.iOperationIndex = this.iOperationIndex + 1
End Function

'Resizes the operations list so there are no more empty operations
Private Sub finishOperations()
    ReDim Preserve this.operations(0 To this.iOperationIndex)
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

'Serializes the argument array passed to a string.
'@param {ByRef Variant()} Arguments to serialize
'@returns {String} Serialized representation of the arguments.
'@remark Objects cannot be split into their components and thus are cached as a conglomerate of type and pointer (e.g. Dictionary<12341234123>). 
'@TODO: Potentially use [StgSerializePropVariant ](https://docs.microsoft.com/en-us/windows/win32/api/propvarutil/nf-propvarutil-stgserializepropvariant) as this'd be more optimal
'@example
'  Debug.Print getPerformanceCacheID(Array())=""
'  Debug.Print getPerformanceCacheID(Array(Array(1, 2, Null), "yop", Empty, "", Nothing, New Collection, DateSerial(2020, 1, 1), False, True)) = "Array[1;2;null;];""yop"";empty;"""";Nothing;Collection<1720260481920>;01/01/2020;False;True;"
'returns
'  True
'  True
Private Function getPerformanceCacheID(ByRef Arguments As Variant) As String
    Dim length As Long: length = UBound(Arguments) - LBound(Arguments) + 1
    If length > 0 Then
        Dim sSerialized As String: sSerialized = ""
        For i = LBound(Arguments) To UBound(Arguments)
            Select Case VarType(Arguments(i))
              Case vbBoolean, vbByte, vbInteger, vbLong, vbLongLong, vbCurrency, vbDate, vbDecimal, vbDouble, vbSingle
                  sSerialized = sSerialized & Arguments(i) & ";"
              Case vbString
                  sSerialized = sSerialized & """" & Arguments(i) & """;"
              Case vbObject, vbDataObject
                  If Arguments(i) Is Nothing Then
                      sSerialized = sSerialized & "Nothing;"
                  Else
                      sSerialized = sSerialized & TypeName(Arguments(i)) & "<" & ObjPtr(Arguments(i)) & ">;"
                  End If
              Case vbEmpty
                  sSerialized = sSerialized & "empty;"
              Case vbNull
                  sSerialized = sSerialized & "null;"
              Case vbError
                  sSerialized = sSerialized & "error;"
              Case Else
                  If CBool(VarType(Arguments(i)) And vbArray) Then
                      sSerialized = sSerialized & "Array[" & getPerformanceCacheID(Arguments(i)) & "];"
                  Else
                      sSerialized = sSerialized & "Unknown;"
                  End If
            End Select
        Next
    End If
    getPerformanceCacheID = sSerialized
End Function

'-------
'general
'-------

'Used to obtain missing
'@Param {Variant} The value to be returned - Please do not populate this parameter.
'@returns {Missing} Missing value
Private Function GetMissing(Optional arg As Variant) As Variant
  GetMissing = arg
End Function

'Copies one variant to a destination
'@param {ByRef Variant} dest Destination to copy variant to
'@param {Variant} value Source to copy variant from.
'@perf This appears to be a faster variant of "oleaut32.dll\VariantCopy" + it's multi-platform
Private Sub CopyVariant(ByRef dest As Variant, ByVal value As Variant)
  If IsObject(value) Then
    Set dest = value
  Else
    Let dest = value
  End If
End Sub


'TODO: Better error handling
'Throws an error
'@param {string} The error message to be thrown
'@returns {void}
Private Sub Throw(ByVal sMessage As String)
    Err.Raise 1, "stdLambda", sMessage, vbCritical
    End
End Sub


'Used by Bind() for binding arguments ontop of BoundArgs and binding bound args to passed arguments
'@param {Variant()} The 1st array which will
'@param {Variant()} The 2nd array which will be concatenated after the 1st
'@complexity O(1)
Private Function ConcatArrays(ByVal Arr1 As Variant, ByVal Arr2 As Variant) As Variant
    Dim ub1 As Long: ub1 = UBound(Arr1)
    Dim lb1 As Long: lb1 = LBound(Arr1)
    Dim ub2 As Long: ub2 = UBound(Arr2)
    Dim lb2 As Long: lb2 = LBound(Arr2)
    Dim iub As Long: iub = ub1 + ub2 - lb2 + 1
    
    If iub > -1 Then
        Dim v() As Variant
        ReDim v(lb1 To iub)
        
        
        Dim i As Long
        For i = LBound(v) To UBound(v)
            If i <= ub1 Then
                Call CopyVariant(v(i), Arr1(i))
            Else
                Call CopyVariant(v(i), Arr2(i - ub1 - 1 + lb2))
            End If
        Next
        ConcatArrays = v
    Else
        ConcatArrays = Array()
    End If
End Function

'----------
'Handle exit-token
'----------

'pushes a value onto the exitstack and creates a jump operation
'@param {String}  for what should be exited e.g. for, if, do, function
Private Sub addExit(ExitWhat As String)
    With this
        .ExitStackSize = .ExitStackSize + 1
        ReDim Preserve .ExitStackName(1 To .ExitStackSize) 
        ReDim Preserve .ExitStackIndex(1 To .ExitStackSize)
        .ExitStackName(.ExitStackSize) = ExitWhat
        .ExitStackIndex(.ExitStackSize) = addOperation(oJump, , , -1)
    End With
End Sub

'If Caller is equal to stackname of exit then it will get the jump-condition index of operations and give it a value
' This is important, as something like this is possible:
' let x = 0                
' for i = 1 to 10          
'   if $1 = 10 then        
'       exit for-----------|
'   else                   |
'       {---}              |
'       do                 |
'           if x = 10 then |
'               exit do----|----|
'           end            |    |
'           x = x + 1      |    |
'       loop               |    |
'    end-------------------|---/
' next                     |
' {---} <-----------------/

' The code might be several layers deep, to accomodate for that this exitstack was created
'@param {String} Name of the current calling abstraction e.g for, if, do, function
'@param {Long} Index where to jump to once exit is run
Private Sub consumeExit(Caller As String, JumpIndex As Long)
    Dim i As Long
    Dim TempName() As String
    Dim TempIndex() As Long
    With this
        If .ExitStackSize = 0 Then Exit Sub
        If .ExitStackName(.ExitStackSize) = Caller Then
            .Operations(.ExitStackIndex(.ExitStackSize)).Value = JumpIndex
            .ExitStackSize = ExitStackSize - 1
            If .ExitStackSize > 0 Then
                ReDim TempName(.ExitStackSize)
                ReDim TempIndex(.ExitStackSize)
                For i = 1 To .ExitStackSize
                    TempName(i) = this.ExitStackName(i)
                    TempIndex(i) = this.ExitStackIndex(i)
                Next i
            End If
            .ExitStackName = TempName
            .ExitStackIndex = TempIndex
        End If
    End With
End Sub

'----------
'evaluation Mac
'----------

'Reimplementation of rtcCallByName() but for Mac OS
'@param {ByRef Object} - The object to call
'@param {ByVal String} - The method name to call
'@param {ByVal VbCallType} - The property/method call type
'@param {ByVal Variant()}  - An array of arguments. This function supports up to 30 arguments, akin to Application.Run
'@returns Variant - The return value of the called function
Private Function macCallByName(ByRef obj As Object, ByVal funcName As String, ByVal callerType As VbCallType, ByVal args As Variant) As Variant
    'Get currentLength
    Dim currentLength As Integer: currentLength = UBound(args) - LBound(args) + 1
    Dim i As Long: i = LBound(args)
    
    'Cant use same trick as in stdCallback, as it seems CallByName doesn't support the Missing value... So have to do it this way...
    'Will go up to 30 as per Application.Run() Also seems that you can't pass args array directly to CallByName() because it causes an Overflow error,
    'instead we need to convert the args to vars first... Yes this doesn't look at all pretty, but at least it's compartmentalised to the end of the code...
    Dim a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18, a19, a20, a21, a22, a23, a24, a25, a26, a27, a28, a29
    If currentLength - 1 >= 0 Then Call CopyVariant(a0, args(i + 0)) Else GoTo macJmpCall
    If currentLength - 1 >= 1 Then Call CopyVariant(a1, args(i + 1)) Else GoTo macJmpCall
    If currentLength - 1 >= 2 Then Call CopyVariant(a2, args(i + 2)) Else GoTo macJmpCall
    If currentLength - 1 >= 3 Then Call CopyVariant(a3, args(i + 3)) Else GoTo macJmpCall
    If currentLength - 1 >= 4 Then Call CopyVariant(a4, args(i + 4)) Else GoTo macJmpCall
    If currentLength - 1 >= 5 Then Call CopyVariant(a5, args(i + 5)) Else GoTo macJmpCall
    If currentLength - 1 >= 6 Then Call CopyVariant(a6, args(i + 6)) Else GoTo macJmpCall
    If currentLength - 1 >= 7 Then Call CopyVariant(a7, args(i + 7)) Else GoTo macJmpCall
    If currentLength - 1 >= 8 Then Call CopyVariant(a8, args(i + 8)) Else GoTo macJmpCall
    If currentLength - 1 >= 9 Then Call CopyVariant(a9, args(i + 9)) Else GoTo macJmpCall
    If currentLength - 1 >= 10 Then Call CopyVariant(a10, args(i + 10)) Else GoTo macJmpCall
    If currentLength - 1 >= 11 Then Call CopyVariant(a11, args(i + 11)) Else GoTo macJmpCall
    If currentLength - 1 >= 12 Then Call CopyVariant(a12, args(i + 12)) Else GoTo macJmpCall
    If currentLength - 1 >= 13 Then Call CopyVariant(a13, args(i + 13)) Else GoTo macJmpCall
    If currentLength - 1 >= 14 Then Call CopyVariant(a14, args(i + 14)) Else GoTo macJmpCall
    If currentLength - 1 >= 15 Then Call CopyVariant(a15, args(i + 15)) Else GoTo macJmpCall
    If currentLength - 1 >= 16 Then Call CopyVariant(a16, args(i + 16)) Else GoTo macJmpCall
    If currentLength - 1 >= 17 Then Call CopyVariant(a17, args(i + 17)) Else GoTo macJmpCall
    If currentLength - 1 >= 18 Then Call CopyVariant(a18, args(i + 18)) Else GoTo macJmpCall
    If currentLength - 1 >= 19 Then Call CopyVariant(a19, args(i + 19)) Else GoTo macJmpCall
    If currentLength - 1 >= 20 Then Call CopyVariant(a20, args(i + 20)) Else GoTo macJmpCall
    If currentLength - 1 >= 21 Then Call CopyVariant(a21, args(i + 21)) Else GoTo macJmpCall
    If currentLength - 1 >= 22 Then Call CopyVariant(a22, args(i + 22)) Else GoTo macJmpCall
    If currentLength - 1 >= 23 Then Call CopyVariant(a23, args(i + 23)) Else GoTo macJmpCall
    If currentLength - 1 >= 24 Then Call CopyVariant(a24, args(i + 24)) Else GoTo macJmpCall
    If currentLength - 1 >= 25 Then Call CopyVariant(a25, args(i + 25)) Else GoTo macJmpCall
    If currentLength - 1 >= 26 Then Call CopyVariant(a26, args(i + 26)) Else GoTo macJmpCall
    If currentLength - 1 >= 27 Then Call CopyVariant(a27, args(i + 27)) Else GoTo macJmpCall
    If currentLength - 1 >= 28 Then Call CopyVariant(a28, args(i + 28)) Else GoTo macJmpCall
    If currentLength - 1 >= 29 Then Call CopyVariant(a29, args(i + 29)) Else GoTo macJmpCall

macJmpCall:
    Select Case currentLength
        Case 0:  Call CopyVariant(macCallByName, CallByName(obj, funcName, callerType))
        Case 1:  Call CopyVariant(macCallByName, CallByName(obj, funcName, callerType, a0))
        Case 2:  Call CopyVariant(macCallByName, CallByName(obj, funcName, callerType, a0, a1))
        Case 3:  Call CopyVariant(macCallByName, CallByName(obj, funcName, callerType, a0, a1, a2))
        Case 4:  Call CopyVariant(macCallByName, CallByName(obj, funcName, callerType, a0, a1, a2, a3))
        Case 5:  Call CopyVariant(macCallByName, CallByName(obj, funcName, callerType, a0, a1, a2, a3, a4))
        Case 6:  Call CopyVariant(macCallByName, CallByName(obj, funcName, callerType, a0, a1, a2, a3, a4, a5))
        Case 7:  Call CopyVariant(macCallByName, CallByName(obj, funcName, callerType, a0, a1, a2, a3, a4, a5, a6))
        Case 8:  Call CopyVariant(macCallByName, CallByName(obj, funcName, callerType, a0, a1, a2, a3, a4, a5, a6, a7))
        Case 9:  Call CopyVariant(macCallByName, CallByName(obj, funcName, callerType, a0, a1, a2, a3, a4, a5, a6, a7, a8))
        Case 10: Call CopyVariant(macCallByName, CallByName(obj, funcName, callerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9))
        Case 11: Call CopyVariant(macCallByName, CallByName(obj, funcName, callerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10))
        Case 12: Call CopyVariant(macCallByName, CallByName(obj, funcName, callerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11))
        Case 13: Call CopyVariant(macCallByName, CallByName(obj, funcName, callerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12))
        Case 14: Call CopyVariant(macCallByName, CallByName(obj, funcName, callerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13))
        Case 15: Call CopyVariant(macCallByName, CallByName(obj, funcName, callerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14))
        Case 16: Call CopyVariant(macCallByName, CallByName(obj, funcName, callerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15))
        Case 17: Call CopyVariant(macCallByName, CallByName(obj, funcName, callerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16))
        Case 18: Call CopyVariant(macCallByName, CallByName(obj, funcName, callerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17))
        Case 19: Call CopyVariant(macCallByName, CallByName(obj, funcName, callerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18))
        Case 20: Call CopyVariant(macCallByName, CallByName(obj, funcName, callerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18, a19))
        Case 21: Call CopyVariant(macCallByName, CallByName(obj, funcName, callerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18, a19, a20))
        Case 22: Call CopyVariant(macCallByName, CallByName(obj, funcName, callerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18, a19, a20, a21))
        Case 23: Call CopyVariant(macCallByName, CallByName(obj, funcName, callerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18, a19, a20, a21, a22))
        Case 24: Call CopyVariant(macCallByName, CallByName(obj, funcName, callerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18, a19, a20, a21, a22, a23))
        Case 25: Call CopyVariant(macCallByName, CallByName(obj, funcName, callerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18, a19, a20, a21, a22, a23, a24))
        Case 26: Call CopyVariant(macCallByName, CallByName(obj, funcName, callerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18, a19, a20, a21, a22, a23, a24, a25))
        Case 27: Call CopyVariant(macCallByName, CallByName(obj, funcName, callerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18, a19, a20, a21, a22, a23, a24, a25, a26))
        Case 28: Call CopyVariant(macCallByName, CallByName(obj, funcName, callerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18, a19, a20, a21, a22, a23, a24, a25, a26, a27))
        Case 29: Call CopyVariant(macCallByName, CallByName(obj, funcName, callerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18, a19, a20, a21, a22, a23, a24, a25, a26, a27, a28))
        Case 30: Call CopyVariant(macCallByName, CallByName(obj, funcName, callerType, a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, a17, a18, a19, a20, a21, a22, a23, a24, a25, a26, a27, a28, a29))
    End Select
End Function