VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdLambda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@module
'@description A class which implements a syntax for creating lambda functions from string expressions
'within VBA. Use this class to create sandboxed lambda functions, and call them with arguments.
'
'Below shows a trivial use case of `stdLambda` and `stdArray`. 
'
'```vb
'Dim arr as stdArray: Set arr = stdArray.Create(1,2,3,4,5,6,7,8,9,10)
''Filter to include only even numbers:
'set arr = arr.filter(stdLambda.Create("$1 mod 2 = 0"))  '=> Array(2,4,6,8,10)
''Take every element and multiply it by 10
'set arr = arr.map(stdLambda.Create("$1*10"))            '=> Array(20,40,60,80,100)
''And calculate the sum:
'set arr = arr.reduce(stdLambda.Create("$1+$2"), 0)      '=> 300
'```
'
'As you can see, within 3 lines of code (excluding the comments), we have achieved the equivalent of this code: 
'
'```vb
'Dim element as variant, totalSum as Long
'For each element in Array(1,2,3,4,5,6,7,8,9,10)
'  if element mod 2 = 0 then
'    totalSum = totalSum + element * 10
'  end if
'next
'```
'
'This is a very powerful feature of the `stdLambda` class, and one of the main reasons why it was created.
'It allows you to write code in a more natural way, and achieve the same result in a much shorter amount of
'code and improving maintainability. 
'
'For a more detailed breakdown of syntax and use cases see the [stdLambda guide](https://github.com/sancarn/stdVBA/blob/master/docs/stdLambda.md)
'@remark This module implements `stdICallable` meaning it can be utilised by many other stdVBA classes.
'@example This example shows how to create a simple lambda function which returns the sum of two numbers.
'```vb
'Dim cbSum As stdLambda: Set cbSum = stdLambda.Create("$1+$2")
'Debug.Print cbSum.Run(1, 2)   '=> 3
'Debug.print cbSum.Run(3,-3)   '=> 0
'```
'@example Find a window by it's title.
'```vb
'Dim desktop as stdWindow: set desktop = stdWindow.CreateFromDesktop()
'Dim win as stdWindow: set win = desktop.FindFirst(stdLambda.Create("$1.Caption = ""Title to find"""))
'```


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
    Instruction As IInstruction
    value As Variant
End Type

'Evaluation operation types
Private Enum IInstruction
    ' Data Instructions
    iPush = 1
    iPop
    iMerge

    ' Access Instructions
    iAccess_General
    iAccess_Argument

    ' Set Instructions, not to be confused with Object#PropSet
    iSet_General

    ' Arithmetic Instructions
    iArithmetic_Add
    iArithmetic_Subtract
    iArithmetic_Multiply
    iArithmetic_Divide
    iArithmetic_Power
    iArithmetic_Modulo
    iArithmetic_Negate

    ' Logical Instructions
    iLogic_And
    iLogic_Or
    iLogic_Not
    iLogic_Xor

    ' Function Instructions
    iFunc_Call

    ' Comparison Instructions
    iComparison_Equal
    iComparison_NotEqual
    iComparison_GreaterThan
    iComparison_GreaterThanOrEqual
    iComparison_LessThan
    iComparison_LessThanOrEqual
    iComparison_Like
    iComparison_Is

    ' Miscellaneous Instructions
    iMisc_Concatenate

    ' Jump Instructions
    iJump_Unconditional
    iJump_IfTrue
    iJump_IfFalse

    ' Return Instructions
    iReturn_WithValue
    iReturn_NoValue     'Implemented in VM, but not in parser yet

    'Object Instructions
    iObject_PropGet
    iObject_PropLet
    iObject_PropSet
    iObject_MethodCall
    iObject_FieldCall    'PropGet OR MethodCall
End Enum

Private Enum LambdaType
    iStandardLambda = 1
    iBoundLambda = 2
End Enum

'Special constant used in parsing:
Private Const UniqueConst As String = "3207af79-30df-4890-ade1-640f9f28f309"
Private Const vbGetOrMethod As Long = VbGet Or VbMethod
Private Const minStackSize As Long = 30 'note that the stack size may become smaller than this

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
    Equation As String

    'Instance var on both class and object
    FunctionExtensions As Object
    isBoundLambda As Boolean
    Bound As TBoundLambda
    
    UsePerformanceCache As Boolean
    PerformanceCache As Object
End Type
Private This As TThis








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
    If This.Singleton.Cache Is Nothing Then Set This.Singleton.Cache = CreateObject("Scripting.Dictionary")
    Dim sID As String: sID = bUsePerformanceCache & "-" & bSandboxExtras & ")" & sEquation
    If Not This.Singleton.Cache.exists(sID) Then
        Set This.Singleton.Cache(sID) = New stdLambda
        Call This.Singleton.Cache(sID).protInit(LambdaType.iStandardLambda, sEquation, bUsePerformanceCache, bSandboxExtras)
    End If
    
    'Return cached lambda
    Set Create = This.Singleton.Cache(sID)
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
            This.Equation = sEquation
            This.UsePerformanceCache = params(1)
            Dim bSandboxExtras As Boolean: bSandboxExtras = params(2)
            
            'Performance cache
            If This.UsePerformanceCache Then Set This.PerformanceCache = CreateObject("Scripting.Dictionary")

            'Function extensions
            Set This.FunctionExtensions = stdLambda.oFunctExt
            If bSandboxExtras Or This.FunctionExtensions Is Nothing Then Set This.FunctionExtensions = CreateObject("Scripting.Dictionary")

            This.isBoundLambda = False
            This.tokens = Tokenise(sEquation)
            This.iTokenIndex = 1
            This.iOperationIndex = 0
            This.stackSize = 0
            This.scopeCount = 0
            This.funcScope = 0
            Call parseBlock("eof")
            Call finishOperations

        Case LambdaType.iBoundLambda
            This.isBoundLambda = True
            Set This.Bound.Lambda = params(0)
            This.Bound.Args = params(1)
            This.Equation = "BOUND..."
            
            'Function extensions
            Set This.FunctionExtensions = stdLambda.oFunctExt
            If bSandboxExtras Or This.FunctionExtensions Is Nothing Then Set This.FunctionExtensions = CreateObject("Scripting.Dictionary")
        Case Else
            Err.Raise 1, "stdLambda::Init", "No lambda with that type."
    End Select
End Sub

'Run the lambda from passed parameters
'@param params as Array<Variant> - Array of parameters to run the lambda with
'@returns - The result of the lambda
Public Function Run(ParamArray params() As Variant) As Variant
Attribute Run.VB_UserMemId = 0
    If Not This.isBoundLambda Then
        'Execute top-down parser
        Call CopyVariant(Run, evaluate(This.operations, params))
    Else
        Call CopyVariant(Run, This.Bound.Lambda.RunEx(ConcatArrays(This.Bound.Args, params)))
    End If
End Function

'Run the lambda from an array of parameters
'@param params as Variant<Array<Variant>> - Array of parameters to run the lambda with
'@returns - The result of the lambda
Public Function RunEx(ByVal params As Variant) As Variant
    If Not This.isBoundLambda Then
        If Not IsArray(params) Then
            Err.Raise 1, "params to be supplied as array of arguments", ""
        End If
        
        'Execute top-down parser
        Call CopyVariant(RunEx, evaluate(This.operations, params))
    Else
        Call CopyVariant(RunEx, This.Bound.Lambda.RunEx(ConcatArrays(This.Bound.Args, params)))
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
Public Function BindGlobal(ByVal sGlobalName As String, ByVal variable As Variant) As stdLambda
    Set BindGlobal = Me
    If This.isBoundLambda Then
        Call This.Bound.Lambda.BindGlobal(sGlobalName, variable)
    Else
        If This.FunctionExtensions Is Nothing Then Set This.FunctionExtensions = CreateObject("Scripting.Dictionary")
        If IsObject(variable) Then
            Set This.FunctionExtensions(sGlobalName) = variable
        Else
            Let This.FunctionExtensions(sGlobalName) = variable
        End If
    End If
End Function

'Extend the lambda with new functions and named global variables
'@returns Object<Dictionary<string,stdICallable> | Dictionary<string,variant>> - Dictionary of functions and named global variables
Public Property Get oFunctExt() As Object
  Set oFunctExt = This.FunctionExtensions
End Property

'Get the equation which was used to create the lambda
'@returns - The equation used to create the lambda
Public Property Get Equation() as String
    Equation = This.Equation
End Property

'Implementation of stdICallable::Run
'@param params as Array<Variant> - Parameters to run the lambda with
'@returns - The result of the lambda
Private Function stdICallable_Run(ParamArray params() As Variant) As Variant
    If Not This.isBoundLambda Then
        'Execute top-down parser
        Call CopyVariant(stdICallable_Run, evaluate(This.operations, params))
    Else
        Call CopyVariant(stdICallable_Run, This.Bound.Lambda.RunEx(ConcatArrays(This.Bound.Args, params)))
    End If
End Function

'Implementation of stdICallable::RunEx
'@param params as Variant<Array<Variant>> - Array of parameters to run the lambda with
'@returns - The result of the lambda
Private Function stdICallable_RunEx(ByVal params As Variant) As Variant
    If Not IsArray(params) Then
        Err.Raise 1, "params to be supplied as array of arguments", ""
    End If
    
    If Not This.isBoundLambda Then
        'Execute top-down parser
        Call CopyVariant(stdICallable_RunEx, evaluate(This.operations, params))
    Else
        Call CopyVariant(stdICallable_RunEx, This.Bound.Lambda.RunEx(ConcatArrays(This.Bound.Args, params)))
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
Private Function stdICallable_SendMessage(ByVal sMessage As String, ByRef success As Boolean, ByVal params As Variant) As Variant
    Select Case sMessage
        Case "obj"
            Set stdICallable_SendMessage = Me
            success = True
        Case "className"
            stdICallable_SendMessage = "stdLambda"
            success = True
        Case "bindGlobal"
            'Bind global based whether this is a bound lambda or not
            Call BindGlobal(params(0), params(1))
            success = True
        Case Else
            success = False
    End Select
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
    Dim size As Integer: size = This.stackSize + 1
    
    ' Consume multiple lines
    Dim bLoop As Boolean: bLoop = True
    Do
        While optConsume("colon"): Wend
        Call parseStatement
        While optConsume("colon"): Wend
        
        Dim i As Long
        For i = LBound(endToken) To UBound(endToken)
            If peek(endToken(i)) Then
                bLoop = False
            End If
        Next
    Loop While bLoop
    
    ' Get rid of all extra expression results and declarations
    While This.stackSize > size
        Call addOperation(iMerge, , -1)
    Wend
    This.scopeCount = This.scopeCount - 1
End Sub

'Increment the number of scopes and initialise them. Scopes are used to store variables, functions and function arg counts.
Private Sub addScope()
    This.scopeCount = This.scopeCount + 1
    Dim scope As Long: scope = This.scopeCount
    ReDim Preserve This.scopes(1 To scope)
    ReDim Preserve This.scopesArgCount(1 To scope)
    Set This.scopes(scope) = CreateObject("Scripting.Dictionary")
    Set This.scopesArgCount(scope) = CreateObject("Scripting.Dictionary")
End Sub

'Parse a statement of code.
'@remark - A statement consists of either a variable assignment, a function declaration or an expression
'(typically this wouldn't classify as a statement, but for the purpose of stdLambda and simplifying parsing it does).
Private Sub parseStatement()
    If (peek("set") Or peek("let")) And peek("var", 2) And peek("equal", 3) Then
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
'14. Value (numbers, $vars, strings, booleans, brackets)
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
            Call addOperation(iLogic_Xor, , -1)
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
            Call addOperation(iLogic_Or, , -1)
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
            Call addOperation(iLogic_And, , -1)
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
        Call addOperation(iLogic_Not)
        If invert = False Then
            Call addOperation(iLogic_Not)
        End If
    End If
End Sub

'Parse comparison operators (=, <>, <, <=, >, >=, is, Like)
Private Sub parseComparisonPriority1()
    Call parseArithmeticPriority1
    Dim bLoop As Boolean: bLoop = True
    Do
        If optConsume("lessThan") Then
            Call parseArithmeticPriority1
            Call addOperation(iComparison_LessThan, , -1)
        ElseIf optConsume("lessThanEqual") Then
            Call parseArithmeticPriority1
            Call addOperation(iComparison_LessThanOrEqual, , -1)
        ElseIf optConsume("greaterThan") Then
            Call parseArithmeticPriority1
            Call addOperation(iComparison_GreaterThan, , -1)
        ElseIf optConsume("greaterThanEqual") Then
            Call parseArithmeticPriority1
            Call addOperation(iComparison_GreaterThanOrEqual, , -1)
        ElseIf optConsume("equal") Then
            Call parseArithmeticPriority1
            Call addOperation(iComparison_Equal, , -1)
        ElseIf optConsume("notEqual") Then
            Call parseArithmeticPriority1
            Call addOperation(iComparison_NotEqual, , -1)
        ElseIf optConsume("is") Then
            Call parseArithmeticPriority1
            Call addOperation(iComparison_Is, , -1)
        ElseIf optConsume("like") Then
            Call parseArithmeticPriority1
            Call addOperation(iComparison_Like, , -1)
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
            Call addOperation(iMisc_Concatenate, , -1)
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
            Call addOperation(iArithmetic_Add, , -1)
        ElseIf optConsume("subtract") Then
            Call parseArithmeticPriority3
            Call addOperation(iArithmetic_Subtract, , -1)
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
            Call addOperation(iArithmetic_Modulo, , -1)
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
            Call addOperation(iArithmetic_Multiply, , -1)
        ElseIf optConsume("divide") Then
            Call parseArithmeticPriority5
            Call addOperation(iArithmetic_Divide, , -1)
        Else
            bLoop = False
        End If
    Loop While bLoop
End Sub

'Parse unary + and - operators (i.e. -1) 
Private Sub parseArithmeticPriority5()
    If optConsume("subtract") Then
        Call parseArithmeticPriority5 'recurse
        Call addOperation(iArithmetic_Negate)
    ElseIf optConsume("add") Then
        Call parseArithmeticPriority5 'recurse
    Else
        Call parseArithmeticPriority6
    End If
End Sub

'Parse power (^) operator
Private Sub parseArithmeticPriority6() '^
    Call parseFlowPriority1
    Dim bLoop As Boolean
    Do
        If optConsume("power") Then
            Call parseArithmeticPriority6andahalf '- and + are still identity operators
            Call addOperation(iArithmetic_Power, , -1)
        Else
            bLoop = False
        End If
    Loop While bLoop
End Sub

'Parse unary + and - operators on the RHS of a power operator (i.e. 2^-1)
Private Sub parseArithmeticPriority6andahalf()
    If optConsume("subtract") Then
        Call parseArithmeticPriority6andahalf 'recurse
        Call addOperation(iArithmetic_Negate)
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
        Dim skipThenJumpIndex As Integer: skipThenJumpIndex = addOperation(iJump_IfFalse, , -1)
        
        Dim size As Integer: size = This.stackSize
        Call consume("then")
        Call parseBlock("else", "end")
        Dim skipElseJumpIndex As Integer: skipElseJumpIndex = addOperation(iJump_Unconditional)
        This.operations(skipThenJumpIndex).value = This.iOperationIndex
        This.stackSize = size
        
        If optConsume("end") Then
            Call addOperation(iPush, 0, 1) 'Expressions should always return a value
            This.operations(skipElseJumpIndex).value = This.iOperationIndex
        Else
            Call consume("else")
            Call parseBlock("eof", "rBracket", "end")
            This.operations(skipElseJumpIndex).value = This.iOperationIndex
        
            Call optConsume("end")
        End If
    Else
        Call parseValuePriority1
    End If
End Sub

'Parse evaluation of numbers, arguments, strings, booleans, variable names and brackets. Will also parse accessors on these values.
'i.e. `varName.someMethod(1,2,3).someProp`
Private Sub parseValuePriority1()
    'Prefix unary operators for set/let
    Dim iOperationType As IInstruction
    Select Case True
        Case optConsume("let"): iOperationType = iObject_PropLet
        Case optConsume("set"): iOperationType = iObject_PropSet
        Case Else:              iOperationType = iObject_FieldCall
    End Select

    If peek("literalNumber") Then
        Call addOperation(iPush, CDbl(consume("literalNumber")), 1)
    ElseIf peek("arg") Then
        Call addOperation(iAccess_Argument, Val(Mid(consume("arg"), 2)), 1)
        Call parseManyAccessors(iOperationType)
    ElseIf peek("literalString") Then
        Call parseString
    ElseIf peek("literalBoolean") Then
        Call addOperation(iPush, consume("literalBoolean") = "true", 1)
    ElseIf peek("var") Then
        If Not parseScopeAccess Then
            Call parseFunction
        End If
        Call parseManyAccessors(iOperationType)
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
    Call addOperation(iPush, consume("var"), 1)
    Dim size As Integer: size = This.stackSize
    Call parseOptParameters
    Call addOperation(iFunc_Call)
    This.stackSize = size
End Function

'Parse object field, property and method accessors. Specifically allows for more than one accessor to be chained together
'i.e. `obj.someMethod(1,"hi").someProp`
Private Sub parseManyAccessors(Optional ByVal instruction As IInstruction = iObject_FieldCall)
    Dim bLoop As Boolean: bLoop = True
    Do
        bLoop = False
        If parseOptObjectField(instruction) Then bLoop = True
        If parseOptObjectProperty(instruction) Then bLoop = True
        If parseOptObjectMethod() Then bLoop = True
    Loop While bLoop
End Sub

'Parse an object's method or property call (i.e. `obj.someMethod` or `obj.someProp`) and add it to the operations stack
'@param instruction - Whether this is a property get or property let/set
'@returns - Whether a method or property was found
Private Function parseOptObjectField(Optional ByVal instruction As IInstruction = iObject_FieldCall) As Boolean
    parseOptObjectField = False
    If optConsume("fieldAccess") Then
        Dim size As Integer: size = This.stackSize
        Call addOperation(iPush, consume("var"), 1)
        Call parseOptParameters
        'Parse Let/Set ... = ... as a special case
        If peek("equal") And instruction <> iObject_FieldCall Then
            Call consume("equal")
            Call parseExpression
            Call addOperation(instruction)
        Else
            Call addOperation(iObject_FieldCall)
        End If
        This.stackSize = size
        parseOptObjectField = True
    End If
End Function

'Parse an object's property access (i.e. Obj.someProp) and add it to the operations stack
'@param instruction - Whether this is a property get or property let/set
'@returns - Whether a property was found
Private Function parseOptObjectProperty(Optional ByVal instruction As IInstruction = iObject_PropGet) As Boolean
    parseOptObjectProperty = False
    If optConsume("propertyAccess") Then
        Dim size As Integer: size = This.stackSize
        Call addOperation(iPush, consume("var"), 1)
        Call parseOptParameters
        'Parse Let/Set ... = ... as a special case
        If peek("equal") And instruction <> iObject_PropGet Then
            Call consume("equal")
            Call parseExpression
            Call addOperation(instruction)
        Else
            Call addOperation(iObject_PropGet)
        End If
        This.stackSize = size
        parseOptObjectProperty = True
    End If
End Function

'Parse an object's method call (i.e. Obj.someMethod) and add it to the operations stack
'@returns - Whether a method was found
Private Function parseOptObjectMethod() As Boolean
    parseOptObjectMethod = False
    If optConsume("methodAccess") Then
        Dim size As Integer: size = This.stackSize
        Call addOperation(iPush, consume("var"), 1)
        Call parseOptParameters
        Call addOperation(iObject_MethodCall)
        This.stackSize = size
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
            Call addOperation(iPush, iArgCount, 1)
        End If
        parseOptParameters = True
    End If
End Function

'Parse a string literal and add it to the operations stack
Private Sub parseString()
    Dim sRes As String: sRes = consume("literalString")
    sRes = Mid(sRes, 2, Len(sRes) - 2)
    sRes = Replace(sRes, """""", """")
    Call addOperation(iPush, sRes, 1)
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
        Call addOperation(iAccess_General, 1 + offset, 1)
    Else
        This.iTokenIndex = This.iTokenIndex - 1 ' Revert token consumption
    End If
End Function

'Parse an assignment and add it to the current scope and operations stack
Private Sub parseAssignment()
    'Consume set/let keyword
    If Not optConsume("let") Then Call consume("set")
    
    Dim varName As String: varName = consume("var")
    Call consume("equal")
    Call parseExpression
    Dim offset As Long: offset = findVariable(varName)
    If offset >= 0 Then
        ' If the variable already existed, move the data to that pos on the stack
        Call addOperation(iSet_General, offset, -1)
        Call addOperation(iAccess_General, offset, 1) ' To keep a return value
    Else
        ' If the variable didn't exist yet, treat this stack pos as its source
        Call This.scopes(This.scopeCount).add(varName, This.stackSize)
    End If
End Sub

'Find the position on the Operations stack of a variable by name
'@param varName - Name of variable to find
'@returns - The position of the variable on the Operations stack
Private Function findVariable(varName As String) As Long
    Dim scope As Long: scope = This.scopeCount
    findVariable = -1
    While scope > 0
        If This.scopes(scope).exists(varName) Then
            If scope < This.funcScope Then
                Call Throw("Can't access """ & varName & """, functions can unfortunately not access data outside their block")
            ElseIf This.scopesArgCount(scope).exists(varName) Then
                Call Throw("Expected a variable, but found a function for name " & varName)
            Else
                findVariable = This.stackSize - This.scopes(scope).item(varName)
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
        Dim returnPosIndex As Integer: returnPosIndex = addOperation(iPush, , 1)
        
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
        Call addOperation(iJump_Unconditional, funcPos, -iArgCount) 'only -argCount since pushing Result and popping return pos cancel out
        This.operations(returnPosIndex).value = This.iOperationIndex
    Else
        This.iTokenIndex = This.iTokenIndex - 1 ' Revert token consumption
    End If
End Function

'Parse a function declaration and add it to the current scope and operations stack
Private Sub parseFunctionDeclaration()
    ' Create a dedicated scope for this funcion
    Call addScope
    Dim prevFuncScope As Long: prevFuncScope = This.funcScope
    This.funcScope = This.scopeCount
    
    ' Add operation to skip this code in normal operation flow
    Dim skipToIndex As Integer: skipToIndex = addOperation(iJump_Unconditional)
    
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
    Call This.scopes(This.scopeCount - 1).add(funcName, This.iOperationIndex)
    Call This.scopesArgCount(This.scopeCount - 1).add(funcName, iArgCount)
    
    ' Obtain the body
    Call parseBlock("end")
    Call consume("end")
    While iArgCount > 0
        Call addOperation(iMerge, , -1)
        iArgCount = iArgCount - 1
    Wend
    Call addOperation(iReturn_WithValue, , -1)
    This.operations(skipToIndex).value = This.iOperationIndex
    
    ' Reset the scope
    This.scopeCount = This.scopeCount - 1
    This.funcScope = prevFuncScope
End Sub

'Parse a parameter declaration and add it to the current scope
Private Sub parseParameterDeclaration()
    Dim varName As String: varName = consume("var")
    Dim offset As Long: offset = findVariable(varName)
    If offset >= 0 Then
        Call Throw("You can't declare multiple parameters with the same name")
    Else
        ' Reserve a spot for this parameter, it will be pushed by the caller
        This.stackSize = This.stackSize + 1
        Call This.scopes(This.scopeCount).add(varName, This.stackSize)
    End If
End Sub

'Find function by name
'@param varName as string - Name of function to find
'@param argCount as long - Number of arguments the function takes (out)
'@returns - The position of the function in the Operations array
Private Function findFunction(varName As String, Optional ByRef argCount As Long) As Long
    Dim scope As Long: scope = This.scopeCount
    findFunction = -1
    While scope > 0
        If This.scopes(scope).exists(varName) Then
            If Not This.scopesArgCount(scope).exists(varName) Then
                Call Throw("Expected a function, but found a variable for name " & varName)
            Else
                findFunction = This.scopes(scope).item(varName)
                argCount = This.scopesArgCount(scope).item(varName)
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
    
    ' Temporary variables for GoSub stack operations
    Dim tempItemToPush As Variant
    Dim tempPoppedLoc As Long ' Now stores the index of the popped item
    Dim tempArgs() As Variant
    
    ' Temporary index variables for multiple pops
    Dim idxV1 As Long, idxV2 As Long, idxV3 As Long
    
    ' If result is in performance cache then return it immediately
    If This.UsePerformanceCache Then
        Dim sPerformanceCacheID As String: sPerformanceCacheID = getPerformanceCacheID(vLastArgs)
        If This.PerformanceCache.exists(sPerformanceCacheID) Then
            Call CopyVariant(evaluate, This.PerformanceCache(sPerformanceCacheID))
            Exit Function
        End If
    End If

    ' Evaluate operations to identify result
    While opIndex <= opCount
        op = ops(opIndex)
        opIndex = opIndex + 1
        
        ' The following is a jump table for operations. This is significantly faster than using a `select case` statement. The label order must be identical to the enum order.
        ' Recommend generating this call using the following regex and list operation:
        ' Regex: "    i(\w+)"
        ' List:  "lbl$1, "
        On op.Instruction GoTo lblPush, lblPop, lblMerge, lblAccess_General, lblAccess_Argument, lblSet_General, lblArithmetic_Add, lblArithmetic_Subtract, lblArithmetic_Multiply, lblArithmetic_Divide, lblArithmetic_Power, lblArithmetic_Modulo, lblArithmetic_Negate, lblLogic_And, lblLogic_Or, lblLogic_Not, lblLogic_Xor, lblFunc_Call, lblComparison_Equal, lblComparison_NotEqual, lblComparison_GreaterThan, lblComparison_GreaterThanOrEqual, lblComparison_LessThan, lblComparison_LessThanOrEqual, lblComparison_Like, lblComparison_Is, lblMisc_Concatenate, lblJump_Unconditional, lblJump_IfTrue, lblJump_IfFalse, lblReturn_WithValue, lblReturn_NoValue, lblObject_PropGet, lblObject_PropLet, lblObject_PropSet, lblObject_MethodCall, lblObject_FieldCall, lblEndLoop
        
        ' If not handled by the above, there are 2 options:
        ' If zero then it is an end of loop instruction, skip it
        ' If >0 then it is an unknown instruction, throw an error
        If op.Instruction = 0 Then
            opIndex = opCount + 1
            GoTo lblEndLoop ' Jump to the end of the loop
        Else
            Throw "Unknown instruction in Lambda Virtual Machine: " & op.Instruction
            Exit Function
        End If

        ' NOTE: The following labels are in declaration order, but this has no effect on evaluation. The labels are in their current order purely for readability.
lblPush:
        'Inline `CopyVariant` for performance, as this is a very common operation
        If IsObject(op.value) Then
            Set tempItemToPush = op.value
        Else
            tempItemToPush = op.value ' Assign value to temp variable
        End If

        GoSub StackPush
        GoTo lblEndLoop

lblPop:
        GoSub StackPop ' Call GoSub to pop
        ' tempPoppedLoc now holds the popped index, use it as needed
        GoTo lblEndLoop

lblMerge:
        GoSub StackPop: idxV1 = tempPoppedLoc
        Call CopyVariant(stack(stackPtr - 1), stack(idxV1))
        GoTo lblEndLoop

lblAccess_General:
        Dim vAccessVar As Variant: Call CopyVariant(vAccessVar, stack(stackPtr - op.value))
        tempItemToPush = vAccessVar ' Assign value to temp variable for push
        GoSub StackPush
        GoTo lblEndLoop

lblAccess_Argument:
        Dim iArgIndex As Long: iArgIndex = op.value + LBound(vLastArgs) - 1
        If iArgIndex <= UBound(vLastArgs) Then
            'Inline `CopyVariant` for performance, as this is a very common operation
            If IsObject(vLastArgs(iArgIndex)) Then
                Set tempItemToPush = vLastArgs(iArgIndex) ' Assign object reference to temp variable for push
            Else
                tempItemToPush = vLastArgs(iArgIndex) ' Assign value to temp variable for push
            End If
            GoSub StackPush
        Else
            Call Throw("Argument " & iArgIndex & " not supplied to Lambda.")
        End If
        GoTo lblEndLoop

lblSet_General:
        GoSub StackPop: idxV1 = tempPoppedLoc
        stack(stackPtr - op.value) = stack(idxV1) ' Set directly from stack
        GoTo lblEndLoop

lblFunc_Call:
        GoSub StackGetFunctionArgs ' Get arguments from stack, writes output to `tempArgs`
        GoSub StackPop: idxV1 = tempPoppedLoc ' Function object/pointer index
        Call CopyVariant(tempItemToPush, evaluateFunc(stack(idxV1), tempArgs))
        GoSub StackPush
        GoTo lblEndLoop

lblArithmetic_Add:
        GoSub StackPop: idxV2 = tempPoppedLoc
        GoSub StackPop: idxV1 = tempPoppedLoc
        tempItemToPush = stack(idxV1) + stack(idxV2)
        GoSub StackPush
        GoTo lblEndLoop

lblArithmetic_Subtract:
        GoSub StackPop: idxV2 = tempPoppedLoc
        GoSub StackPop: idxV1 = tempPoppedLoc
        tempItemToPush = stack(idxV1) - stack(idxV2)
        GoSub StackPush
        GoTo lblEndLoop

lblArithmetic_Multiply:
        GoSub StackPop: idxV2 = tempPoppedLoc
        GoSub StackPop: idxV1 = tempPoppedLoc
        tempItemToPush = stack(idxV1) * stack(idxV2)
        GoSub StackPush
        GoTo lblEndLoop

lblArithmetic_Divide:
        GoSub StackPop: idxV2 = tempPoppedLoc
        GoSub StackPop: idxV1 = tempPoppedLoc
        tempItemToPush = stack(idxV1) / stack(idxV2)
        GoSub StackPush
        GoTo lblEndLoop

lblArithmetic_Power:
        GoSub StackPop: idxV2 = tempPoppedLoc
        GoSub StackPop: idxV1 = tempPoppedLoc
        tempItemToPush = stack(idxV1) ^ stack(idxV2)
        GoSub StackPush
        GoTo lblEndLoop

lblArithmetic_Modulo:
        GoSub StackPop: idxV2 = tempPoppedLoc
        GoSub StackPop: idxV1 = tempPoppedLoc
        tempItemToPush = stack(idxV1) Mod stack(idxV2)
        GoSub StackPush
        GoTo lblEndLoop

lblArithmetic_Negate:
        GoSub StackPop: idxV2 = tempPoppedLoc
        tempItemToPush = -stack(idxV2)
        GoSub StackPush
        GoTo lblEndLoop

lblLogic_And:
        GoSub StackPop: idxV2 = tempPoppedLoc
        GoSub StackPop: idxV1 = tempPoppedLoc
        tempItemToPush = stack(idxV1) And stack(idxV2)
        GoSub StackPush
        GoTo lblEndLoop

lblLogic_Or:
        GoSub StackPop: idxV2 = tempPoppedLoc
        GoSub StackPop: idxV1 = tempPoppedLoc
        tempItemToPush = stack(idxV1) Or stack(idxV2)
        GoSub StackPush
        GoTo lblEndLoop

lblLogic_Not:
        GoSub StackPop: idxV2 = tempPoppedLoc
        tempItemToPush = Not stack(idxV2)
        GoSub StackPush
        GoTo lblEndLoop

lblLogic_Xor:
        GoSub StackPop: idxV2 = tempPoppedLoc
        GoSub StackPop: idxV1 = tempPoppedLoc
        tempItemToPush = stack(idxV1) Xor stack(idxV2)
        GoSub StackPush
        GoTo lblEndLoop

lblComparison_Equal:
        GoSub StackPop: idxV2 = tempPoppedLoc
        GoSub StackPop: idxV1 = tempPoppedLoc
        If (VarType(stack(idxV1)) = vbError Or VarType(stack(idxV2)) = vbError) Then
            tempItemToPush = False
        Else
            tempItemToPush = (stack(idxV1) = stack(idxV2))
        End If
        GoSub StackPush
        GoTo lblEndLoop

lblComparison_NotEqual:
        GoSub StackPop: idxV2 = tempPoppedLoc
        GoSub StackPop: idxV1 = tempPoppedLoc
        If (VarType(stack(idxV1)) = vbError Or VarType(stack(idxV2)) = vbError) Then
            tempItemToPush = False
        Else
            tempItemToPush = (stack(idxV1) <> stack(idxV2))
        End If
        GoSub StackPush
        GoTo lblEndLoop

lblComparison_GreaterThan:
        GoSub StackPop: idxV2 = tempPoppedLoc
        GoSub StackPop: idxV1 = tempPoppedLoc
        If (VarType(stack(idxV1)) = vbError Or VarType(stack(idxV2)) = vbError) Then
            tempItemToPush = False
        Else
            tempItemToPush = (stack(idxV1) > stack(idxV2))
        End If
        GoSub StackPush
        GoTo lblEndLoop

lblComparison_GreaterThanOrEqual:
        GoSub StackPop: idxV2 = tempPoppedLoc
        GoSub StackPop: idxV1 = tempPoppedLoc
        If (VarType(stack(idxV1)) = vbError Or VarType(stack(idxV2)) = vbError) Then
            tempItemToPush = False
        Else
            tempItemToPush = (stack(idxV1) >= stack(idxV2))
        End If
        GoSub StackPush
        GoTo lblEndLoop

lblComparison_LessThan:
        GoSub StackPop: idxV2 = tempPoppedLoc
        GoSub StackPop: idxV1 = tempPoppedLoc
        If (VarType(stack(idxV1)) = vbError Or VarType(stack(idxV2)) = vbError) Then
            tempItemToPush = False
        Else
            tempItemToPush = (stack(idxV1) < stack(idxV2))
        End If
        GoSub StackPush
        GoTo lblEndLoop

lblComparison_LessThanOrEqual:
        GoSub StackPop: idxV2 = tempPoppedLoc
        GoSub StackPop: idxV1 = tempPoppedLoc
        If (VarType(stack(idxV1)) = vbError Or VarType(stack(idxV2)) = vbError) Then
            tempItemToPush = False
        Else
            tempItemToPush = (stack(idxV1) <= stack(idxV2))
        End If
        GoSub StackPush
        GoTo lblEndLoop

lblComparison_Like:
        GoSub StackPop: idxV2 = tempPoppedLoc
        GoSub StackPop: idxV1 = tempPoppedLoc
        If (VarType(stack(idxV1)) = vbError Or VarType(stack(idxV2)) = vbError) Then
            tempItemToPush = False
        Else
            tempItemToPush = (stack(idxV1) Like stack(idxV2))
        End If
        GoSub StackPush
        GoTo lblEndLoop

lblComparison_Is:
        GoSub StackPop: idxV2 = tempPoppedLoc
        GoSub StackPop: idxV1 = tempPoppedLoc
        If (VarType(stack(idxV1)) = vbError Or VarType(stack(idxV2)) = vbError) Then
            tempItemToPush = False
        Else
            tempItemToPush = (stack(idxV1) Is stack(idxV2))
        End If
        GoSub StackPush
        GoTo lblEndLoop

lblMisc_Concatenate:
        GoSub StackPop: idxV2 = tempPoppedLoc
        GoSub StackPop: idxV1 = tempPoppedLoc
        tempItemToPush = stack(idxV1) & stack(idxV2)
        GoSub StackPush
        GoTo lblEndLoop

lblJump_Unconditional:
        opIndex = op.value
        GoTo lblEndLoop

lblJump_IfTrue:
        GoSub StackPop: idxV1 = tempPoppedLoc
        If stack(idxV1) Then
            opIndex = op.value
        End If
        GoTo lblEndLoop

lblJump_IfFalse:
        GoSub StackPop: idxV1 = tempPoppedLoc
        If Not stack(idxV1) Then
            opIndex = op.value
        End If
        GoTo lblEndLoop

lblReturn_WithValue:
        GoSub StackPop: idxV1 = tempPoppedLoc 'Get the return value's index
        opIndex = stack(stackPtr - 1) 'Get the return position
        Call CopyVariant(stack(stackPtr - 1), stack(idxV1)) 'Set the return value from stack
        GoTo lblEndLoop

lblReturn_NoValue:
        opIndex = stack(stackPtr - 1) 'Get and Set the return position (no pop as it's a fixed index)
        GoTo lblEndLoop
        
lblObject_PropGet:
lblObject_PropLet:
lblObject_PropSet:
lblObject_MethodCall:
lblObject_FieldCall:
        'Get the name and arguments
        Dim bIsSetter As Boolean: bIsSetter = op.Instruction = iObject_PropSet Or op.Instruction = iObject_PropLet 'Corrected for op.Instruction usage
        Dim iOriginalPos As Long: iOriginalPos = stackPtr
        stackPtr = stackPtr - IIf(bIsSetter, 1, 0)
        GoSub StackGetFunctionArgs ' Get arguments from stack, writes output to `tempArgs`

        If bIsSetter Then
            ReDim Preserve tempArgs(0 To UBound(tempArgs) + 1)
            Call CopyVariant(tempArgs(UBound(tempArgs)), stack(iOriginalPos - 1))
            #If devMode Then
                stack(iOriginalPos - 1) = Empty
            #End If
        End If

        GoSub StackPop: idxV1 = tempPoppedLoc ' Pop the function name's index
        Dim funcName As Variant: funcName = stack(idxV1) ' Get function name from stack
        
        'Get caller type
        Dim callerType As VbCallType
        Select Case op.Instruction
            Case iObject_FieldCall:   callerType = vbGetOrMethod
            Case iObject_PropGet:     callerType = vbGet
            Case iObject_MethodCall:  callerType = vbMethod
            Case iObject_PropLet:     callerType = vbLet
            Case iObject_PropSet:     callerType = vbSet
        End Select
                    
        'Call rtcCallByName
        GoSub StackPop: idxV1 = tempPoppedLoc ' Pop the object's index
        Dim objToCall As Object: Set objToCall = stack(idxV1) ' Get object from stack
        
        If bIsSetter Then
            Call stdCallByName(objToCall, funcName, callerType, tempArgs)
            Call CopyVariant(tempItemToPush, tempArgs(UBound(tempArgs))) ' Set the return value to the RHS of the assignment 
        Else
            Call CopyVariant(tempItemToPush, stdCallByName(objToCall, funcName, callerType, tempArgs)) 'Obtain the return value
        End If
        GoSub StackPush
        GoTo lblEndLoop

lblEndLoop:
    Wend

    ' Add result to performance cache
    If This.UsePerformanceCache Then
        GoSub StackPop: idxV1 = tempPoppedLoc ' Get final result from stack
        Call CopyVariant(v1, stack(idxV1)) ' Copy to temp var for cache check
        If IsObject(v1) Then
            Set This.PerformanceCache(sPerformanceCacheID) = v1
        Else
            Let This.PerformanceCache(sPerformanceCacheID) = v1
        End If
    End If

    GoSub StackPop: idxV1 = tempPoppedLoc ' Final result is the top of the stack's index
    Call CopyVariant(evaluate, stack(idxV1)) ' Assign final result from stack

    Exit Function ' Exit evaluate function normally

' --- GoSub Routines for Stack Operations ---
'@remark minStackSize is a module-level constant
Dim size As Long

StackPush:
    size = UBound(stack)
    If stackPtr > size Then
        ReDim Preserve stack(0 To size * 2)
    End If
    
    'Inline CopyVariant for performance, as this is a very common operation
    If IsObject(tempItemToPush) Then
        Set stack(stackPtr) = tempItemToPush
    Else
        stack(stackPtr) = tempItemToPush
    End If
    
    stackPtr = stackPtr + 1
    Return

'Returns the index of the item that was popped
'@remark It's important to note that this function does not return the popped value, but rather the index of the item that was popped.
'Thus make sure to read the value from the stack before pushing a new value.
StackPop:
    size = UBound(stack)
    If stackPtr <= (size / 3) And stackPtr > minStackSize Then
        ReDim Preserve stack(0 To CLng(size / 2))
    End If
    
    ' Decrement the stack pointer
    stackPtr = stackPtr - 1

    ' Store the index of the item that is *about to be* popped
    tempPoppedLoc = stackPtr
    
    Return

' --- GoSub Routine for Get Arguments ---
StackGetFunctionArgs:
    'This value will either be the number of arguments to be passed to the function, or the name of the function itself (if no arguments are passed).
    Dim argCount As Variant: argCount = stack(stackPtr - 1) ' Get argCount from stack, without popping it
    
    'If argCount is a string, it means that no arguments were passed, and we should treat it as an empty array. Else
    'it is a number indicating the number of arguments to be passed.
    If VarType(argCount) = vbString Then
        tempArgs = Array()
    Else
        GoSub StackPop ' Pop the argCount from stack, no need to transfer from `tempPoppedLoc` again. 
        ReDim tempArgs(1 To argCount)
        
        ' Arguments are held on the stack in order, which means that we need to fill the array in reverse order.
        Dim iParamIndex As Long
        For iParamIndex = argCount To 1 Step -1
            GoSub StackPop
            Call CopyVariant(tempArgs(iParamIndex), stack(tempPoppedLoc)) ' Copy argument value from stack
        Next iParamIndex
    End If
    
    Return
End Function

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
                  Case vbSet
                    Set obj(funcName) = args(0)
                  Case vbLet
                    Let obj(funcName) = args(0)
                End Select
                Exit Function
        End Select
    End If

    'Call CallByName from DLL or
    #If Mac Then
        Call CopyVariant(stdCallByName, macCallByName(obj, funcName, callerType, args))
    #Else
        On Error GoTo ErrorInRTCCallByName
        Call CopyVariant(stdCallByName, rtcCallByName(obj, StrPtr(funcName), callerType, args, &H409))
    #End If
    Exit Function
ErrorInRTCCallByName:
    'HACK: Rarely objects which are poorly implemented might throw missing member errors. This can for instance happen with RecordSet. This try loop is an attempt to catch this
    Dim iTry As Long: iTry = iTry + 1
    If iTry < 5 Then
        Dim sCallerTypeName As String
        Select Case callerType
            Case vbGetOrMethod: sCallerTypeName = "Property or Method "
            Case vbGet: sCallerTypeName = "Property "
            Case vbMethod: sCallerTypeName = "Method "
        End Select
        Err.Raise Err.Number, "", sCallerTypeName & funcName & " doesn't exist on object with type " & TypeName(obj)
        Resume
    Else
        Resume
    End If
End Function

'Evaluates the built in standard functions
'@param {String} sFuncName The name of the function to invoke
'@param {Variant} args() The arguments
'@returns The result
Private Function evaluateFunc(ByVal sFuncName As String, ByVal args As Variant) As Variant
    Dim iArgStart As Long: iArgStart = LBound(args)
    If TypeName(This.FunctionExtensions) = "Dictionary" Then
        If This.FunctionExtensions.exists(sFuncName) Then
            Dim vInjectedVar As Variant
            Call CopyVariant(vInjectedVar, This.FunctionExtensions(sFuncName))
            If TypeOf vInjectedVar Is stdICallable Then
                Call CopyVariant(evaluateFunc, This.FunctionExtensions(sFuncName).RunEx(args))
            Else
                Call CopyVariant(evaluateFunc, vInjectedVar)
            End If
            Exit Function
        End If
    End If
    
    Dim i As Long
    Select Case LCase(sFuncName)
        'Useful OOP constants
        Case "thisworkbook": If IsObject(ThisWorkbook) Then Set evaluateFunc = ThisWorkbook
        Case "application":  If IsObject(Application) Then Set evaluateFunc = Application

        'Data structures
        Case "dict":
            Dim oRetDict As Object: Set oRetDict = CreateObject("Scripting.Dictionary")
            For i = iArgStart To UBound(args) Step 2
                Call oRetDict.add(args(i), args(i + 1))
            Next
            Set evaluateFunc = oRetDict

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
        Case "nothing":         Set evaluateFunc = Nothing
        Case "empty":           evaluateFunc = Empty
        Case "missing":         evaluateFunc = GetMissing()

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
        Case "val":     evaluateFunc = VBA.Conversion.Val(args(iArgStart))
        
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
            Set evaluateFunc = stdLambda.Create(args(iArgStart))
            Set evaluateFunc.oFuncExt = oFunctExt
        Case "isnumeric": evaluateFunc = IsNumeric(args(iArgStart))
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
    If Me Is stdLambda Then Set This.FunctionExtensions = CreateObject("Scripting.Dictionary")
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
    If This.iTokenIndex = 0 Then This.iTokenIndex = 1
    
    'Get next token
    ShiftTokens = This.tokens(This.iTokenIndex)
    
    'Increment token index
    This.iTokenIndex = This.iTokenIndex + 1
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
    If This.iTokenIndex = 0 Then This.iTokenIndex = 1
    If This.iTokenIndex + offset - 1 <= UBound(This.tokens) Then
        peek = This.tokens(This.iTokenIndex + offset - 1).Type.name = sTokenType
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
'@param kInstruction - The main type of the operation
'@param value - The value associated with the operation
'@param stackDelta - The effect this has on the stack size (increasing or decreasing it)
'@returns - The index of the created operation
Private Function addOperation(kInstruction As IInstruction, Optional value As Variant, Optional stackDelta As Integer) As Integer
    If This.iOperationIndex = 0 Then
        ReDim Preserve This.operations(0 To 1)
    Else
        Dim size As Long: size = UBound(This.operations)
        If This.iOperationIndex > size Then
            ReDim Preserve This.operations(0 To size * 2)
        End If
    End If
    
    With This.operations(This.iOperationIndex)
        .Instruction = kInstruction
        Call CopyVariant(.value, value)
    End With
    addOperation = This.iOperationIndex
    This.stackSize = This.stackSize + stackDelta
    
    This.iOperationIndex = This.iOperationIndex + 1
End Function

'Resizes the operations list so there are no more empty operations
Private Sub finishOperations()
    ReDim Preserve This.operations(0 To This.iOperationIndex)
End Sub

'Serializes the argument array passed to a string.
'@param {ByRef Variant()} Arguments to serialize
'@returns {String} Serialized representation of the arguments.
'@remark Objects cannot be split into their components and thus are cached as a conglomerate of type and pointer (e.g. Dictionary<12341234123>). 
'@TODO: Potentially use [StgSerializePropVariant ](https://docs.microsoft.com/en-us/windows/win32/api/propvarutil/nf-propvarutil-stgserializepropvariant) as this'd be more optimal
'@example ```vb
'Debug.Print getPerformanceCacheID(Array())=""
'Debug.Print getPerformanceCacheID(Array(Array(1, 2, Null), "yop", Empty, "", Nothing, New Collection, DateSerial(2020, 1, 1), False, True)) = "Array[1;2;null;];""yop"";empty;"""";Nothing;Collection<1720260481920>;01/01/2020;False;True;"
'```
'
'returns
'
'```
'True
'True
'```
Private Function getPerformanceCacheID(ByRef Arguments As Variant) As String
    Dim length As Long: length = UBound(Arguments) - LBound(Arguments) + 1
    If length > 0 Then
        Dim sSerialized As String: sSerialized = ""
        Dim i As Long
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
    Err.Raise 1, "stdLambda", sMessage
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

'--------------------
' Testing functions
'--------------------

'Creates an operation with the given instruction and value
'@param instruction - The instruction to use for the operation
'@param value - The value to use for the operation
'@returns - The created operation
Private Function testCreateOp(ByVal instruction As IInstruction, ByVal value As Variant) As Operation
    Dim op As Operation
    op.Instruction = instruction
    Call CopyVariant(op.value, value)
    testCreateOp = op
End Function

'Creates an array of operations with the given parameters
'@param opsParams - A set of IInstruction/Value pairs to create operations from
'@returns - An array of created operations
Private Function testCreateOps(ParamArray opsParams()) As Operation()
    Dim ops() As Operation
    If (UBound(opsParams) + 1) Mod 2 <> 0 Then Call Throw("Expected an even number of parameters")
    
    Dim size As Long: size = (UBound(opsParams) + 1) \ 2
    ReDim ops(0 To size)

    Dim i As Long
    For i = 0 To size - 1
        ops(i) = testCreateOp(opsParams(i * 2), opsParams(i * 2 + 1))
    Next i

    testCreateOps = ops
End Function

'Runs tests on the class
'@returns - An array of booleans indicating whether each test passed
Public Sub protRunVMTests()
    'iPush
    Debug.Assert evaluate(testCreateOps(iPush, 10), Array()) = 10

    'iPop
    ' Cannot directly assert by return value of evaluate unless combined with other ops
    ' To make it testable: Push 10, Push 20, Pop. If pop was a no-op, result would be 20. If pop works result should be 10.
    Debug.Assert evaluate(testCreateOps(iPush, 10, iPush, 20, iPop, Empty), Array()) = 10
    
    ' iMerge
    Debug.Assert evaluate(testCreateOps(iPush, 10, iPush, 20, iMerge, Empty), Array()) = 20
    
    'TODO: iAccess_General

    ' iAccess_Argument
    Debug.Assert evaluate(testCreateOps(iAccess_Argument, 1), Array("Hello")) = "Hello"
    Debug.Assert evaluate(testCreateOps(iAccess_Argument, 1), Array("Hello", "World")) = "Hello"
    Debug.Assert evaluate(testCreateOps(iAccess_Argument, 2), Array("Hello", "World")) = "World"
    Debug.Assert evaluate(testCreateOps(iAccess_Argument, 1), Array(Application)) Is Application
    
    'TODO: iSet_General

    ' Arithmetic operations
    Debug.Assert evaluate(testCreateOps(iPush, 10, iPush, 20, iArithmetic_Add, Empty), Array()) = 30
    Debug.Assert evaluate(testCreateOps(iPush, 10, iPush, 20, iArithmetic_Subtract, Empty), Array()) = -10
    Debug.Assert evaluate(testCreateOps(iPush, 10, iPush, 20, iArithmetic_Multiply, Empty), Array()) = 200
    Debug.Assert evaluate(testCreateOps(iPush, 10, iPush, 20, iArithmetic_Divide, Empty), Array()) = 0.5
    Debug.Assert evaluate(testCreateOps(iPush, 10, iPush, 3, iArithmetic_Power, Empty), Array()) = 1000
    Debug.Assert evaluate(testCreateOps(iPush, 20, iPush, 10, iArithmetic_Modulo, Empty), Array()) = 0
    Debug.Assert evaluate(testCreateOps(iPush, 20, iPush, 9, iArithmetic_Modulo, Empty), Array()) = 2
    Debug.Assert evaluate(testCreateOps(iPush, -20, iPush, 10, iArithmetic_Modulo, Empty), Array()) = 0
    Debug.Assert evaluate(testCreateOps(iPush, 10, iArithmetic_Negate, Empty), Array()) = -10
    
    'iLogic_And
    'A | 1 1 0 0
    'B | 1 0 1 0
    '-----------
    'O | 1 0 0 0
    Debug.Assert evaluate(testCreateOps(iPush, True, iPush, True, iLogic_And, Empty), Array()) = True
    Debug.Assert evaluate(testCreateOps(iPush, True, iPush, False, iLogic_And, Empty), Array()) = False
    Debug.Assert evaluate(testCreateOps(iPush, False, iPush, True, iLogic_And, Empty), Array()) = False
    Debug.Assert evaluate(testCreateOps(iPush, False, iPush, False, iLogic_And, Empty), Array()) = False

    'iLogic_Or
    'A | 1 1 0 0
    'B | 1 0 1 0
    '-----------
    'O | 1 1 1 0
    Debug.Assert evaluate(testCreateOps(iPush, True, iPush, True, iLogic_Or, Empty), Array()) = True
    Debug.Assert evaluate(testCreateOps(iPush, True, iPush, False, iLogic_Or, Empty), Array()) = True
    Debug.Assert evaluate(testCreateOps(iPush, False, iPush, True, iLogic_Or, Empty), Array()) = True
    Debug.Assert evaluate(testCreateOps(iPush, False, iPush, False, iLogic_Or, Empty), Array()) = False

    'iLogic_Not
    'A | 1 0
    '-----------
    'O | 0 1
    Debug.Assert evaluate(testCreateOps(iPush, True, iLogic_Not, Empty), Array()) = False
    Debug.Assert evaluate(testCreateOps(iPush, False, iLogic_Not, Empty), Array()) = True

    'iLogic_Xor
    'A | 1 1 0 0
    'B | 1 0 1 0
    '-----------
    'O | 0 1 1 0
    Debug.Assert evaluate(testCreateOps(iPush, True, iPush, True, iLogic_Xor, Empty), Array()) = False
    Debug.Assert evaluate(testCreateOps(iPush, True, iPush, False, iLogic_Xor, Empty), Array()) = True
    Debug.Assert evaluate(testCreateOps(iPush, False, iPush, True, iLogic_Xor, Empty), Array()) = True
    Debug.Assert evaluate(testCreateOps(iPush, False, iPush, False, iLogic_Xor, Empty), Array()) = False

    'iFunc_Call - Use `Len` as an example
    Debug.Assert evaluate(testCreateOps(iPush, "Len", iPush, "TestString", iPush, 1, iFunc_Call, Empty), Array()) = Len("TestString")
    
    'Test oFunctExts
    Set This.FunctionExtensions = CreateObject("Scripting.Dictionary")
    This.FunctionExtensions.Add "Proxy", stdLambda.Create("$1")
    Debug.Assert evaluate(testCreateOps(iPush, "Proxy", iPush, "TestString", iPush, 1, iFunc_Call, Empty), Array()) = "TestString"
    'Ensure returns objects too
    Debug.Assert TypeName(evaluate(testCreateOps(iPush, "Proxy", iPush, Application, iPush, 1, iFunc_Call, Empty), Array())) = "Application"
    Set This.FunctionExtensions = Nothing

    'TODO: Test other functions

    'Comparison operations
    Debug.Assert evaluate(testCreateOps(iPush, 10, iPush, 10, iComparison_Equal, Empty), Array()) = True
    Debug.Assert evaluate(testCreateOps(iPush, 10, iPush, 20, iComparison_Equal, Empty), Array()) = False
    Debug.Assert evaluate(testCreateOps(iPush, 10, iPush, 10, iComparison_NotEqual, Empty), Array()) = False
    Debug.Assert evaluate(testCreateOps(iPush, 10, iPush, 20, iComparison_NotEqual, Empty), Array()) = True
    Debug.Assert evaluate(testCreateOps(iPush, 10, iPush, 20, iComparison_GreaterThan, Empty), Array()) = False
    Debug.Assert evaluate(testCreateOps(iPush, 10, iPush, 9, iComparison_GreaterThan, Empty), Array()) = True
    Debug.Assert evaluate(testCreateOps(iPush, 10, iPush, 10, iComparison_GreaterThan, Empty), Array()) = False
    Debug.Assert evaluate(testCreateOps(iPush, 10, iPush, 11, iComparison_GreaterThan, Empty), Array()) = False
    Debug.Assert evaluate(testCreateOps(iPush, 10, iPush, 9, iComparison_GreaterThanOrEqual, Empty), Array()) = True
    Debug.Assert evaluate(testCreateOps(iPush, 10, iPush, 10, iComparison_GreaterThanOrEqual, Empty), Array()) = True
    Debug.Assert evaluate(testCreateOps(iPush, 10, iPush, 11, iComparison_GreaterThanOrEqual, Empty), Array()) = False
    Debug.Assert evaluate(testCreateOps(iPush, 10, iPush, 9, iComparison_LessThan, Empty), Array()) = False
    Debug.Assert evaluate(testCreateOps(iPush, 10, iPush, 10, iComparison_LessThan, Empty), Array()) = False
    Debug.Assert evaluate(testCreateOps(iPush, 10, iPush, 11, iComparison_LessThan, Empty), Array()) = True
    Debug.Assert evaluate(testCreateOps(iPush, 10, iPush, 9, iComparison_LessThanOrEqual, Empty), Array()) = False
    Debug.Assert evaluate(testCreateOps(iPush, 10, iPush, 10, iComparison_LessThanOrEqual, Empty), Array()) = True
    Debug.Assert evaluate(testCreateOps(iPush, 10, iPush, 11, iComparison_LessThanOrEqual, Empty), Array()) = True
    Debug.Assert evaluate(testCreateOps(iPush, "A", iPush, "B", iComparison_Like, Empty), Array()) = False
    Debug.Assert evaluate(testCreateOps(iPush, "A", iPush, "A", iComparison_Like, Empty), Array()) = True
    Debug.Assert evaluate(testCreateOps(iPush, "AB", iPush, "A*", iComparison_Like, Empty), Array()) = True
    Debug.Assert evaluate(testCreateOps(iPush, "ABC", iPush, "A*", iComparison_Like, Empty), Array()) = True
    Debug.Assert evaluate(testCreateOps(iPush, "AB", iPush, "A?", iComparison_Like, Empty), Array()) = True
    Debug.Assert evaluate(testCreateOps(iPush, "ABC", iPush, "A?", iComparison_Like, Empty), Array()) = False
    
    Dim o1 As Collection: Set o1 = New Collection
    Dim o2 As Collection: Set o2 = New Collection
    Dim o3 As Collection
    Debug.Assert evaluate(testCreateOps(iPush, o1, iPush, o2, iComparison_Is, Empty), Array()) = False
    Debug.Assert evaluate(testCreateOps(iPush, o1, iPush, o1, iComparison_Is, Empty), Array()) = True
    Debug.Assert evaluate(testCreateOps(iPush, Nothing, iPush, Nothing, iComparison_Is, Empty), Array()) = True
    Debug.Assert evaluate(testCreateOps(iPush, o3, iPush, Nothing, iComparison_Is, Empty), Array()) = True

    'Miscellaneous operations
    Debug.Assert evaluate(testCreateOps(iPush, "Hello", iPush, " World", iMisc_Concatenate, Empty), Array()) = "Hello World"

    'Jump instructions
    'These change instruction pointer, need careful sequencing for observable results
    'iJump_Unconditional
    '00: Push 10
    '01: Jump to 3
    '02: Push 20   (skipped)
    '03: <end>
    'Should return 10
    Debug.Assert evaluate(testCreateOps(iPush, 10, iJump_Unconditional, 3, iPush, 20), Array()) = 10

    'iJump_IfTrue
    '00 Push True
    '01 JumpIfTrue 4
    '02 Push 10              (skipped)
    '03 JumpUnconditional 5  (skipped)
    '04 Push 20
    '05 <end>
    'Should return 20.
    Debug.Assert evaluate(testCreateOps(iPush, True, iJump_IfTrue, 4, iPush, 10, iJump_Unconditional, 5, iPush, 20), Array()) = 20

    'Counter-example for iJump_IfTrue
    '00 Push False
    '01 JumpIfTrue 4
    '02 Push 10
    '03 JumpUnconditional 5
    '04 Push 20              (skipped)
    '05 <end>
    'Should return 10.
    Debug.Assert evaluate(testCreateOps(iPush, False, iJump_IfTrue, 4, iPush, 10, iJump_Unconditional, 5, iPush, 20), Array()) = 10

    'iJump_IfFalse (inverse of the above)
    Debug.Assert evaluate(testCreateOps(iPush, False, iJump_IfFalse, 4, iPush, 10, iJump_Unconditional, 5, iPush, 20), Array()) = 20
    Debug.Assert evaluate(testCreateOps(iPush, True, iJump_IfFalse, 4, iPush, 10, iJump_Unconditional, 5, iPush, 20), Array()) = 10

    'TODO: iReturn_WithValue - this is a utility used for returning values from functions, need to think how to test it properly
    'TODO: iReturn_NoValue - as above.
    

    'iObject_PropGet
    'TEST1 - With object return value
    '00: Push Application    'Application _
    '01: Push "ThisWorkbook" '  .ThisWorkbook
    '02: Object_PropGet
    'Should return `ThisWorkbook`
    Debug.Assert evaluate(testCreateOps(iPush, Application, iPush, "ThisWorkbook", iObject_PropGet, Empty), Array()) Is ThisWorkbook

    'TEST2 - With literal return value
    '00: Push Application    'Application _
    '01: Push "ThisWorkbook" '  .Name
    '02: Object_PropGet
    'Should return Application.Name e.g. "Microsoft Excel"
    Debug.Assert evaluate(testCreateOps(iPush, Application, iPush, "Name", iObject_PropGet, Empty), Array()) = Application.Name

    'TEST3 - With parameters
    '00: Push Application    'Application _
    '01: Push "Range"        '  .Range( _
    '02: Push "A1"           '    "A1")
    '03: Push 1              '(1 arg)
    '04: Object_PropGet
    'Should return `Range("A1")` but note these won't return the same objects so `is` will still fail.
    'Instead we check the type and address.
    Debug.Assert TypeName(evaluate(testCreateOps(iPush, Application, iPush, "Range", iPush, "A1", iPush, 1, iObject_PropGet, Empty), Array())) = "Range"
    Debug.Assert evaluate(testCreateOps(iPush, Application, iPush, "Range", iPush, "A1", iPush, 1, iObject_PropGet, Empty), Array()).Address = "$A$1"

    'TODO: iObject_MethodCall
    'TODO: iObject_MethodCall_WithArgs
    'TODO: iObject_FieldCall           as method
    'TODO: iObject_FieldCall_WithArgs  as method
    'TODO: iObject_FieldCall           as property
    'TODO: iObject_FieldCall_WithArgs  as property
    'TODO: iObject_PropLet
    'TODO: iObject_PropLet_WithArgs
    'TODO: iObject_PropSet
    'TODO: iObject_PropSet_WithArgs
End Sub


