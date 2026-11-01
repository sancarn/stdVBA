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
