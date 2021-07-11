Attribute VB_Name = "stdLambdaTests"
Sub testAll()
    test.Topic "stdLambda"
    
    On Error Resume Next
    Test.Assert "Arithmetic operations", stdLambda.Create("(3*(2+5)+5*8/2^(2+1))/26").Run()=1
    Test.Assert "Logical operations", stdLambda.Create("5<3 or 5>3").Run() = true
    Test.Assert "Arguments", stdLambda.Create("$1 + $2").Run(5, 9) = 14
    test.Assert "Property access", stdLambda.Create("$1.Range(""A1"")").Run(Sheets(1)).Address(true,true,xlA1,true) Is Sheets(1).Range("A1").Address(true,true,xlA1,true)
    
    'Call methods default
    Call stdLambda.Create("$1.select").Run(Range("A1"))
    Test.Assert "Evaluate methods access", selection.Address(true,true,xlA1,true) = Range("A1").Address(true,true,xlA1,true)

    'Call methods explicit
    Call stdLambda.Create("$1.#select").Run(Range("A1"))
    Test.Assert "Evaluate methods with method caller access", selection.Address(true,true,xlA1,true) = Range("A1").Address(true,true,xlA1,true)
    
    'Call methods explicit (backwards compatbility)
    Call stdLambda.Create("$1.#select").Run(Range("A1"))
    Test.Assert "Evaluate methods with method caller access (deprecated)", selection.Address(true,true,xlA1,true) = Range("A1").Address(true,true,xlA1,true)

    'inline if
    Dim lambda As Variant
    Set lambda = stdLambda.Create("if $1 then 0 else if $2 then 1 else 1 + 1")
    Test.Assert "Inline if 1", lambda.Run(True, True)=0
    Test.Assert "Inline if 2", lambda.Run(False, True)=1
    Test.Assert "Inline if 3", lambda.Run(False, False)=2
    
    Test.Assert "Pure functions", stdLambda.Create("uCase(trim(""          oranges        "")) & len(""potatoes"")").Run() = "ORANGES8"
    Test.Assert "Multiline using :", stdLambda.Create("2+2: 5*2").Run()=10 'not really a test for whether the 1st line executed

    
    'variables
    With stdLambda.CreateMultiline(Array( _
         "test = 2", _
         "if $1 then", _
         "   smth = test + 2", _
         "   test = smth * 2", _
         "else", _
         "   test = test + 4", _
         "end", _
         "test" _
    ))
        Test.Assert "Variables 1", .Run(True)=8
        Test.Assert "Variables 2", .Run(False)=6 
    End With
    With stdLambda.Create("test = 2: if $1 then smth = test + 2: test = smth * 2 else test = test + 4 end: test ")
        Test.Assert "Variables 3", .Run(True)=8
        Test.Assert "Variables 4", .Run(False)=6
    End With
    
    'function definition
    Test.Assert "Function 1 fibonacci recursion", stdLambda.CreateMultiline(Array( _
         "fun fib(v)", _
         "  if v<=1 then", _
         "    v", _
         "  else ", _
         "    fib(v-2) + fib(v-1)", _
         "  end", _
         "end", _
         "fib($1)" _
    )).Run(20)=6765

    Test.Assert "Function 2 functions calling functions", stdLambda.CreateMultiline(Array( _
         "fun mul3(v) v * 3 end", _
         "fun mul3Add1(v) mul3(v) + 2 end", _
         "mul3Add1(2) + mul3Add1(2)" _
    )).Run()=16
    
    Test.Assert "Function 3 local vars", stdLambda.CreateMultiline(Array( _
         "someVar = 12", _
         "fun localVars(v)", _
         "  smth = 3", _
         "  if v < 2 then ", _
         "    smth = smth + 2", _
         "  end ", _
         "  smth", _
         "end", _
         "someVar + localVars(1)" _
    )).Run()=17
    
    Test.Assert "Function 4 nested functions", stdLambda.CreateMultiline(Array( _
         "fun somth()", _
         "  fun nested()", _
         "    2", _
         "  end", _
         "  nested() + nested()", _
         "end", _
         "somth()" _
    )).Run()=4
    
    'not allowed
    'Test.Assert "", stdLambda.CreateMultiline(Array( _
    '     "fun somth()", _
    '     "  fun nested()", _
    '     "    2", _
    '     "  end", _
    '     "  nested() + nested()", _
    '     "end", _
    '     "nested()" _
    ')).Run()
    
    'Test.Assert "", stdLambda.CreateMultiline(Array( _
    '     "someVar = 12", _
    '     "fun globalVars(v)", _
    '     "  smth = 3", _
    '     "  if v < 2 then ", _
    '     "    smth = smth + someVar", _
    '     "  end ", _
    '     "  smth", _
    '     "end", _
    '     "someVar + globalVars(1)" _
    ')).Run()

    'Test stdLambda::bind()
    With stdLambda.Create("Array($1,$2,$3)").Bind(1)
        Test.Assert "stdLambda::Bind() 1 Example", Join(.Run(2, 3), "|") = "1|2|3"
        With .Bind(2)
            Test.Assert "stdLambda::Bind() 2 Example", Join(.Run(3), "|") = "1|2|3"
            With .Bind(3)
                Test.Assert "stdLambda::Bind() 3 Example", Join(.Run(), "|") = "1|2|3"
            End With
        End With
        
        'In a historical version of stdLambda these would fail:
        Test.Assert "stdLambda::Bind() 4 Ensure creation of new bindings doesn't erase old bindings", Join(stdLambda.Create("Array($1,$2,$3)").Run(1, 2, 3), "|") = "1|2|3"
        Test.Assert "stdLambda::Bind() 5 Ensure creation of new bindings doesn't erase old bindings", Join(.Run(2, 3), "|") = "1|2|3"
        
        'Can also bind multiple arguments simultaneously
        With .Bind(2, "hello")
            Test.Assert "stdLambda::Bind() 6 multiple arg binding", Join(.Run(), "|") = "1|2|hello"
        End With
    End With

    'Testing global variables
    With stdLambda.Create("hello + 2")
        Call .BindGlobal("hello", 1)
        Test.Assert "stdLambda::BindGlobal() 1 Can bind global variables", .Run() = 3
    End With

    Dim oDict as object: set oDict = CreateObject("Scripting.Dictionary")

    'Testing dictionary declaration
    With stdLambda.Create("$1.TEST")
        oDict("TEST") = True
        Test.Assert "Dictionary.Method syntax", .Run(oDict)
    End With

    'Ensure late-bound bindGlobal() works
    Dim iCallable as stdICallable, bSuccess as boolean
    set iCallable = stdLambda.Create("hello")
    Call iCallable.SendMessage("bindGlobal", bSuccess, Array("hello",true))
    Test.Assert "SendMessage Success Parameter set", bSuccess
    Test.Assert "SendMessage Successful binding", iCallable.Run()
    Call iCallable.SendMessage("", bSuccess, Null)
    Test.Assert "SendMessage Fail Parameter set", Not bSuccess

    'Only perform these on a full test as they are quite slow:
    if Test.FullTest then
        'Ensure performance chaching is faster when params are the same:
        oDict("number") = 1
        
        'Time without performance cache
        iStart = Timer
        With stdLambda.Create("$1.number+1")
        For i = 1 To 10 ^ 4
            Call .Run(oDict)
        Next
        End With
        Dim timeA As Double: timeA = (Timer - iStart)
        
        'Time with performance cache
        iStart = Timer
        With stdLambda.Create("$1.number+1", True)
        For i = 1 To 10 ^ 4
            Call .Run(oDict)
        Next
        End With
        Dim timeB As Double: timeB = (Timer - iStart)
        
        test.assert "Performance cache faster than without", TimeB < TimeA
    end if
End Sub
