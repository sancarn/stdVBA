Attribute VB_Name = "stdCallbackTests"
Public Sub testAll()
  Test.Topic "stdCallback"

  With stdCallback.CreateFromModule("stdCallbackTests", "testCallbackTest")
    'Run tests
    Dim v As Variant
    v = .Run(1, 2, 3, 4)
    Test.Assert "Run() 1 is array", isArray(v)
    If isArray(v) Then
      Test.Assert "Run() 2 array correct", Join(v, "|") = "1|2|3|4"
    End If

    'RunEx tests
    v = .RunEx(Array(1, 2, 3, 4))
    Test.Assert "RunEx() 1 is array", isArray(v)
    If isArray(v) Then
      Test.Assert "RunEx() 2 array correct", Join(v, "|") = "1|2|3|4"
    End If
  End With

  'Test object method
  With stdCallback.CreateFromObjectMethod(Test, "TestMethod")

  End With

  'Test object property
  With stdCallback.CreateFromObjectProperty(Test, "TestProperty", VbGet)
    
  End With

  'Historic evaluator method
  Test.Assert "CreateEvaluator --> stdLambda", TypeOf stdCallback.CreateEvaluator("1") Is stdICallable

  'Test stdLambda::bind()
  With stdCallback.CreateFromModule("stdCallbackTests", "testCallbackTest").Bind(1)
    Test.Assert "stdCallback::Bind() 1 Example", Join(.Run(2, 3), "|") = "1|2|3"
    With .Bind(2)
        Test.Assert "stdCallback::Bind() 2 Example", Join(.Run(3), "|") = "1|2|3"
        With .Bind(3)
          Test.Assert "stdCallback::Bind() 3 Example", Join(.Run(), "|") = "1|2|3"
        End With
    End With
    
    'In a historical version of stdLambda these would fail:
    Test.Assert "stdCallback::Bind() 4 Ensure creation of new bindings doesn't erase old bindings", Join(stdCallback.CreateFromModule("stdCallbackTests", "testCallbackTest").Run(1, 2, 3), "|") = "1|2|3"
    Test.Assert "stdCallback::Bind() 5 Ensure creation of new bindings doesn't erase old bindings", Join(.Run(2, 3), "|") = "1|2|3"
    
    'Can also bind multiple arguments simultaneously
    With .Bind(2, "hello")
      Test.Assert "stdCallback::Bind() 6 multiple arg binding", Join(.Run(), "|") = "1|2|hello"
    End With
  End With
End Sub



Public Function testCallbackTest(ParamArray params() As Variant) As Variant
  Dim v As Variant: v = params
  testCallbackTest = v
End Function
