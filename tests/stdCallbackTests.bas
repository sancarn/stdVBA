Attribute VB_Name = "stdCallbackTests"
'@lang VBA

#If Win64 Then
  Private Const vbLongPtr = vbLongLong
#Else
  Private Const vbLongPtr = vbLong
#End If

Private mTest As String

Public Sub testAll()
  Test.Topic "stdCallback"

  With stdCallback.CreateFromModule("stdCallbackTests", "stdCallbackTest")
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

  'Test stdLambda::bind()
  With stdCallback.CreateFromModule("stdCallbackTests", "stdCallbackTest").Bind(1)
    Test.Assert "stdCallback::Bind() 1 Example", Join(.Run(2, 3), "|") = "1|2|3"
    With .Bind(2)
        Test.Assert "stdCallback::Bind() 2 Example", Join(.Run(3), "|") = "1|2|3"
        With .Bind(3)
          Test.Assert "stdCallback::Bind() 3 Example", Join(.Run(), "|") = "1|2|3"
        End With
    End With
    
    'In a historical version of stdLambda these would fail:
    Test.Assert "stdCallback::Bind() 4 Ensure creation of new bindings doesn't erase old bindings", Join(stdCallback.CreateFromModule("stdCallbackTests", "stdCallbackTest").Run(1, 2, 3), "|") = "1|2|3"
    Test.Assert "stdCallback::Bind() 5 Ensure creation of new bindings doesn't erase old bindings", Join(.Run(2, 3), "|") = "1|2|3"
    
    'Can also bind multiple arguments simultaneously
    With .Bind(2, "hello")
      Test.Assert "stdCallback::Bind() 6 multiple arg binding", Join(.Run(), "|") = "1|2|hello"
    End With
  End With
  
  mTest = ""
  With stdCallback.CreateFromPointer(AddressOf stdCallbackTest_Sub)
    Call .Run
    Test.Assert "stdCallback::CreateFromPointer 1 Sub no args", mTest = "hello"
  End With
  
  Const sSet As String = "hello"
  
  mTest = ""
  With stdCallback.CreateFromPointer(AddressOf stdCallbackTest_SubArg)
    Call .Run(sSet)
    Test.Assert "stdCallback::CreateFromPointer 2 Sub ByVal string arg, predicted param types", mTest = sSet
  End With
  
  mTest = ""
  With stdCallback.CreateFromPointer(AddressOf stdCallbackTest_SubArgByRef)
    Call .Run(VarPtr(sSet))
    Test.Assert "stdCallback::CreateFromPointer 3 Sub ByRef string arg, predicted param types", mTest = sSet
  End With
  
  Const vSet As Variant = "hello"
  
  mTest = ""
  With stdCallback.CreateFromPointer(AddressOf stdCallbackTest_SubArgVariant, , Array(vbVariant))
    Call .Run(vSet)
    Test.Assert "stdCallback::CreateFromPointer 4 Sub ByVal variant arg, non-predicted param types", mTest = sSet
  End With
  
  mTest = ""
  With stdCallback.CreateFromPointer(AddressOf stdCallbackTest_SubArgByRefVariant)
    Call .Run(VarPtr(vSet))
    Test.Assert "stdCallback::CreateFromPointer 5 Sub ByRef variant arg, predicted param types", mTest = sSet
  End With
  
  With stdCallback.CreateFromPointer(AddressOf stdCallbackTest_Return, vbString)
    Test.Assert "stdCallback::CreateFromPointer 6 Function returns correct data", .Run("Jim") = "Jim_jr"
  End With
  
End Sub



Public Function stdCallbackTest(ParamArray params() As Variant) As Variant
  Dim v As Variant: v = params
  stdCallbackTest = v
End Function

Public Sub stdCallbackTest_Sub()
  mTest = "hello"
End Sub
Public Sub stdCallbackTest_SubArg(ByVal s As String)
  mTest = s
End Sub
Public Sub stdCallbackTest_SubArgByRef(s As String)
  mTest = s
End Sub
Public Sub stdCallbackTest_SubArgVariant(ByVal s)
  mTest = s
End Sub
Public Sub stdCallbackTest_SubArgByRefVariant(s)
  mTest = s
End Sub

Public Function stdCallbackTest_Return(ByVal name As String) As String
  stdCallbackTest_Return = name & "_jr"
End Function
