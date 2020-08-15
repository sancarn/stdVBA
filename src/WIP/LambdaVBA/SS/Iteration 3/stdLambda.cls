VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdLambda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





'
'
'

'Can implement extensions (additional functions/operators) e.g.
'set Functions(".") = stdCallback.Create("Object",someObj,"MethodCaller", vbMethod)
'set Functions("$") = stdCallback.Create("Object",someObj,"GetterCaller", vbGet   )                
Public Functions As Object


'Examples:
'Resolve({".":[{"var":"a"},"x"], {"a":myObj}) --> myObj.x
Function Resolve(ByVal o As Object, Optional data as object = nothing,Optional bTypeChecks As Boolean = True) As Variant
    If o.Count > 1 Then
        Call Err.Raise(1, "Resolve", "Objects intrinsically cannot have more than 1 key,value pair")
    ElseIf o.Count = 0 Then
        Resolve = 0
        Exit Function
    Else
        Dim sKey As String
        sKey = o.keys()(0)
        
        Dim params As stdArray
        Set params = stdArray.Create()
        
        Dim v As Variant
        For Each v In o(sKey).arr
            If IsObject(v) Then
                params.push Resolve(v, bTypeChecks)
            Else
                params.push v
            End If
        Next
        
        'Ensure function exists and is bound
        If bTypeChecks Then
            If Functions.exists(sKey) Then
                'Ensure
                If TypeName(Functions(sKey)) = "stdCallback" Then
                    Resolve = Functions(sKey)(params)
                Else
                    Err.Raise 1, "", "Function isn't of type stdCallback"
                End If
            Else
                Err.Raise 1, "", "Function doesn't exist in Functions object"
            End If
        Else
            Resolve = Functions(sKey)(params)
        End If
    End If
End Function

Public Sub Test()
    Debug.Assert TestOperator("+",nothing,1,2) = 3
    Debug.Assert TestOperator("-",nothing,1,2) = -1
    Debug.Assert TestOperator("*",nothing,3,2) = 6
    Debug.Assert TestOperator("/",nothing,1,2) = 0.5
    Debug.Assert TestOperator("%",nothing,1,2) = 1
    Debug.Assert TestOperator("%",nothing,2,2) = 0
End Sub
Public Function TestOperator(ByVal sOperator as string, ByVal data as object, ParamArray params() as variant)
    Dim o as object: set o = CreateObject("Scripting.Dictionary")
    set o(sOperator) = stdArray.fromArray(params)
    TestOperator = stdLambda.Resolve(o)
End Function






Friend Function zAdd(a, b) As Double
    zAdd = params(1) + params(2)
End Function
Friend Function zSubtract(params As Collection) As Double
    zSubtract = params(1) - params(2)
End Function
Friend Function zMultiply(params As Collection) As Double
    zMultiple = params(1) * params(2)
End Function
Friend Function zDivide(params As Collection) As Double
    zDivide = params(1) / params(2)
End Function
Friend Function zModulus(params As Collection) As Double
    zModulus = params(1) Mod params(2)
End Function
Friend Function zGT(params As Collection) As Boolean
    zGT = params(1) > params(2)
End Function
Friend Function zGTE(params As Collection) As Boolean
    zGTE = params(1) >= params(2)
End Function
Friend Function zLT(params As Collection) As Boolean
    zLT = params(1) < params(2)
End Function
Friend Function zLTE(params As Collection) As Boolean
    zLTE = params(1) <= params(2)
End Function

Private Sub Class_Initialize()
    Set Functions = CreateObject("Scripting.Dictionary")
    Set Functions("+") = stdCallback.Create("Object", stdLambda, "zAdd", VbCallType.VbMethod)
    Set Functions("-") = stdCallback.Create("Object", stdLambda, "zSubtract", VbCallType.VbMethod)
    Set Functions("*") = stdCallback.Create("Object", stdLambda, "zMultiply", VbCallType.VbMethod)
    Set Functions("/") = stdCallback.Create("Object", stdLambda, "zDivide", VbCallType.VbMethod)
    Set Functions("%") = stdCallback.Create("Object", stdLambda, "zModulus", VbCallType.VbMethod)
    Set Functions("<") = stdCallback.Create("Object", stdLambda, "zLT", VbCallType.VbMethod)
    Set Functions("<=") = stdCallback.Create("Object", stdLambda, "zLTE", VbCallType.VbMethod)
    Set Functions(">") = stdCallback.Create("Object", stdLambda, "zGT", VbCallType.VbMethod)
    Set Functions(">=") = stdCallback.Create("Object", stdLambda, "zGTE", VbCallType.VbMethod)
End Sub
