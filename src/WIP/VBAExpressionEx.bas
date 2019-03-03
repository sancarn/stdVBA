Attribute VB_Name = "VBAExpressionEx"
Public Variables As Collection

Public Function getVar(guid As String) As Object
  Set getVar = Variables(guid)
End Function

Public Function addVar(obj As Object) As String
  'Initialise variable if not already initialised
  If Variables Is Nothing Then Set Variables = New Collection
  
  Dim guid As String
  guid = getGUID()
  
  'Save variable with identifier
  Variables.Add obj, guid
  
  'Return identifier
  addVar = guid
End Function

Private Function getGUID() As String
  'Create an identifier
  Dim guid As String, rnd As Integer
  guid = ""
  For i = 1 To 32
    rnd = Int(VBA.Math.rnd() * 16)
    guid = guid & Hex(IIf(rnd = 16, 15, rnd))
  Next
  getGUID = guid
End Function

Public Function getChild(ByVal x As Object, ByVal name As String, ParamArray params() As Variant) As Variant
    If UBound(params) - LBound(params) + 1 = 0 Then
      On Error GoTo tryVal
        Set getChild = CallByName(x, name, VbGet)
        Exit Function
tryVal:
      getChild = CallByName(x, name, VbGet)
    Else
      On Error GoTo tryVal2
        Set getChild = CallByName(x, name, VbGet, params)
        Exit Function
tryVal2:
      If Err.Number = 438 Then GoTo tryMethod
      getChild = CallByName(x, name, VbGet, params)
      Exit Function
tryMethod:
      On Error GoTo tryMethod2
      Set getChild = CallByName(x, name, VbMethod, params)
      Exit Function
tryMethod2:
      getChild = CallByName(x, name, VbMethod, params)
      Exit Function
    End If
End Function
