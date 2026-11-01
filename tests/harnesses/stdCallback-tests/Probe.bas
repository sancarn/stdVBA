Attribute VB_Name = "CallbackProbe"
Option Explicit
Private expected As Object
Private seen As Boolean

Public Function Probe(ByVal scenario As Long) As String
  On Error GoTo Failed
  Set expected = New Collection
  Dim cb As stdCallback
  Dim args(0 To 0) As Variant
  Set args(0) = expected
  Select Case scenario
    Case 0: Set cb = stdCallback.CreateFromPointer(AddressOf ReceiveValue)
    Case 1: Set cb = stdCallback.CreateFromPointer(AddressOf ReceiveValue, , Array(vbObject))
    Case 2: Set cb = stdCallback.CreateFromPointer(AddressOf ReceiveReference)
    Case 3: Set cb = stdCallback.CreateFromPointer(AddressOf ReceiveReference, , Array(vbObject))
    Case 4: Set cb = stdCallback.CreateFromPointer(AddressOf ReceiveReference)
    Case 5: Set cb = stdCallback.CreateFromPointer(AddressOf ReceiveVariant, , Array(vbVariant))
    Case 6: Set cb = stdCallback.CreateFromPointer(AddressOf ReceiveCollection)
    Case 7: Set cb = stdCallback.CreateFromPointer(AddressOf ReceiveValue)
    Case 8: Set cb = stdCallback.CreateFromPointer(AddressOf ReceiveValue)
    Case 9: Set cb = stdCallback.CreateFromPointer(AddressOf ReceiveValue)
  End Select
  seen = False
  Select Case scenario
    Case 4: cb.Run VarPtr(expected)
    Case 7: cb.RunEx args
    Case 8
      Set expected = Nothing
      cb.Run expected
    Case 9
      Set expected = CreateObject("Scripting.Dictionary")
      cb.Run expected
    Case Else: cb.Run expected
  End Select
  Probe = "PASS identity=" & CStr(seen)
  Exit Function
Failed:
  Probe = "ERROR " & Err.Number & ": " & Err.Description
End Function

Private Sub ReceiveValue(ByVal value As Object)
  seen = (value Is expected)
End Sub
Private Sub ReceiveReference(ByRef value As Object)
  seen = (value Is expected)
End Sub
Private Sub ReceiveVariant(ByVal value As Variant)
  seen = (value Is expected)
End Sub
Private Sub ReceiveCollection(ByVal value As Collection)
  seen = (value Is expected)
End Sub
