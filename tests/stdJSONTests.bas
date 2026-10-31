Attribute VB_Name = "stdJSONTests"
Option Explicit
Private calls As Long
Private order As String
Public Function RunTests(Optional ByVal unused As Long = 0) As String
  On Error GoTo Failed
  Dim pre As stdICallable: Set pre = stdCallback.CreateFromModule("stdJSONTests", "Pre")
  Dim post As stdICallable: Set post = stdCallback.CreateFromModule("stdJSONTests", "Post")
  Dim j As stdJSON
  calls = 0
  Set j = stdJSON.CreateFromVariant(Array(Array(1, 2), Null, "replace"), pre)
  If calls <> 7 Then Err.Raise 5, , "pre count " & calls
  If j.Item(1).Item(1) <> 11 Then Err.Raise 5, , "nested number"
  If Not IsNull(j.Item(2)) Then Err.Raise 5, , "null"
  If j.Item(3).Item(1) <> 13 Then Err.Raise 5, , "replacement children"
  order = ""
  Set j = stdJSON.CreateFromString("{""a"":[],""b"":{}}", postProcessor:=post)
  If order <> "AOO" Then Err.Raise 5, , "post order " & order
  If j.Item("b").Item("processed") <> True Then Err.Raise 5, , "replacement not retained"
  Set j = stdJSON.CreateFromString("{a:{},b:'yes'}", True, True, , post)
  If j.Item("b") <> "yes" Then Err.Raise 5, , "reentrant options"
  Set j = stdJSON.CreateFromVariant(Array(1, Array(2)))
  If j.Item(2).Item(1) <> 2 Then Err.Raise 5, , "no callback"
  RunTests = "PASS recursive pre-processing, replacement traversal, post-order, replacements, reentrant options and no-callback conversion"
  Exit Function
Failed:
  RunTests = "FAIL " & Err.Number & ": " & Err.Description
End Function
Public Function Pre(ByVal value As Variant) As Variant
  calls = calls + 1
  If IsObject(value) Then
    Set Pre = value
  ElseIf IsArray(value) Then
    Pre = value
  ElseIf IsNull(value) Then
    Pre = Null
  ElseIf VarType(value) = vbString Then
    Pre = Array(3)
  Else
    Pre = value + 10
  End If
End Function
Public Function Post(ByVal value As stdJSON) As stdJSON
  Dim nested As stdJSON: Set nested = stdJSON.CreateFromString("{}")
  If value.JsonType = eJSONArray Then
    order = order & "A"
    Set Post = value
  Else
    order = order & "O"
    If value.Length = 0 Then
      Set Post = stdJSON.Create()
      Post.Add "processed", True
    Else
      Set Post = value
    End If
  End If
End Function
