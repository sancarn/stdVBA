Attribute VB_Name = "stdRefArrayTests"
Option Explicit

Public Function RunTests() As Long
    Dim v As Variant, r As stdRefArray
    v = Array(1, 2, 3)
    Set r = stdRefArray.Create(v)
    Check LBound(r.Data) = 0 And UBound(r.Data) = 2, "Variant bounds"
    r.Data(1) = 42
    Check v(1) = 42 And r(1) = 42, "Variant mutation and default accessor"
    Set r = Nothing
    Check v(1) = 42, "Wrapper destruction preserves owner"

    Dim a(-2 To 2) As Long
    a(-2) = 7
    Set r = stdRefArray.Create(a)
    Check LBound(r.Data) = -2 And UBound(r.Data) = 2, "Typed bounds"
    r.Data(0) = 99
    Check a(0) = 99, "Typed mutation"
    a(-2) = 8
    Check r.Data(-2) = 8, "Source mutation"
    Set r = Nothing

    Dim s(2 To 3, -1 To 1) As String
    s(2, -1) = "hello"
    Set r = stdRefArray.Create(s)
    Check LBound(r.Data, 1) = 2 And UBound(r.Data, 2) = 1, "2D bounds"
    r.Data(3, 0) = "world"
    Check s(3, 0) = "world", "String mutation"
    Set r = Nothing
    v = s
    Set r = stdRefArray.Create(v)
    Check r.Data(2, -1) = "hello", "Owned typed 2D array"
    Set r = Nothing

    Dim objects(0 To 0) As Object, obj As Collection
    Set obj = New Collection
    Set objects(0) = obj
    Set r = stdRefArray.Create(objects)
    Check r.Data(0) Is obj, "Object identity"
    Set obj = Nothing
    r.Data(0).Add "alive"
    Check objects(0).Count = 1, "Source retains object"
    Set r = Nothing
    Check objects(0).Count = 1, "Alias release preserves object"
    Set objects(0) = Nothing
    RunTests = 12
End Function

Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 145, "stdRefArrayTests", message
End Sub
