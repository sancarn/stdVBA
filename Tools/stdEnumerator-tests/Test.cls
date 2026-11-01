VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "Test"
Attribute VB_PredeclaredId = True
Option Explicit
Public Count As Long
Public Sub Topic(ByVal label As String)
End Sub
Public Sub Assert(ByVal label As String, ByVal passed As Boolean)
    Count = Count + 1
    If Not passed Then Err.Raise 5, "Test.Assert", label
End Sub
Public Sub ResetTimer()
End Sub
Public Property Get Range() As Object
    Set Range = ThisWorkbook.Worksheets(1).Range("A1")
End Property
