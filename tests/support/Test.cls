VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public checks As Long
Public failures As Long
Public lastCheck As String
Public FullTest As Boolean
Public Sub Refresh()
  checks = 0
  failures = 0
  lastCheck = ""
End Sub
Public Sub Topic(ByVal name As String)
End Sub
Public Sub Assert(ByVal name As String, ByVal passed As Boolean)
  checks = checks + 1
  If Not passed Then
    failures = failures + 1
    lastCheck = name
  End If
End Sub
Public Sub ResetTimer()
End Sub
Public Sub Activate()
End Sub
Public Property Get Range() As Object
  Dim host As Object: Set host = Application
  Set Range = host.ActiveWorkbook.Worksheets(1).Range("A1")
End Property
