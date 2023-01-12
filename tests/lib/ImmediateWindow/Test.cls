VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub Topic(ByVal sName As String)
  Debug.Print sName
  Debug.Print "============================="
End Sub

Public Sub Assert(ByVal sMsg As String, ByVal bCond As Boolean)
  Debug.Print IIf(bCond, "PASS", "FAIL") & ": " & sMsg
End Sub

Public Sub Activate()
  DoEvents
End Sub
Public Property Get Range()
  Err.Raise 1, "", "Incompatible"
End Property

Public Property Get FullTest() As Boolean
  FullTest = True
End Property

