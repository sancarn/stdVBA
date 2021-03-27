VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdAccTestHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private i As Long
Public Sub Add()
  i = i + 1
End Sub
Public Sub Reset()
  i = 0
End Sub
Public Property Get Count() As Long
  Count = i
End Property
