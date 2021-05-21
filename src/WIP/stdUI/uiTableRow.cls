VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "uiTableRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function Create() as uiTableRow
  set Create = new uiTableRow
  Call Create.protInit()
End Function
Public Sub protInit()
  '...
End Sub

Public Property Get Height()

End Function
Public Property Let Height()

End Property

Public Property Cells() as Collection

End Property