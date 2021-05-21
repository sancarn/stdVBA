VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "uiTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function Create() as uiTable
  set Create = new uiTable
  Call Create.protInit()
End Function
Friend Sub protInit()
  '...
End Sub
