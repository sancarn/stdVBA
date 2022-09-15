VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Public Function Create() as ClassTemplate
  set Create = new ClassTemplate
  Call Create.protInit()
End Function

Friend Sub protInit()

End Sub