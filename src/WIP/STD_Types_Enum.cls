VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "STD_Types_Enum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Wrap - Attempt to extend class with IEnumVARIANT interface (not sure if this is possible)
Public Sub Wrap(pointer as LongPtr)
  
End Sub

'CreateMartialer - Create an Enum object which martials foreach events to parent object
'   Dim WithEvents col as STD_Types_Enum
'   Private function col_Next() as Variant
'    ...
'   end sub
'   Private function col_Length() as Long
'    ...
'   end sub
'   ...
Public Function CreateMartialer() as STD_Types_Enum
  
End Sub

'Create
