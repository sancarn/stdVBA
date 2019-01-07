VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "STD__NSTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public tArray As STD_Types_Array
Public tDate As STD_Types_Date

Private Sub Class_Initialize()
  Set tArray = STD_Types_Array
  Set tDate = STD_Types_Date
End Sub
