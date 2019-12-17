VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdRefArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Status WIP
'High level wrapper around flood risk.
'Ultimate goal is to make a class which calculates affordability, total project benefit, historic and overridden likelihood,
'and ultimately port code out of HFRR core into a class wrapper.

#Const DEBUG_PERF = False

'Variables for pData
Private Declare PtrSafe Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cbCopy As Long)


Public Data As Variant

Private Const VARIANT_SIZE As Long = 16

Public Function Create(ByRef Data As Variant) As stdRefArray
    Set Create = New stdRefArray
    Call Create.Init(Data)
End Function
Public Sub Init(ByRef DataIn As Variant)
    'Create direct reference to array:
    CopyMemory Data, DataIn, VARIANT_SIZE
End Sub

Private Sub Class_Terminate()
   'Clean up array reference
   FillMemory Data, VARIANT_SIZE, 0
End Sub

Public Function GetData(ByVal iRow As Long, ByVal iCol As Long) As Variant
Attribute GetData.VB_UserMemId = 0
  GetData = Data(iRow, iCol)
End Function
