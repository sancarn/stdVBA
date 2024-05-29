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
#If Mac Then
    #If VBA7 Then
        Private Declare PtrSafe Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As LongPtr) As LongPtr
    #Else
        Private Declare Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As Long) As Long
    #End If
#Else 'Windows
    'https://msdn.microsoft.com/en-us/library/mt723419(v=vs.85).aspx
    #If VBA7 Then
        Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    #Else
        Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    #End If
#End If

Private Const VT_BYREF As Long = &H4000
#If Win64 Then
    Private Const PTR_SIZE As Long = 8
#Else
    Private Const PTR_SIZE As Long = 4
#End If


Public Data As Variant



Public Function Create(ByRef Data As Variant) As stdRefArray
    Set Create = New stdRefArray
    Call Create.protInit(Data)
End Function
Public Sub protInit(ByRef DataIn As Variant)
    if isArray(DataIn) then
      'Set the variant data to a pointer
      Data = VarPtrArr(DataIn)

      'Set the variant type to VT_BYREF - VBA will handle the rest
      Dim vt as VbVarType: vt = VarType(DataIn) or VT_BYREF
      Call CopyMemory(Data, vt, 2)
    else
      Err.Raise 5, "stdRefArray#protInit", "Array required"
    end if
End Sub

'@param arr - Array to get the pointer of
'@returns - Pointer to the array
'@devRemark see implementation by Cristian https://stackoverflow.com/a/71081452/6302131
#If Win64 Then
Private Function VarPtrArr(ByRef arr As Variant) As LongLong
#Else
Private Function VarPtrArr(ByRef arr As Variant) As Long
#End If
    Const vtArrByRef As Long = vbArray + VT_BYREF
    Dim vt As VbVarType
    Call CopyMemory(vt, arr, 2)
    If (vt And vtArrByRef) = vtArrByRef Then
        Const pArrayOffset As Long = 8
        Call CopyMemory(VarPtrArr, ByVal VarPtr(arr) + pArrayOffset, PTR_SIZE)
    Else
        Err.Raise 5, "stdRefArray~VarPtrArr", "Array required"
    End If
End Function