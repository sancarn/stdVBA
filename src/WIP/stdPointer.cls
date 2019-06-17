'Pre requisites:



Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal bytes As Long)
Private Enum iTypeEnum
  Unknown
  Integer
  String
  Single
  Float
  Double
  Currency
  Variant
  ArrayBase
  Object
  Function
  'Anything else...?
End Enum

Private iType as iTypeEnum
Private iPtrVal as LongPtr


Public Property Get Size(Optional piType as iTypeEnum = iTypeEnum.Unknown) as Integer
  if piType = iTypeEnum.Unknown then
    piType = iType
  end if
  
  select case piType
    case iTypeEnum.Object
      Size = 4
    'case ...
      'Size = ...
    case else
      '--> Raise Error
  end select
End Function


Public Function From(ByRef x as variant) as stdPointer
  'If typename(x) = ...
  
  'end if
end function
Public Function FromPtr(ByVal piPtrVal as LongPtr, piType as iTypeEnum) as stdPointer
  Dim ret as stdPointer
  set ret = new stdPointer
  ret.init(piPtrVal, piType)
end function


Public Sub init(ByVal piPtrVal as LongPtr, piType as iTypeEnum)
  if not initialised then
    iPtrVal = piPtrVal
    iType = piType
  else
    '--> Raise error
  end if
End Sub

Public Funcion DeRef() as variant
  if iType = iTypeEnum.Object
    set DeRef = DeRefAsObject()
  else
    '--> Raise error
  end if
End Function

'Public Function Invoke() ?





'Object deref example:
'  Note: object size is 4 bytes
'    Returns an object given its pointer
'    This function reverses the effect of the ObjPtr function
Public Function DeRefAsObject() As Object
    Dim obj As Object
    ' force the value of the pointer into the temporary object variable
    CopyMemory obj, iPtrVal, Size(iTypeEnum.Object)
    ' assign to the result (this increments the ref counter)
    Set DeRefAsObject = obj
    ' manually destroy the temporary object variable 
    ' (if you omit this step you'll get a GPF!)
    CopyMemory obj, 0&, Size(iTypeEnum.Object)
End Function
