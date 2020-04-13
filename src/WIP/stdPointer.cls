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



'Gets a pointer to the last private method of an object stored in memory
'@param obj Object to get the pointer form         
'@param [iScanLimit] Limit to scan the object to. In general if your class has less than 512 methods/fields this number will not need to be altered.
Public Function GetLastPrivateMethod(ByRef obj as object, optional byval iScanLimit as long = 512) as Long
  Dim v as Long, n as Long, nAddr As Long
  Dim b as Byte, m As Byte
  
  GetMem4 ObjPtr(Me), nAddr                       'get address of the form's vtable
  nAddr = nAddr + &H6F8                           'bump to the user part of the form's vtable
  GetMem4 nAddr, n                                'read the address of the first entry point
  GetMem1 n, m                                    'read the jump opcode at the first entry point [&H33 for psuedo code, &HE9 for native code]
  For v = 1 To iScanLimit                         'scan a number of vtable entries (specified by user, default is 512)
    nAddr = nAddr + 4                             'next entry address
    GetMem4 nAddr, n                              'read the address of the entry point
    If IsBadCodePtr(n) Then GoTo vTableEnd        'is the entry point address valid code?
    GetMem1 n, b                                  'read the jump opcode at the entry point
    If b <> m Then GoTo vTableEnd                 'does the jump opcode match that of the first vtable entry?
  Next v
  Exit Function                                   'last vtable entry not found... increase the For limit?
vTableEnd:
  GetMem4 nAddr - 4, nAddr                        'back one entry to the last private method
  GetLastPrivateMethod = nAddr
End Function
