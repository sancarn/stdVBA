Attribute VB_Name = "modMyClassFactory"
Option Explicit
 
Private Type tMyObject 'the Object-Instances will occupy only 8Bytes (that's half the size of a Variant-Type)
  pVTable As Long
  RefCount As Long
End Type
 
'IUnknown-Implementation
Public Function QueryInterface(This As tMyObject, ByVal pReqIID As Long, ppObj As stdole.IUnknown) As Long '<- HResult
  QueryInterface = &H80004002 'E_NOINTERFACE ... although since there will be no casts, this method is not called in our little Demo
End Function

Public Function AddRef(This As tMyObject) As Long
  This.RefCount = This.RefCount + 1
  AddRef = This.RefCount
End Function

Public Function Release(This As tMyObject) As Long
  This.RefCount = This.RefCount - 1
  Release = This.RefCount
  If This.RefCount = 0 Then CoTaskMemFree VarPtr(This)
End Function

'IMyClass-implementation (IMyClass only contains this single method)
Public Function AddTwoLongs(This As tMyObject, ByVal L1 As Long, ByVal L2 As Long, Result As Long) As Long '<- HResult
  Result = L1 + L2 'note, that we set the Result ByRef-Parameter - not the Function-Result (which would be used for Error-Transport)
End Function

'Factory Helper-Function to create a new Class-Instance (a new Object) of type IMyClass
Public Function CreateInstance() As IMyClass
Dim MyObj As tMyObject 'we use our UDT-based Object-Type in a Stack-Variable for more convenience
    MyObj.pVTable = modMyClassDef.VTablePtr 'whilst filling its members (as e.g. pVTable here)
    MyObj.RefCount = 1 '<- the obvious value, since we are about to create a "fresh instance"

Dim pMem As Long
    pMem = CoTaskMemAlloc(LenB(MyObj)) 'allocate space for our little 8Byte large Object
    Assign ByVal pMem, MyObj, LenB(MyObj) 'copy-over the Data from our local MyObj-UDT-Variable
    Assign CreateInstance, pMem 'assign the new initialized Object-Reference to the Function-Result
End Function
