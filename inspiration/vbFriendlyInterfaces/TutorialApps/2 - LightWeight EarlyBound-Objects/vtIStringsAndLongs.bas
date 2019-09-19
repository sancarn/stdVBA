Attribute VB_Name = "vtIStringsAndLongs"
'This vtable-module-definition (similar to a "static-class" in other languages) is matching *exactly*
'the implementation-scheme which was used for IDispatch- and IEnumVariant-vTable-delegation in vbInterfaces.dll
'so, any useful stuff someone might want to contribute into the Dll (different from StringReflect or AddLongs <g>)
'can be taken over into the vbInterfaces.dll-Project quite easily, available "native" and usable by others then
'
'But as this example shows, sidewards enhancements (on top of what's already in vbInterfaces.dll) are doable just fine
Option Explicit
 
Private Type tIUnknownCallback 'our Object-Instances will occupy only 16Bytes (that's the size of a Variant-Type)
  pVTable As Long
  RefCount As Long
  oIUnknown As vbIUnknown
  UserData As Long
End Type

Private Type tIUnknown     'VTablePointers to IUnknown
  Methods(1 To 3) As Long  'static space for 3 Methods
End Type

Private Type tIDispatch    'VTablePointers to IDispatch
  vtIUnknown As tIUnknown  'static space, to inherit the whole VTable from the tIUnknown-Type above
  Methods(1 To 4) As Long  'static space for 4 Methods
End Type

Private Type tIStringsAndLongs 'VTablePointers to the IStringsAndLongs-Interface (VB-defined Interfaces derive from IDispatch)
  vtIDispatch As tIDispatch    'static space, to inherit the whole VTable from the tIDispatch-Type above
  Methods(1 To 2) As Long      'static space for 2 Methods (we get more specific with Names in InitVTable)
End Type
 
Private mVTable As tIStringsAndLongs 'preallocated (static, non-Heap) Space for the VTable in mVTable

'**** the following two Properties are generic, and can be left as they are in all implementations *****
Property Get Methods() As Long()
  ReDim M(1 To Len(mVTable) \ 4) As Long
  vbI.MemCopyPtr VarPtr(M(1)), pVTable, Len(mVTable)
  Methods = M
End Property

Property Get pVTable() As Long
  If mVTable.Methods(1) = 0 Then InitVTable 'initialize only when not already done
  pVTable = VarPtr(mVTable)
End Property
'**** end of the code-block for the two generic Default-Properties ****

Private Sub InitVTable() 'this method will be called only once
  vbI.CopyMethods vbI.vtIDispatchMethods, VarPtr(mVTable) 'inherit the VTable-Entries from: vtIDispatch
  
  vbI.AddTo mVTable.Methods, AddressOf StringReflection
  vbI.AddTo mVTable.Methods, AddressOf AddLongs
End Sub

Private Function StringReflection(This As tIUnknownCallback, S As String, Result As String) As HRESULT
  Dim Impl As vbIStringsAndLongs: Set Impl = This.oIUnknown 'do the proper cast, before performing the call
  Result = Impl.StringReflection(S)
End Function
 
Private Function AddLongs(This As tIUnknownCallback, ByVal L1 As Long, ByVal L2 As Long, Result As Long) As HRESULT
  Dim Impl As vbIStringsAndLongs: Set Impl = This.oIUnknown 'do the proper cast, before performing the call
  Result = Impl.AddLongs(L1, L2)
End Function
 
