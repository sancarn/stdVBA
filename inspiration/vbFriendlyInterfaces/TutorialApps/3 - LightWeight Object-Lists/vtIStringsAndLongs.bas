Attribute VB_Name = "vtIAddress"
'This vtable-module-definition (similar to a "static-class" in other languages) is matching *exactly* the
'implementation-scheme and philosophy which was used for IDispatch- and IEnumVariant-vTable-delegation in vbInterfaces.dll
'so, any useful stuff someone might want to contribute into the Dll (different from StringReflect or AddLongs <g>)
'can be taken over into the vbInterfaces.dll-Project quite easily, available "native" and usable by others then
'
'This example shows, that sidewards enhancements (on top of what's already in vbInterfaces.dll) are doable just fine
Option Explicit

Public Type tmAddress 'this is the TypeDef for the "Extended-UserData" we pass at Object-construction in the factory
  ID As Long
  Name As String
  BirthDay As String
End Type
 
Private Type tIUnknownCallback 'our Object-Instances will (when non-extended) occupy only 16Bytes (that's the size of a Variant-Type)
  pVTable As Long
  RefCount As Long
  oIUnknown As vbIUnknown
  UserData As Long
  
  'but here we ensure additional Space for the Address-specific m-Variable (based on the tmAddress-Type directly above)
  m As tmAddress
End Type

Private Type tIUnknown     'VTablePointers to IUnknown
  Methods(1 To 3) As Long  'static space for the 3 IUnknown-Methods
End Type

Private Type tIDispatch    'VTablePointers to IDispatch
  vtIUnknown As tIUnknown  'static space, to inherit the whole VTable from the tIUnknown-Type above
  Methods(1 To 4) As Long  'static space for the 4 IDispatch-Methods
End Type

Private Type tIAddress      'VTablePointers to the IAddress-Interface (VB-defined Interfaces derive from IDispatch)
  vtIDispatch As tIDispatch 'static space, to inherit the whole VTable from the tIDispatch-Type above
  Methods(1 To 7) As Long   'static space for 7 Methods (we get more specific with Names in InitVTable)
End Type
 
Private mVTable As tIAddress 'preallocated (static, non-Heap) Space for the VTable in mVTable

'**** the following two Properties are generic, and can be left as they are in all implementations *****
Property Get Methods() As Long()
  Static statM(1 To Len(mVTable) \ 4) As Long
  vbI.MemCopyPtr VarPtr(statM(1)), pVTable, Len(mVTable)
  Methods = statM
End Property

Property Get pVTable() As Long
  If mVTable.Methods(1) = 0 Then InitVTable 'initialize only when not already done
  pVTable = VarPtr(mVTable)
End Property
'**** end of the code-block for the two generic Default-Properties ****

Private Sub InitVTable() 'this method will be called only once
  vbI.CopyMethods vbI.vtIDispatchMethods, VarPtr(mVTable)  'inherit the VTable-Entries from: vtIDispatch
  
  'the Interface we implement here in this module, was defined in vbIAddress this way (just pasted here for easier, direct comparison):
    'Public ID As Long
    'Public Name As String
    'Public BirthDay As Date
    'Public Function BirthDayToday() As Boolean:End Function
  
  'Ok, we add our Functions in proper VTable-Order now, in accordance to the method-order of the VB-Interface above
  vbI.AddTo mVTable.Methods, AddressOf Get_ID '<- ... Get always comes first...
  vbI.AddTo mVTable.Methods, AddressOf Let_ID '<- ... before the Let-method
  vbI.AddTo mVTable.Methods, AddressOf Get_Name
  vbI.AddTo mVTable.Methods, AddressOf Let_Name
  vbI.AddTo mVTable.Methods, AddressOf Get_BirthDay
  vbI.AddTo mVTable.Methods, AddressOf Let_BirthDay
  vbI.AddTo mVTable.Methods, AddressOf BirthDayToday 'this last one not being a Property, but a single normal Function
End Sub

'other than in the example in Demo-Folder #2, we don't delegate back into the Callback-instance (our cAddressLWeightFactory)
'(although we could) - but this time we handle it directly here in the vTable-Implementation, which *is* already our lightweight Class
Private Function Get_ID(This As tIUnknownCallback, Result As Long) As HRESULT
  Result = This.m.ID
End Function
Private Function Let_ID(This As tIUnknownCallback, ByVal RHS As Long) As HRESULT
  This.m.ID = RHS
End Function

Private Function Get_Name(This As tIUnknownCallback, Result As String) As HRESULT
  Result = This.m.Name
End Function
Private Function Let_Name(This As tIUnknownCallback, ByVal RHS As String) As HRESULT
  This.m.Name = RHS
End Function

Private Function Get_BirthDay(This As tIUnknownCallback, Result As Date) As HRESULT
  Result = This.m.BirthDay
End Function
Private Function Let_BirthDay(This As tIUnknownCallback, ByVal RHS As Date) As HRESULT
  This.m.BirthDay = RHS
End Function

Private Function BirthDayToday(This As tIUnknownCallback, Result As Boolean) As HRESULT
  Dim Today As Date
      Today = Date
  With This.m
    Result = Day(.BirthDay) = Day(Today) And Month(.BirthDay) = Month(Today)
  End With
End Function


