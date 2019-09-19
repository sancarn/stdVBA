Attribute VB_Name = "vtIEnumVariant"
Option Explicit

Private Type tIUnknownCallback 'our Object-Instances will occupy only 16Bytes (that's the size of a Variant-Type)
  pVTable As Long
  RefCount As Long
  oIUnknown As vbIUnknown
  UserData As Long
End Type

Private Type tIUnknown    'VTablePointers to IUnknown
  Methods(1 To 3) As Long 'static space for 3 Methods (we get more specific with Names in InitVTable)
End Type

Private Type tIEnumVariant 'VTablePointers to IEnumVariant
  vtIUnknown As tIUnknown  'space, to inherit the whole VTable from the tIUnknown-Type
  Methods(1 To 4) As Long  'static space for 4 Methods (we get more specific with Names in InitVTable)
End Type

Private mVTable As tIEnumVariant 'preallocated (static, non-Heap) Space for the VTable in mVTable
 
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
  vbI.CopyMethods vbI.vtIUnknownMethods, VarPtr(mVTable) 'inherit the VTable-Entries from: vtIUnknown
 
  vbI.AddTo mVTable.Methods, AddressOf NextElm
  vbI.AddTo mVTable.Methods, AddressOf Skip
  vbI.AddTo mVTable.Methods, AddressOf Reset
  vbI.AddTo mVTable.Methods, AddressOf Clone
End Sub

Private Function NextElm(This As tIUnknownCallback, ByVal cElements As Long, VariantArrayFirstElement As Variant, pElementsFetched As Long) As HRESULT
  Dim ElmFetched As Long, Impl As vbIEnumVariant: Set Impl = This.oIUnknown 'do the proper cast, before performing the call
  NextElm = Impl.NextElm(This.UserData, cElements, VariantArrayFirstElement, ElmFetched)
  If VarPtr(pElementsFetched) Then pElementsFetched = ElmFetched
End Function

Private Function Skip(This As tIUnknownCallback, ByVal cElements As Long) As HRESULT
  Dim Impl As vbIEnumVariant: Set Impl = This.oIUnknown 'do the proper cast, before performing the call
  Skip = Impl.Skip(This.UserData, cElements)
End Function

Private Function Reset(This As tIUnknownCallback) As HRESULT
  Dim Impl As vbIEnumVariant: Set Impl = This.oIUnknown 'do the proper cast, before performing the call
  Reset = Impl.Reset(This.UserData)
End Function

Private Function Clone(This As tIUnknownCallback, NewInstanceWithClonedState As stdole.IUnknown) As HRESULT
  CreateBaseInstance This.pVTable, This.oIUnknown, VarPtr(NewInstanceWithClonedState), This.UserData
End Function


