Attribute VB_Name = "vtIPicture"
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
 
Private Type tIPicture     'VTablePointers to the IPicture-Interface
  vtIUnknown As tIUnknown  'static space, to inherit the whole VTable from the tIUnknown-Type above
  Methods(1 To 15) As Long 'static space for 15 Methods (we get more specific with Names in InitVTable)
End Type
 
Private mVTable As tIPicture 'preallocated (static, non-Heap) Space for the VTable in mVTable

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
  vbI.CopyMethods vbI.vtIUnknownMethods, VarPtr(mVTable) 'inherit the VTable-Entries from: vtIUnknown
  
  vbI.AddTo mVTable.Methods, AddressOf GetHandle
  vbI.AddTo mVTable.Methods, AddressOf GetHPal
  vbI.AddTo mVTable.Methods, AddressOf GetPictureType
  vbI.AddTo mVTable.Methods, AddressOf GetWidth
  vbI.AddTo mVTable.Methods, AddressOf GetHeight
  vbI.AddTo mVTable.Methods, AddressOf Render
  vbI.AddTo mVTable.Methods, AddressOf SetHPal
  vbI.AddTo mVTable.Methods, AddressOf GetCurDC
  vbI.AddTo mVTable.Methods, AddressOf SelectPicture
  vbI.AddTo mVTable.Methods, AddressOf GetKeepOriginalFormat
  vbI.AddTo mVTable.Methods, AddressOf SetKeepOriginalFormat
  vbI.AddTo mVTable.Methods, AddressOf PictureChanged
  vbI.AddTo mVTable.Methods, AddressOf SaveAsFile
  vbI.AddTo mVTable.Methods, AddressOf GetAttributes
  vbI.AddTo mVTable.Methods, AddressOf SethDC
End Sub

Private Function GetHandle(This As tIUnknownCallback, Result As Long) As HRESULT
On Error GoTo Fail
Dim Impl As vbIPicture: Set Impl = This.oIUnknown 'do the proper cast, before performing the call
    Result = Impl.GetHandle(This.UserData)
Exit Function
Fail: GetHandle = E_FAIL
End Function

Private Function GetHPal(This As tIUnknownCallback, Result As Long) As HRESULT
On Error GoTo Fail
Dim Impl As vbIPicture: Set Impl = This.oIUnknown 'do the proper cast, before performing the call
    Result = Impl.GetHPal(This.UserData)
Exit Function
Fail: GetHPal = E_FAIL
End Function

Private Function GetPictureType(This As tIUnknownCallback, Result As Long) As HRESULT
On Error GoTo Fail
Dim Impl As vbIPicture: Set Impl = This.oIUnknown 'do the proper cast, before performing the call
    Result = Impl.GetPictureType(This.UserData)
Exit Function
Fail: GetPictureType = E_FAIL
End Function

Private Function GetWidth(This As tIUnknownCallback, Result As Long) As HRESULT
On Error GoTo Fail
Dim Impl As vbIPicture: Set Impl = This.oIUnknown 'do the proper cast, before performing the call
    Result = Impl.GetWidth(This.UserData)
Exit Function
Fail: GetWidth = E_FAIL
End Function

Private Function GetHeight(This As tIUnknownCallback, Result As Long) As HRESULT
On Error GoTo Fail
Dim Impl As vbIPicture: Set Impl = This.oIUnknown 'do the proper cast, before performing the call
    Result = Impl.GetHeight(This.UserData)
Exit Function
Fail: GetHeight = E_FAIL
End Function

Private Function Render(This As tIUnknownCallback, ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, _
                  ByVal xSrc As Long, ByVal ySrc As Long, ByVal cxSrc As Long, ByVal cySrc As Long, ByVal pRcBounds As Long) As HRESULT
On Error GoTo Fail
Dim Impl As vbIPicture: Set Impl = This.oIUnknown 'do the proper cast, before performing the call
    Impl.Render This.UserData, hDC, x, y, cx, cy, xSrc, ySrc, cxSrc, cySrc, pRcBounds
Exit Function
Fail: Render = E_FAIL
End Function

Private Function SetHPal(This As tIUnknownCallback, ByVal NewHPal As Long) As HRESULT
On Error GoTo Fail
Dim Impl As vbIPicture: Set Impl = This.oIUnknown 'do the proper cast, before performing the call
    Impl.SetHPal This.UserData, NewHPal
Exit Function
Fail: SetHPal = E_FAIL
End Function

Private Function GetCurDC(This As tIUnknownCallback, Result As Long) As HRESULT
On Error GoTo Fail
Dim Impl As vbIPicture: Set Impl = This.oIUnknown 'do the proper cast, before performing the call
    Result = Impl.GetCurDC(This.UserData)
Exit Function
Fail: GetCurDC = E_FAIL
End Function

Private Function SelectPicture(This As tIUnknownCallback, ByVal hDCToSelectInto As Long, hDCPrevious As Long, hBmp As Long) As HRESULT
On Error GoTo Fail
Dim Impl As vbIPicture: Set Impl = This.oIUnknown 'do the proper cast, before performing the call
    Dim hDCPrev_local As Long, hBMP_local As Long
    Impl.SelectPicture This.UserData, hDCToSelectInto, hDCPrev_local, hBMP_local
    If VarPtr(hDCPrevious) Then hDCPrevious = hDCPrev_local
    If VarPtr(hBmp) Then hBmp = hBMP_local
Exit Function
Fail: SelectPicture = E_FAIL
End Function

Private Function GetKeepOriginalFormat(This As tIUnknownCallback, Result As Long) As HRESULT
  If VarPtr(Result) Then Result = 1
End Function
Private Function SetKeepOriginalFormat(This As tIUnknownCallback, ByVal NewKeepFormat As Long) As HRESULT
End Function

Private Function PictureChanged(This As tIUnknownCallback) As HRESULT
On Error GoTo Fail
Dim Impl As vbIPicture: Set Impl = This.oIUnknown 'do the proper cast, before performing the call
    Impl.PictureChanged This.UserData
Exit Function
Fail: PictureChanged = E_FAIL
End Function

Private Function SaveAsFile(This As tIUnknownCallback, ByVal pStm As Long, ByVal fSaveMemCopy As Long, SavedBytes As Long) As HRESULT
On Error GoTo Fail
Dim Impl As vbIPicture: Set Impl = This.oIUnknown 'do the proper cast, before performing the call
    Dim SavedBytes_local As Long
    Impl.SaveAsFile This.UserData, pStm, CBool(fSaveMemCopy), SavedBytes_local
    If VarPtr(SavedBytes) Then SavedBytes = SavedBytes_local
Exit Function
Fail: SaveAsFile = E_FAIL
End Function

Private Function GetAttributes(This As tIUnknownCallback, Result As Long) As HRESULT
On Error GoTo Fail
Dim Impl As vbIPicture: Set Impl = This.oIUnknown 'do the proper cast, before performing the call
    Result = Impl.GetAttributes(This.UserData)
Exit Function
Fail: GetAttributes = E_FAIL
End Function

Private Function SethDC(This As tIUnknownCallback, ByVal NewhDC As Long) As HRESULT
'just to have the Function in place ...because stdole2.tlb defines this last VTable-member -
'though according to the MSDN it is not an official part of IPicture, anyway - here it is,
'but for the moment we will not re-delegate this call into the friendly vbIPicture-Interface
End Function
