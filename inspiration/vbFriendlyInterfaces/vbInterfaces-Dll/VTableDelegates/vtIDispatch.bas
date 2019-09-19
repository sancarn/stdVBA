Attribute VB_Name = "vtIDispatch"
Option Explicit

Private Type tDispParams
  pVarArgs As Long
  pDispIDs As Long
  cArgs As Long
  cNamedArgs As Long
End Type

Private Type tExcepInfo
  wCode As Integer
  wReserved As Integer
  Source As String
  Description As String
  HelpFile As String
  HelpContext As Long
  pvReserved As Long
  pfnDeferredFillIn As Long
  scode As Long
End Type

Private Type SAFEARRAY1D
  cDims As Integer
  fFeatures As Integer
  cbElements As Long
  cLocks As Long
  pvData As Long
  cElements1D As Long
  lLbound1D As Long
End Type

Private Type tIUnknownCallback 'our Object-Instances will occupy only 16Bytes (that's the size of a Variant-Type)
  pVTable As Long
  RefCount As Long
  oIUnknown As vbIUnknown
  UserData As Long
End Type
 
Private Type tIUnknown    'VTablePointers to IUnknown
  Methods(1 To 3) As Long 'static space for 3 Methods (we get more specific with Names in InitVTable)
End Type

Private Type tIDispatch   'VTablePointers to IDispatch
  vtIUnknown As tIUnknown 'space, to inherit the whole VTable from the tIUnknown-Type
  Methods(1 To 4) As Long 'static space for 4 Methods (we get more specific with Names in InitVTable)
End Type
 
Private mVTable As tIDispatch 'preallocated (static, non-Heap) Space for the VTable in mVTable
 
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
 
  vbI.AddTo mVTable.Methods, AddressOf GetTypeInfoCount
  vbI.AddTo mVTable.Methods, AddressOf GetTypeInfo
  vbI.AddTo mVTable.Methods, AddressOf GetIDsOfNames
  vbI.AddTo mVTable.Methods, AddressOf Invoke
End Sub

Private Function GetTypeInfoCount(This As tIUnknownCallback, cTInfo As Long) As HRESULT
  If VarPtr(cTInfo) = 0 Then GetTypeInfoCount = E_INVALIDARG: Exit Function
  cTInfo = 0 'we don't support TypeInfos by default
End Function

Private Function GetTypeInfo(This As tIUnknownCallback, ByVal iTInfo As Long, ByVal LCID As Long, oTypeInfo As stdole.IUnknown) As HRESULT
  If VarPtr(oTypeInfo) = 0 Then GetTypeInfo = E_INVALIDARG: Exit Function
  If iTInfo <> 0 Then GetTypeInfo = DISP_E_BADINDEX: Exit Function

  If oTypeInfo Is Nothing Then GetTypeInfo = E_NOTIMPL 'when no TypeInfo was set in the callee, then we return "not implemented"
End Function

Private Function GetIDsOfNames(This As tIUnknownCallback, ByVal pIID As Long, pNames As Long, ByVal cNames As Long, ByVal LCID As Long, pDispIDFirstMember As Long) As HRESULT
  Dim Impl As vbIDispatch: Set Impl = This.oIUnknown 'do the proper cast, before performing the call
  If cNames <> 1 Or VarPtr(pDispIDFirstMember) = 0 Then GetIDsOfNames = E_INVALIDARG: Exit Function
 
  pDispIDFirstMember = Impl.GetIDForMemberName(This.UserData, This.pVTable, GetStringFromPointerW(pNames, 256))
  If pDispIDFirstMember < 1 Then pDispIDFirstMember = -1: GetIDsOfNames = DISP_E_UNKNOWNNAME 'and return the correct HRESULT, in case the callee "did nothing" in the callback
End Function

Private Function Invoke(This As tIUnknownCallback, ByVal DispID&, ByVal pIID&, ByVal LCID&, ByVal Flags%, Params As tDispParams, ByVal pVarResult&, ExcepInfo As tExcepInfo, ByVal pArgErr&) As HRESULT
  Dim Impl As vbIDispatch: Set Impl = This.oIUnknown 'do the proper cast, before performing the call
  If Params.cArgs > 0 And Params.cNamedArgs > 0 And Flags < 4 Then Invoke = DISP_E_NONAMEDARGS: Exit Function
  
  On Error Resume Next 'we'll do this callback "OnError-buffered", to catch what happened on the outside
  Dim P() As Variant, saP As SAFEARRAY1D, VResult
  If Params.cArgs <= 0 Then
    Invoke = Impl.Invoke(This.UserData, This.pVTable, DispID, Flags, VResult)
  ElseIf Params.cArgs <= 10 Then
    saP.cDims = 1
    saP.cbElements = 16
    saP.cElements1D = Params.cArgs
    saP.pvData = Params.pVarArgs
    BindArray P, VarPtr(saP)
      Select Case Params.cArgs
        Case 1:  Invoke = Impl.Invoke(This.UserData, This.pVTable, DispID, Flags, VResult, P(0))
        Case 2:  Invoke = Impl.Invoke(This.UserData, This.pVTable, DispID, Flags, VResult, P(1), P(0))
        Case 3:  Invoke = Impl.Invoke(This.UserData, This.pVTable, DispID, Flags, VResult, P(2), P(1), P(0))
        Case 4:  Invoke = Impl.Invoke(This.UserData, This.pVTable, DispID, Flags, VResult, P(3), P(2), P(1), P(0))
        Case 5:  Invoke = Impl.Invoke(This.UserData, This.pVTable, DispID, Flags, VResult, P(4), P(3), P(2), P(1), P(0))
        Case 6:  Invoke = Impl.Invoke(This.UserData, This.pVTable, DispID, Flags, VResult, P(5), P(4), P(3), P(2), P(1), P(0))
        Case 7:  Invoke = Impl.Invoke(This.UserData, This.pVTable, DispID, Flags, VResult, P(6), P(5), P(4), P(3), P(2), P(1), P(0))
        Case 8:  Invoke = Impl.Invoke(This.UserData, This.pVTable, DispID, Flags, VResult, P(7), P(6), P(5), P(4), P(3), P(2), P(1), P(0))
        Case 9:  Invoke = Impl.Invoke(This.UserData, This.pVTable, DispID, Flags, VResult, P(8), P(7), P(6), P(5), P(4), P(3), P(2), P(1), P(0))
        Case 10: Invoke = Impl.Invoke(This.UserData, This.pVTable, DispID, Flags, VResult, P(9), P(8), P(7), P(6), P(5), P(4), P(3), P(2), P(1), P(0))
      End Select
    ReleaseArray P
  Else
    Err.Raise vbObjectError, , "currently only up to 10 Arguments are supported by vbIDispatch->Invoke"
  End If
 
  If Err.Number Then 'an error occured in the callee, ...
    Invoke = DISP_E_EXCEPTION 'so we return the needed HResult in this case and try to fill-in the ExcepInfo-Struct
    If VarPtr(ExcepInfo) Then  'a VarPtr-check is necessary before we do so, because it might have come in as a zero-pointer
      AssignZero ByVal VarPtr(ExcepInfo), LenB(ExcepInfo)
      ExcepInfo.Source = Err.Source
      ExcepInfo.Description = Err.Description
      ExcepInfo.HelpFile = Err.HelpFile
      ExcepInfo.HelpContext = Err.HelpContext
      If Err > 0 And Err < 1000 Then ExcepInfo.wCode = Err Else ExcepInfo.scode = Err
    End If
    Err.Clear
  ElseIf Invoke = S_OK And pVarResult <> 0 Then 'we need to copy over the Variant-Result into the pVarResult-Pointer
    Assign ByVal pVarResult, ByVal VarPtr(VResult), 16  'and do so in an efficient manner
    AssignZero ByVal VarPtr(VResult), 16  'cleanup the internal Variable, to avoid double-freeing of the variant-content
  End If
End Function
