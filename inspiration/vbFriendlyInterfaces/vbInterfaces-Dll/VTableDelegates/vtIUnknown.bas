Attribute VB_Name = "vtIUnknown"
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

Private mVTable As tIUnknown 'preallocated (static, non-Heap) Space for the VTable in mVTable

Property Get Methods() As Long()
  If mVTable.Methods(1) = 0 Then InitVTable
  Methods = mVTable.Methods
End Property

Property Get pVTable() As Long
  If mVTable.Methods(1) = 0 Then InitVTable 'initialize only when not already done
  pVTable = VarPtr(mVTable)
End Property
'**** end of the code-block for the two generic Default-Properties ****

Private Sub InitVTable() 'this method will be called only once
  vbI.AddTo mVTable.Methods, AddressOf QueryInterface
  vbI.AddTo mVTable.Methods, AddressOf AddRef
  vbI.AddTo mVTable.Methods, AddressOf Release
End Sub

'IUnknown-Delegation
Private Function QueryInterface(This As tIUnknownCallback, reqIID As tIID, ppObj As stdole.IUnknown) As HRESULT
  If VarPtr(ppObj) = 0 Then QueryInterface = E_POINTER: Exit Function
 
  If IsEqualGUID(reqIID, IID_IUnknown) <> 0 Then 'in case of a downcast to IUnknown we assign ourselves (increasing the RefCount)
    This.RefCount = This.RefCount + 1
    Assign ppObj, VarPtr(This)

  Else ' all other IID-requests are delegated to the outside
    If InVBAStopModeOrResetting Then QueryInterface = E_POINTER: Exit Function
    Dim RefCount As Long, Unk As stdole.IUnknown
        RefCount = This.RefCount
    If Not This.oIUnknown Is Nothing Then This.oIUnknown.QueryInterface This.UserData, This.pVTable, RefCount, IIDtoStr(reqIID), Unk
 
    If Not Unk Is Nothing Then
      AssignZero ppObj
      Set ppObj = Unk
      This.RefCount = RefCount
    ElseIf This.RefCount + 1 = RefCount Then 'RefCount was incremented in the callback, to signalize that the instance is supported directly
      This.RefCount = This.RefCount + 1
      Assign ppObj, VarPtr(This)
    Else
      QueryInterface = E_NOINTERFACE
    End If
  End If
End Function

Private Function AddRef(This As tIUnknownCallback) As Long
  This.RefCount = This.RefCount + 1
  AddRef = This.RefCount
End Function

Private Function Release(This As tIUnknownCallback) As Long
  This.RefCount = This.RefCount - 1
  Release = This.RefCount
 
  If This.RefCount <= 0 Then
    If Not InVBAStopModeOrResetting Then 'we cleanup the callback-refs ourselves
      If Not This.oIUnknown Is Nothing Then This.oIUnknown.Terminate This.UserData, This.pVTable
      Set This.oIUnknown = Nothing 'release the CallBack-Object we were linked to
    End If
    CoTaskMemFree VarPtr(This)   'and destroy the Memory-allocation behind our This-Pointer
  End If
End Function

'Helper-Function to create new Instances
Public Sub CreateBaseInstance(pVTableToUse As Long, ByVal ImplementingCallbackObj As vbIUnknown, pVarPtrNewInstance As Long, _
                              Optional ByVal UserData As Long = -1, Optional ByVal pExtUserData As Long, Optional ByVal ExtByteLen As Long)
Dim pMem As Long, IUC As tIUnknownCallback
    pMem = CoTaskMemAlloc(LenB(IUC) + ExtByteLen)
 If pMem = 0 Then Err.Raise vbObjectError, , "Couldn't create Memory for the new Instance"
    
Set IUC.oIUnknown = ImplementingCallbackObj
    IUC.pVTable = pVTableToUse
    IUC.RefCount = 1 'init the RefCount-Member to 1
 
    Assign ByVal pVarPtrNewInstance, pMem 'assign the new initialized Object-Reference
    AssignZero ImplementingCallbackObj    'to avoid weak-referencing of the ImplementingCallbackObj
    
    If pExtUserData Then 'for optional extension of the only 16Byte long Base-Allocation
      Assign ByVal pMem + LenB(IUC), ByVal pExtUserData, ExtByteLen
      AssignZero ByVal pExtUserData, ExtByteLen
      IUC.UserData = pMem + LenB(IUC) 'in case of Extended UserData, the UserData.Member itself will carry the start-address of this extended "bag"
    Else
      IUC.UserData = UserData
    End If
    Assign ByVal pMem, IUC, LenB(IUC)
End Sub

