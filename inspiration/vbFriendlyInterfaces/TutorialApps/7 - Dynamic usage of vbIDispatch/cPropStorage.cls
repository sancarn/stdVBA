VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPropStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
Implements vbIUnknown
Implements vbIDispatch
 
Private mDict As New Scripting.Dictionary, VStorage() As Variant
 
Public Function Props() As Object 'our IDispatch-supporting DispObject, which allows LateBound Method-Calling "per Dot"
  vbI.NewInstance vbI.pVT(vtbl_IDispatch), Me, VarPtr(Props)
End Function
 
'************* IUnknown-Implementation *****************
Private Sub vbIUnknown_QueryInterface(UserData As Long, ByVal pVTable As Long, RefCount As Long, sReqIID As String, Unk As stdole.IUnknown)
  If vbI.IIDsEqual(sReqIID, vbI.sIID_IDispatch) Then RefCount = RefCount + 1
End Sub

Private Sub vbIUnknown_Terminate(UserData As Long, ByVal pVTable As Long)
End Sub


'************* IDispatch-Implementation ****************
Private Function vbIDispatch_GetIDForMemberName(UserData As Long, ByVal pVTable As Long, MemberName As String) As Long
  If mDict.Exists(MemberName) Then
    vbIDispatch_GetIDForMemberName = mDict(MemberName)
  Else
    mDict.Add MemberName, mDict.Count + 1
    vbIDispatch_GetIDForMemberName = mDict.Count
    ReDim Preserve VStorage(1 To mDict.Count)
  End If
End Function

Private Function vbIDispatch_Invoke(UserData As Long, ByVal pVTable As Long, ByVal DispID As Long, ByVal CallType As VbCallType, VResult As Variant, ParamArray P() As Variant) As vbInterfaces.HRESULT
  If DispID < 1 Or DispID > mDict.Count Then vbIDispatch_Invoke = DISP_E_MEMBERNOTFOUND: Exit Function
  
  If CallType And (VbGet Or VbMethod) Then 'handle the read-out-requests
    If IsObject(VStorage(DispID)) Then Set VResult = VStorage(DispID) Else VResult = VStorage(DispID)
  ElseIf CallType = VbLet Then
    VStorage(DispID) = P(0)
  ElseIf CallType = VbSet Then
    Set VStorage(DispID) = P(0)
  End If
End Function
 
