VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLightWeightFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'To hand out our IDispatch-supporting LightWeight-Objects, we use this Factory-Class

Implements vbIUnknown  'the interface-stack starts with IUnknown
Implements vbIDispatch 'and ends at IDispatch (we don't need anything more in this simple example)
 
'here's the only Public Function in our "LightWeight-Factory" (returning our new lightweight-instances 'As Object')
Public Function CreateLightWeightObject() As Object
  vbI.NewInstance vbI.pVT(vtbl_IDispatch), Me, VarPtr(CreateLightWeightObject) 'the new Instance is provided over the last (Pointer)Arg
End Function

'************* IUnknown-Implementation *****************
Private Sub vbIUnknown_QueryInterface(UserData As Long, ByVal pVTable As Long, RefCount As Long, sReqIID As String, Unk As stdole.IUnknown)
  If vbI.IIDsEqual(sReqIID, vbI.sIID_IDispatch) Then RefCount = RefCount + 1  '<- signalize that we support the interface on this instance
End Sub

Private Sub vbIUnknown_Terminate(UserData As Long, ByVal pVTable As Long) 'not used here
End Sub


'************* IDispatch-Implementation ****************
Private Function vbIDispatch_GetIDForMemberName(UserData As Long, ByVal pVTable As Long, MemberName As String) As Long
  If StrComp(MemberName, "StringReflection", 1) = 0 Then vbIDispatch_GetIDForMemberName = 1
  If StrComp(MemberName, "AddLongs", 1) = 0 Then vbIDispatch_GetIDForMemberName = 2
End Function
 
Private Function vbIDispatch_Invoke(UserData As Long, ByVal pVTable As Long, ByVal DispID As Long, ByVal CallType As VbCallType, VResult As Variant, ParamArray P() As Variant) As HRESULT
  If DispID = 1 Then 'the DispID-Index for "StringReflection" - now we have to deliver on our promise, we made in GetIDForMemberName
    VResult = StringReflection(CStr(P(0)))
  
  ElseIf DispID = 2 Then 'the DispID-Index for "AddLongs" - now we have to deliver on our promise, we made in GetIDForMemberName
    VResult = AddLongs(P(0), P(1))
  
  Else 'we don't support a 'Member' (a Property or Function) with a DispID-Index that differs from 1 or 2
    vbIDispatch_Invoke = DISP_E_MEMBERNOTFOUND '<- and so we return the appropriate HRESULT
  End If
End Function


'********* two simple Private Helpers, delegated to from within the above Invoke-Callback  ***********
Private Function StringReflection(S As String) As String
  StringReflection = StrReverse(S)
End Function

Private Function AddLongs(ByVal L1 As Long, ByVal L2 As Long) As Long
  AddLongs = L1 + L2
End Function
