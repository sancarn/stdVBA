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

'since this is an "early-bound extension" over the former example #1 (pure IDispatch-Mode), we now have
'to implement one more (concrete) Interface in this Factory-Class (on top of vbIUnknown and vbIDispatch)
Implements vbIUnknown
Implements vbIDispatch
Implements vbIStringsAndLongs

Private sIID_IStringsAndLongs As String
 
Private Sub Class_Initialize()
  Dim Unk As stdole.IUnknown, ConcreteInstance As vbIStringsAndLongs
  Set Unk = CreateLightWeightObject '<- we instantiate into an IUnknown-Type first
  Set ConcreteInstance = Unk 'and now force Query-Interface the very first time (to fill our sIID_IStringsAndLongs-Variable first)
End Sub

'here's the only Public Function in our "LightWeight-Factory" (returning new lightweight-instances 'As vbIStringsAndLongs' now)
Public Function CreateLightWeightObject() As vbIStringsAndLongs
  vbI.NewInstance vtIStringsAndLongs.pVTable, Me, VarPtr(CreateLightWeightObject)
End Function
 

'************* IUnknown-Implementation *****************
Private Sub vbIUnknown_QueryInterface(UserData As Long, ByVal pVTable As Long, RefCount As Long, sReqIID As String, Unk As stdole.IUnknown)
  If Len(sIID_IStringsAndLongs) = 0 Then sIID_IStringsAndLongs = sReqIID 'only done once, to gather the IID for early-binding to vbIStringsAndLongs

  If vbI.IIDInList(sReqIID, vbI.sIID_IDispatch, sIID_IStringsAndLongs) Then RefCount = RefCount + 1 '<- we support both interfaces and tell it by inreasing the RefCount
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
    VResult = vbIStringsAndLongs_StringReflection(CStr(P(0)))
  
  ElseIf DispID = 2 Then 'the DispID-Index for "AddLongs" - now we have to deliver on our promise, we made in GetIDForMemberName
    VResult = vbIStringsAndLongs_AddLongs(P(0), P(1))
  
  Else 'we don't support a 'Member' (a Property or Function) with a DispID-Index that differs from 1 or 2
    vbIDispatch_Invoke = DISP_E_MEMBERNOTFOUND '<- and so we return the appropriate HRESULT
  End If
End Function
 
'************* IStringsAndLongs-Implementation ****************
Private Function vbIStringsAndLongs_StringReflection(S As String) As String
  vbIStringsAndLongs_StringReflection = StrReverse(S)
End Function

Private Function vbIStringsAndLongs_AddLongs(ByVal L1 As Long, ByVal L2 As Long) As Long
  vbIStringsAndLongs_AddLongs = L1 + L2
End Function
