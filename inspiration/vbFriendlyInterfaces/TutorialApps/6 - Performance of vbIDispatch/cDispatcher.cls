VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDispatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
Implements vbIUnknown
Implements vbIDispatch
 
Private foo As String, bar As String 'two internal Private Vars, which take up Values for the Properties foo and bar
 
Public Function DispObj() As Object 'our IDispatch-supporting DispObject, which allows LateBound Method-Calling "per Dot"
  vbI.NewInstance vbI.pVT(vtbl_IDispatch), Me, VarPtr(DispObj)
End Function

Public Function foobarOnHostClass() As String 'just for performance-comparison
  foobarOnHostClass = foo & bar
End Function


'************* IUnknown-Implementation *****************
Private Sub vbIUnknown_QueryInterface(UserData As Long, ByVal pVTable As Long, RefCount As Long, sReqIID As String, Unk As stdole.IUnknown)
  If vbI.IIDsEqual(sReqIID, vbI.sIID_IDispatch) Then RefCount = RefCount + 1
End Sub

Private Sub vbIUnknown_Terminate(UserData As Long, ByVal pVTable As Long)
End Sub


'************* IDispatch-Implementation ****************
Private Function vbIDispatch_GetIDForMemberName(UserData As Long, ByVal pVTable As Long, MemberName As String) As Long
  Select Case LCase$(MemberName) 'for more than a handful of methods, a HashTable for Name-Lookup could be used alternatively of course
    Case "foo":            vbIDispatch_GetIDForMemberName = 1
    Case "bar":            vbIDispatch_GetIDForMemberName = 2
    Case "foobar":         vbIDispatch_GetIDForMemberName = 3
    Case "byrefparamtest": vbIDispatch_GetIDForMemberName = 4
  End Select
End Function

Private Function vbIDispatch_Invoke(UserData As Long, ByVal pVTable As Long, ByVal DispID As Long, ByVal CallType As VbCallType, VResult As Variant, ParamArray P() As Variant) As vbInterfaces.HRESULT
  If DispID < 1 Or DispID > 4 Then vbIDispatch_Invoke = DISP_E_MEMBERNOTFOUND: Exit Function

  If CallType And (VbGet Or VbMethod) Then 'handle the read-out-requests
    Select Case DispID
      Case 1: VResult = foo
      Case 2: VResult = bar
      Case 3: VResult = foo & bar
      Case 4: VResult = "ByRefParamTest: incoming -> " & P(0) & ", " & P(1)
              P(0) = 11: P(1) = 22 'in this case, we want to check whether changed Params are correctly back-transported to the caller
    End Select

  ElseIf CallType = VbLet Then 'non-object Property-Assignments (here we support only two of them for our class-vars foo and bar)
    Select Case DispID
      Case 1: foo = P(0)
      Case 2: bar = P(0)
    End Select
  End If
End Function
 
