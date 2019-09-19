VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbIEnumerable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit 'just a generic Enumeration-Helper-Class (which makes use of IEnumVariant)

Implements vbIUnknown
Implements vbIDispatch
Implements vbIEnumVariant

Private mCallback As vbIEnumerable, mStartIdx As Long, mUserData

Public Function ElementForIndex(ByVal Index As Long, UserData As Variant, CancelEnumeration As Boolean) As Variant
End Function

Public Function EnumerateOn(oCallback As vbIEnumerable, ByVal StartIndex As Long, UserData) As Object
  Set mCallback = oCallback
      If IsObject(UserData) Then Set mUserData = UserData Else mUserData = UserData
      mStartIdx = StartIndex
  vbI.NewInstance vbI.pVT(vtbl_IDispatch), Me, VarPtr(EnumerateOn), StartIndex
End Function

'************* IUnknown-Implementation ******************
Private Sub vbIUnknown_QueryInterface(UserData As Long, ByVal pVTable As Long, RefCount As Long, sReqIID As String, Unk As stdole.IUnknown)
  If vbI.IIDsEqual(sReqIID, vbI.sIID_IEnumVariant) Then RefCount = RefCount + 1
End Sub

Private Sub vbIUnknown_Terminate(UserData As Long, ByVal pVTable As Long)
End Sub
 
'************* IDispatch-Implementation ****************
Private Function vbIDispatch_GetIDForMemberName(UserData As Long, ByVal pVTable As Long, MemberName As String) As Long
End Function

Private Function vbIDispatch_Invoke(UserData As Long, ByVal pVTable As Long, ByVal DispID As Long, ByVal CallType As VbCallType, VResult As Variant, ParamArray P() As Variant) As vbInterfaces.HRESULT
  If DispID = -4 Then '<- here it becomes quite obvious, why we have to "mark" Enumeration-Methods with a -4 in normal Code
    Dim Unk As stdole.IUnknown
    vbI.NewInstance vbI.pVT(vtbl_IEnumVariant), Me, VarPtr(Unk), UserData
    Set VResult = Unk 'return the IEnumVariant-instance we have just created above (in Unk)
  End If
End Function

'************ IEnumVariant-Implementation ***************
Private Function vbIEnumVariant_NextElm(Index As Long, ByVal cElements As Long, VariantArrayFirstElement As Variant, ElementsFetched As Long) As vbInterfaces.HRESULT
  If mCallback Is Nothing Then vbIEnumVariant_NextElm = S_FALSE: Exit Function
  
  Dim CancelEnumeration As Boolean
  
  If cElements = 1 Then 'the normal case (only one single "next-element" is wanted)
    VariantArrayFirstElement = mCallback.ElementForIndex(Index, mUserData, CancelEnumeration) 'just assign the LngArr-Value into the Variant
    If CancelEnumeration Then vbIEnumVariant_NextElm = S_FALSE Else ElementsFetched = 1

  ElseIf cElements > 1 Then 'that one desires more than one element is a rare case - but possible, e.g. in VBs WatchWindow-Mode
    Dim pV As Long: pV = VarPtr(VariantArrayFirstElement)
    Do Until ElementsFetched = cElements
      vbI.VariantCopyToPtr pV, mCallback.ElementForIndex(Index + ElementsFetched, mUserData, CancelEnumeration)
      pV = pV + 16
      If CancelEnumeration Then Exit Do Else ElementsFetched = ElementsFetched + 1
    Loop
  End If
  
  Index = Index + ElementsFetched 'increment the instance-internal (UserData-based)Index for the next call
End Function

Private Function vbIEnumVariant_Reset(Index As Long) As vbInterfaces.HRESULT
  Index = mStartIdx 'reset the Index to the start of the Container or Array, we want to enumerate
End Function

Private Function vbIEnumVariant_Skip(Index As Long, ByVal cElements As Long) As vbInterfaces.HRESULT
  Index = Index + cElements 'Skip is equally simple - just increment the internal Index by cElements
End Function
 
