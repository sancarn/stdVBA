VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPictureEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit 'IPicture-Implementation-Demo ... this Class depends on cGDIPlusCache
                '(to render from truly shared, global resources) and on vbInterfaces.dll ... [Olaf Schmidt 2015]

Public Enum RenderAttributes
  PICTURE_DEFAULT
  PICTURE_SCALABLE = 1
  PICTURE_TRANSPARENT = 2
End Enum

Implements vbIUnknown
Implements vbIDispatch
Implements vbIPicture

Private Const HM2PX_96DPI# = 26.458333333333 'we assume a resolution of 96 DPI on all image-resources

Private mImgKey As String, mPictureType As PictureTypeConstants, mPicAttributes As RenderAttributes
Private mHighStretchQuality As Boolean, mAlpha#, mHBmp&, mHIcon&, mIsCursor As Boolean, mHSpotX&, mHSpotY&
 
Private Sub Class_Terminate() 'cleanup the three Handles, this class might (potentially) create from within the implementation-callbacks
  If mHIcon <> 0 And Not mIsCursor Then GDIPlusCache.DestroyHIcon mHIcon
  If mHIcon <> 0 And mIsCursor Then GDIPlusCache.DestroyHCursor mHIcon
  If mHBmp Then GDIPlusCache.DestroyHBmp mHBmp
End Sub

Public Function Picture(ImgKey As String, Optional ByVal PictureType As PictureTypeConstants = vbPicTypeBitmap, _
                                          Optional ByVal Attributes As RenderAttributes = PICTURE_TRANSPARENT, _
                                          Optional ByVal HighStretchQuality As Boolean, Optional ByVal Alpha As Double = 1) As StdPicture
  mIsCursor = InStr(ImgKey, ",") > 0
  If mIsCursor Then mHSpotX = Val(Split(ImgKey, ",")(1)): mHSpotY = Val(Split(ImgKey, ",")(2))
  mImgKey = Split(ImgKey, ",")(0)
  mPictureType = PictureType
  mPicAttributes = Attributes
  mHighStretchQuality = HighStretchQuality
  mAlpha = Alpha
  
  vbI.NewInstance vbI.pVT(vtbl_IDispatch), Me, VarPtr(Picture)
End Function
 
'************* IUnknown-Implementation *****************
Private Sub vbIUnknown_QueryInterface(UserData As Long, ByVal pVTable As Long, RefCount As Long, sReqIID As String, Unk As stdole.IUnknown)
  Const sIID_IPictureDisp$ = "{7BF80981-BF32-101A-8BBB-00AA00300CAB}", sIID_StdPicture$ = "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
  
  If vbI.IIDsEqual(sReqIID, vbI.sIID_IPicture) Then 'a cast to IPicture was requested
    vbI.NewInstance vbI.pVT(vtbl_IPicture), Me, VarPtr(Unk)
  ElseIf vbI.IIDInList(sReqIID, vbI.sIID_IDispatch, sIID_IPictureDisp, sIID_StdPicture) Then 'a cast to one of the Dispatch-Interfaces
    vbI.NewInstance vbI.pVT(vtbl_IDispatch), Me, VarPtr(Unk)
  End If
End Sub

Private Sub vbIUnknown_Terminate(UserData As Long, ByVal pVTable As Long) 'not used here
End Sub


'************* IDispatch-Implementation ****************
Private Function vbIDispatch_GetIDForMemberName(UserData As Long, ByVal pVTable As Long, MemberName As String) As Long
  If StrComp(MemberName, "Handle", 1) = 0 Then vbIDispatch_GetIDForMemberName = 0
  If StrComp(MemberName, "hPal", 1) = 0 Then vbIDispatch_GetIDForMemberName = 2
  If StrComp(MemberName, "Type", 1) = 0 Then vbIDispatch_GetIDForMemberName = 3
  If StrComp(MemberName, "Width", 1) = 0 Then vbIDispatch_GetIDForMemberName = 4
  If StrComp(MemberName, "Height", 1) = 0 Then vbIDispatch_GetIDForMemberName = 5
  If StrComp(MemberName, "Render", 1) = 0 Then vbIDispatch_GetIDForMemberName = 6
End Function
 
Private Function vbIDispatch_Invoke(UserData As Long, ByVal pVTable As Long, ByVal DispID As Long, ByVal CallType As VbCallType, VResult As Variant, ParamArray P() As Variant) As HRESULT
  Select Case DispID
    Case 0: VResult = vbIPicture_GetHandle(UserData)
    Case 2: VResult = vbIPicture_GetHPal(UserData)
    Case 3: VResult = vbIPicture_GetPictureType(UserData)
    Case 4: VResult = vbIPicture_GetWidth(UserData)
    Case 5: VResult = vbIPicture_GetHeight(UserData)
    Case 6: vbIPicture_Render UserData, P(0), P(1), P(2), P(3), P(4), P(5), P(6), P(7), P(8), P(9)
    Case Else 'we don't support a 'Member' (a Property or Function) with a DispID-Index that differs from 0,2,3,4,5,6
      vbIDispatch_Invoke = DISP_E_MEMBERNOTFOUND '<- and so we return the appropriate HRESULT
  End Select
End Function


'************* vbIPicture-Implementation ****************
Private Function vbIPicture_GetHandle(UserData As Long) As Long
  If mHIcon = 0 And mPictureType = vbPicTypeIcon And mIsCursor Then mHIcon = GDIPlusCache.GetHCursorFromImage(mImgKey, mHSpotX, mHSpotY)
  If mHIcon = 0 And mPictureType = vbPicTypeIcon And Not mIsCursor Then mHIcon = GDIPlusCache.GetHIconFromImage(mImgKey)
  If mHBmp = 0 And mPictureType = vbPicTypeBitmap Then mHBmp = GDIPlusCache.GetHBmpFromImage(mImgKey, vbWhite)
  If mPictureType = vbPicTypeIcon Then vbIPicture_GetHandle = mHIcon
  If mPictureType = vbPicTypeBitmap Then vbIPicture_GetHandle = mHBmp
End Function
Private Function vbIPicture_GetHPal(UserData As Long) As Long
End Function

Private Function vbIPicture_GetPictureType(UserData As Long) As PictureTypeConstants
  vbIPicture_GetPictureType = mPictureType
End Function

Private Function vbIPicture_GetWidth(UserData As Long) As Long
  vbIPicture_GetWidth = GDIPlusCache.Width(mImgKey) * HM2PX_96DPI
End Function

Private Function vbIPicture_GetHeight(UserData As Long) As Long
  vbIPicture_GetHeight = GDIPlusCache.Height(mImgKey) * HM2PX_96DPI
End Function

Private Sub vbIPicture_Render(UserData As Long, ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, _
                              ByVal xSrc As Long, ByVal ySrc As Long, ByVal cxSrc As Long, ByVal cySrc As Long, ByVal pRcBounds As Long)
  xSrc = xSrc / HM2PX_96DPI
  ySrc = GDIPlusCache.Height(mImgKey) - (ySrc + 1) / HM2PX_96DPI
  cxSrc = cxSrc / HM2PX_96DPI
  cySrc = -cySrc / HM2PX_96DPI
  
  GDIPlusCache.AlphaRenderTo hDC, mImgKey, x, y, cx, cy, xSrc, ySrc, cxSrc, cySrc, mAlpha, mHighStretchQuality
End Sub
 
Private Sub vbIPicture_SetHPal(UserData As Long, ByVal NewHPal As Long)
End Sub
Private Function vbIPicture_GetCurDC(UserData As Long) As Long
End Function
Private Sub vbIPicture_SelectPicture(UserData As Long, ByVal hDCToSelectInto As Long, hDCPrevious As Long, hBmp As Long)
End Sub
Private Sub vbIPicture_PictureChanged(UserData As Long)
End Sub
Private Sub vbIPicture_SaveAsFile(UserData As Long, ByVal pStm As Long, ByVal fSaveMemCopy As Boolean, SavedBytes As Long)
End Sub

Private Function vbIPicture_GetAttributes(UserData As Long) As Long
  vbIPicture_GetAttributes = mPicAttributes
End Function

