VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGDIPlusCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'a GDI+ based Cache-Class for Ico- and Png-Alpha-Resources, later accessible per String-Key...(think: ImageList-Replacement)
'it can be used standalone (also without cPictureEx, then able to render Alpha-Content onto any hDC) ... [Olaf Schmidt 2015]
Option Explicit

Private Const PixelFormat32bppPARGB = &HE200B
 
Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, Inbuf As Long, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As Long, Image As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal Image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal Image As Long, Height As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Img As Long, Context As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal Context As Long, ByVal Image As Long, ByVal dstx As Long, ByVal dsty As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As Long, ByVal imageAttributes As Long, ByVal callback As Long, ByVal callbackData As Long) As Long
Private Declare Function GdipSetClipRectI Lib "gdiplus" (ByVal Context As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal CombineMode As Long) As Long
Private Declare Function GdipGraphicsClear Lib "gdiplus" (ByVal Context As Long, ByVal Color As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal Context As Long, ByVal Mode As Long) As Long
Private Declare Function GdipBitmapConvertFormat Lib "gdiplus" (ByVal Image As Long, ByVal Format As Long, ByVal ADitherType As Long, ByVal APaletteType As Long, ByVal pPalette As Long, ByVal AlphaThresholdPercent As Single) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal dx As Long, ByVal dy As Long, ByVal stride As Long, ByVal PixelFormat As Long, ByVal pScanData As Long, Image As Long) As Long
Private Declare Function GdipCreateHICONFromBitmap Lib "gdiplus" (ByVal Image As Long, hIcon As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal Image As Long, hBmp As Long, ByVal BGColor As Long) As Long
Private Declare Function GdipImageGetFrameCount Lib "gdiplus" (ByVal Image As Long, FrDimID As Any, Count As Long) As Long
Private Declare Function GdipImageSelectActiveFrame Lib "gdiplus" (ByVal Image As Long, FrDimID As Any, ByVal FrIdx As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal Context As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
 
Private Declare Function CreateIconFromResourceEx Lib "user32" (presbits As Any, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal Flags As Long) As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon&, IconInfo As Any) As Long
Private Declare Function CreateIconIndirect Lib "user32" (IconInfo As Any) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC&, ByVal x&, ByVal y&, ByVal hIcon&, ByVal dx&, ByVal dy&, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long

Private Declare Function GdiAlphaBlend& Lib "gdi32" (ByVal hDC&, ByVal x&, ByVal y&, ByVal dx&, ByVal dy&, ByVal hdcSrc&, ByVal srcx&, ByVal srcy&, ByVal SrcdX&, ByVal SrcdY&, ByVal lBlendFunction&)
Private Declare Function CreateCompatibleDC& Lib "gdi32" (ByVal hDC&)
Private Declare Function CreateDIBSection& Lib "gdi32" (ByVal hDC&, pBitmapInfo As Any, ByVal un&, ppBits&, ByVal Hdl&, ByVal dw&)
Private Declare Function SelectObject& Lib "gdi32" (ByVal hDC&, ByVal hObject&)
Private Declare Function DeleteObject& Lib "gdi32" (ByVal hObj&)
Private Declare Function DeleteDC& Lib "gdi32" (ByVal hDC&)

Private Declare Function CLSIDFromString Lib "ole32" (ByVal psGUID As Long, id As Any) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlob&, ByVal fDeleteOnRelease As Long, ppstm As stdole.IUnknown) As Long
 
Private mToken As Long, mImages As New Collection, mFrDim(0 To 15) As Byte
Private mhDC&, mOldBmp&, mBBCtx&, mBBImg&, mBBdx&, mBBdy& '<- Variables for the BackBuffer-Handling

Private Sub Class_Initialize()
Dim StartupInput&(0 To 3): StartupInput(0) = 1
    GdiplusStartup mToken, StartupInput(0)
    CLSIDFromString StrPtr("{6AEDBD6D-3FB5-418A-83A6-7F45229DC872}"), mFrDim(0) 'FrameDimensionTime (needed for animated Gifs)
End Sub

Private Sub Class_Terminate()
  Dim Img: For Each Img In mImages: GdipDisposeImage Img: Next
  DestroyBackBuf
  If mToken Then GdiplusShutdown mToken
End Sub

Private Sub PrepareAndClearBackBuf(ByVal dx As Long, ByVal dy As Long)
  If mBBdx < dx Or mBBdy < dy Then
    If mBBdx < dx Then mBBdx = dx
    If mBBdy < dy Then mBBdy = dy
    
    Dim BI(0 To 9) As Long, pDIB As Long
        BI(0) = 40
        BI(1) = mBBdx
        BI(2) = -mBBdy
        BI(3) = 32 * 65536 + 1 '32bpp
    If mhDC = 0 Then mhDC = CreateCompatibleDC(0)
    If mOldBmp Then
      DeleteObject SelectObject(mhDC, CreateDIBSection(0, BI(0), 0, pDIB, 0, 0))
    Else
      mOldBmp = SelectObject(mhDC, CreateDIBSection(0, BI(0), 0, pDIB, 0, 0))
    End If
    If mBBImg Then GdipDisposeImage mBBImg
    GdipCreateBitmapFromScan0 mBBdx, mBBdy, mBBdx * 4, PixelFormat32bppPARGB, pDIB, mBBImg
    If mBBImg Then GdipGetImageGraphicsContext mBBImg, mBBCtx
  End If
  GdipSetClipRectI mBBCtx, 0, 0, dx, dy, 0
  GdipGraphicsClear mBBCtx, 1
End Sub
 
Private Sub DestroyBackBuf()
  If mBBCtx Then GdipDeleteGraphics mBBCtx
  If mBBImg Then GdipDisposeImage mBBImg
  If mOldBmp Then DeleteObject SelectObject(mhDC, mOldBmp)
  If mhDC Then DeleteDC mhDC
End Sub
 
Private Function CreateScaledImg(SrcImg As Long, dxSrc, dySrc, dxDst, dyDst) As Long
Dim Img As Long, Ctx As Long
  GdipCreateBitmapFromScan0 dxDst, dyDst, dxDst * 4, PixelFormat32bppPARGB, 0, Img
  If Img Then CreateScaledImg = Img Else Err.Raise vbObjectError, , "unable to create scaled Img-Resource"
  If Img Then GdipGetImageGraphicsContext Img, Ctx
  If Ctx Then GdipSetInterpolationMode Ctx, 2
  If Ctx <> 0 And SrcImg <> 0 Then GdipDrawImageRectRectI Ctx, SrcImg, 0, 0, dxDst, dyDst, 0, 0, dxSrc, dySrc, 2, 0, 0, 0
  If Ctx Then GdipDeleteGraphics Ctx
End Function

Private Function ReadBytesFromFile(FileName) As Byte()
  With CreateObject("ADODB.Stream")
    .Open
      .Type = 1 'adTypeBinary
      .LoadFromFile FileName
      ReadBytesFromFile = .Read
    .Close
  End With
End Function

Public Function Exists(Key) As Boolean
On Error GoTo ExitFalse
  mImages.Item Key: Exists = True
ExitFalse:
End Function

Public Sub AddImage(Key, FileNameOrBytes, Optional ByVal DesiredWidth& = 0, Optional ByVal DesiredHeight& = 0, Optional ByVal CacheMultiFrames As Boolean)
Dim B() As Byte, Strm As stdole.IUnknown, Tmp&, dx&, dy&, FrameCount&, i&
  If VarType(FileNameOrBytes) = vbString Then B = ReadBytesFromFile(FileNameOrBytes) Else B = FileNameOrBytes

  CreateStreamOnHGlobal VarPtr(B(0)), 0, Strm
  GdipLoadImageFromStream ObjPtr(Strm), Tmp
  GdipGetImageWidth Tmp, dx:  If DesiredWidth <= 0 Then DesiredWidth = dx
  GdipGetImageHeight Tmp, dy: If DesiredHeight <= 0 Then DesiredHeight = dy
  
  mImages.Add CreateScaledImg(Tmp, dx, dy, DesiredWidth, DesiredHeight), Key
  If Tmp <> 0 And CacheMultiFrames Then
    GdipImageGetFrameCount Tmp, mFrDim(0), FrameCount
    For i = 0 To FrameCount - 1
      GdipImageSelectActiveFrame Tmp, mFrDim(0), i
      mImages.Add CreateScaledImg(Tmp, dx, dy, DesiredWidth, DesiredHeight), Key & "|" & i
    Next
  End If
  
  GdipDisposeImage Tmp
End Sub

Public Sub AddIcon(Key, FileNameOrBytes, ByVal DesiredWidth As Long, ByVal DesiredHeight As Long)
Dim B() As Byte, i&, MaxW%(1 To 256), sz&, Offs&, hIcon&
  If VarType(FileNameOrBytes) = vbString Then B = ReadBytesFromFile(FileNameOrBytes) Else B = FileNameOrBytes
  
  For i = 0 To B(4) - 1
    If B(12 + i * 16) = 32 Then MaxW((511 + B(6 + i * 16)) Mod 256 + 1) = i + 1
  Next
  For sz = 1 To 256
    If MaxW(sz) Then Offs = (MaxW(sz) - 1) * 16 + 18
    If sz >= DesiredWidth Then Exit For
  Next
  If Offs = 0 Then Err.Raise vbObjectError, , "the Icon-Resource doesn't contain any 32Bit-Alpha-Icons"
  Offs = B(Offs) + 256& * B(Offs + 1) + 65536 * B(Offs + 2)
  hIcon = CreateIconFromResourceEx(B(Offs), UBound(B) + 1 - Offs, 1, &H30000, sz, sz, 0)

  PrepareAndClearBackBuf sz, sz
  If hIcon Then DrawIconEx mhDC, 0, 0, hIcon, sz, sz, 0, 0, 3
  If hIcon Then DestroyIcon hIcon 'not needed anymore, so we destroy it here

  mImages.Add CreateScaledImg(mBBImg, sz, sz, DesiredWidth, DesiredHeight), Key
End Sub

Public Property Get FrameCount(Key) As Long
  Do While Exists(Key & "|" & FrameCount): FrameCount = FrameCount + 1: Loop
End Property
 
Public Property Get Width(Key) As Long
  GdipGetImageWidth mImages(Key), Width
End Property
Public Property Get Height(Key) As Long
  GdipGetImageHeight mImages(Key), Height
End Property

Public Function GetHIconFromImage(Key) As Long
  GdipCreateHICONFromBitmap mImages(Key), GetHIconFromImage
End Function
Public Sub DestroyHIcon(ByVal hIcon As Long)
  If hIcon Then DestroyIcon hIcon
End Sub

Public Function GetHCursorFromImage(Key, Optional ByVal HSpotX As Long, Optional ByVal HSpotY As Long) As Long
Dim hIcon As Long, IconInfo(0 To 4) As Long
    GdipCreateHICONFromBitmap mImages(Key), hIcon
    If hIcon Then GetIconInfo hIcon, IconInfo(0)
    IconInfo(0) = 0: IconInfo(1) = HSpotX: IconInfo(2) = HSpotY
    GetHCursorFromImage = CreateIconIndirect(IconInfo(0))
    DeleteObject IconInfo(3)
    DeleteObject IconInfo(4)
    DestroyIcon hIcon
End Function
Public Sub DestroyHCursor(ByVal hCursor As Long)
  If hCursor Then DestroyCursor hCursor
End Sub

Public Function GetHBmpFromImage(Key, Optional ByVal BGColor As Long = vbWhite) As Long
  GdipCreateHBITMAPFromBitmap mImages(Key), GetHBmpFromImage, BGColor
End Function
Public Sub DestroyHBmp(ByVal hBmp As Long)
  If hBmp Then DeleteObject hBmp
End Sub
 
Public Sub AlphaRenderTo(ByVal hDC As Long, Key, Optional ByVal x As Long, Optional ByVal y As Long, _
                                                 Optional ByVal dx As Long, Optional ByVal dy As Long, _
                                                 Optional ByVal xSrc As Long, Optional ByVal ySrc As Long, _
                                                 Optional ByVal dxSrc As Long, Optional ByVal dySrc As Long, _
                                                 Optional ByVal Alpha As Double = 1, Optional ByVal HighStretchQuality As Boolean)
  If dxSrc = 0 Then dxSrc = Width(Key)
  If dySrc = 0 Then dySrc = Height(Key)
  If dx = 0 Then dx = Width(Key)
  If dy = 0 Then dy = Height(Key)

  PrepareAndClearBackBuf dx, dy
  GdipSetInterpolationMode mBBCtx, IIf(HighStretchQuality, 2, 1)
  GdipDrawImageRectRectI mBBCtx, mImages(Key), 0, 0, dx, dy, xSrc, ySrc, dxSrc, dySrc, 2, 0, 0, 0
  GdiAlphaBlend hDC, x, y, dx, dy, mhDC, xSrc, ySrc, dx, dy, &H1000000 + &H10000 * Int(255 * Alpha)
End Sub
