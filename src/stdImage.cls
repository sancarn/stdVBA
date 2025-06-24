VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@module
'@description A class which implements a wrapper around an image. Use this class to
'read image data to and from files, clipboard, sheets, UI controls, and more. Additionally
'this class allows the manipulation of images in memory, and perform analytical operations on them.
'@example Simple example of reading an image from a file and saving it to the clipboard:
'```vb
'Dim img as stdImage: set img = stdImage.CreateFromFile("H:\test1.bmp")
'img.ToClipboard()
'''


'Developer notes:
' Useful links:
' * Jaafar
'   * https://www.mrexcel.com/board/threads/working-with-images-in-vba-displaying-png-files.1224206/post-5991836
'     * Demonstrates Bytes -> Stdole.StdPicture using GDIP
'     * Demonstrates hBITMAP -> Bytes
' * Dan_W
'   * https://www.mrexcel.com/board/threads/working-with-images-in-vba-displaying-png-files.1224206/#post-5988472
'     * Demonstrates stdole.IPicture -> Bytes
'   * https://www.mrexcel.com/board/threads/working-with-images-in-vba-displaying-png-files.1224206/#post-5987775
'     * Demonstrates Bytes  -> Stdole.StdPicture using WIA
'     * Demonstrates Base64 -> Stdole.StdPicture using WIA
'     * Demonstrates URL    -> Stdole.StdPicture using WIA
'   * https://www.mrexcel.com/board/threads/working-with-images-in-vba-displaying-png-files.1224206/#post-5987712
'     * Demonstrates File   -> Stdole.Stdpicture using WIA (including png, bmp, gif, ...)
' * Lots of useful stuff here: https://www.vbforums.com/showthread.php?630193-vb6-Alpha-Image-Control-v2-Final-Update-(15-Jan-2012)
' * GdipSaveImageToFile stuff - https://www.vbforums.com/showthread.php?761713-Take-A-screenshot-every-x-minutes-and-upload-to-ftp
'Mostly Wraps GDI+

'Scope:
'stdImage::CreateFromFile(".../test.png")
'stdImage::CreateFromFile(".../test.jpg")
'stdImage::CreateFromFile(".../test.tif")
'stdImage::CreateFromFile(...)
'stdImage::CreateFromShape(shp as Shape)
'stdImage::Create()                                                                                             'Create blank canvas - can be used for charts
'stdImage::CreateFromPicture(...)                                                                               'Create from StdOle.StdPicture
'stdImage::CreateFromPictureDisp(...)                                                                           'Create from StdOle.StdOle.IPictureDisp
'stdImage::CreateFromClipboard()                                                                                'Create image from clipboard
'stdImage::CreateFromHICON(...)                                                                                 '?? Create from HICON
'stdImage::CreateFromSVG()                                                                                      '?? Create picture from SVG
'<#stdImage>.ToClipboard()                                                                                      'Copies to clipboard
'<#stdImage>.ToFile(sFilePath as string)                                                                        'Exports to file
'<#stdImage>.ToUIComponent()                                                                                    'Converts to UserForm Component
'<#stdImage>.ToDataURL()                                                                                        'Converts picture to data url as string
'<#stdImage>.ToStdPicture                                                                                       'Returns StdOle.StdPicture
'<#stdImage>.ToPictureDisp                                                                                      'Returns StdOle.IPictureDisp
'<#stdImage>.HIcon                                                                                              '?? Returns HIcon
'***********************
'* NOT YET IMPLEMENTED *
'***********************
'<#stdImage>.Draw(x as long, y as long, rgbColor as long)                                                       'Draw pixel
'<#stdImage>.DrawLine(x1 as long, y1 as long, x2 as long, y2 as long, rgbColor as long)                         'Draw Line
'<#stdImage>.DrawEllipse(x as long, y as long, w as long, h as long, rgbColor as long)                          'Draw Ellipse
'<#stdImage>.DrawPolyline(points() as Point, rgbLineColor as long, iLineThickness as long)                      'Draw a polyline
'<#stdImage>.DrawPolygon(points() as Point, rgbLineColor as long, iLineThickness as long, rgbFillColor as long) 'Draw a polygon
'<#stdImage>.DrawRect(x,y,w,h)                                                                                  'Draw a rectangle
'<#stdImage>.DrawRoundRect(x,y,w,h,rw,rh)                                                                       'Draw a rectangle with rounded corners
'<#stdImage>.DrawArc(...)                                                                                       'Draw an arc
'<#stdImage>.DrawChord(...)                                                                                     'Draw a chord
'<#stdImage>.DrawPie(...)                                                                                       'Draw a pie
'<#stdImage>.DrawPolyBezier(...)                                                                                'Draws cubic Bézier curves by using the endpoints and control points specified by the lppt parameter.
'<#stdImage>.Resize()
'<#stdImage>.Crop()
'
'************
'* EXAMPLES *
'************
'```vba
'Private Sub UserForm_Initialize()
'  With stdWindow.CreateFromIUnknown(Me)
'    Call .setOwnerHandle(0)
'    .HICON = stdWindow.CreateFromHwnd(Application.VBE.MainWindow.hWnd).HICON       'Set window icon from existing window
'    .HICON = stdImage.CreateFromStdPicture(Image1.picture).HICON                   'Set window icon from Image control content on form
'    .HICON = stdImage.CreateFromShape(Sheet1.Shapes("Picture 2")).HICON            'Set window icon from shape on worksheet
'    .HICON = stdImage.CreateFromFile("C:\Users\sancarn\Pictures\yuumi.png").HICON  'Set window icon from file
'
'    Image1.PictureSizeMode = fmPictureSizeModeStretch
'    Set Image1.picture = stdImage.CreateFromShape(Sheet1.Shapes("Picture 2")).ToStdPicture             'Set image control content from shape on worksheet
'    Set Image1.picture = stdImage.CreateFromFile("C:\Users\sancarn\Pictures\yuumi.png").ToStdPicture   'Set image control content from file
'  End With
'End Sub
'```
'---------------------
'```vba
'Sub SaveShapeToFile()
'  Call stdImage.CreateFromShape(Sheet1.Shapes("Picture 2")).ToFile("C:\Temp\poop.png")
'  Call stdImage.CreateFromShape(Sheet1.Shapes("Picture 2")).ToFile("C:\Temp\poop.bmp")
'  Call stdImage.CreateFromShape(Sheet1.Shapes("Picture 2")).ToFile("C:\Temp\poop.jpeg")
'  Call stdImage.CreateFromShape(Sheet1.Shapes("Picture 2")).ToFile("C:\Temp\poop.tiff")
'End Sub
'```

Public Enum stdImgFormat
  stdImgFormatDefault
  stdImgFormatBMP
  stdImgFormatPNG
  stdImgFormatGIF
  stdImgFormatJPEG
  stdImgFormatTIFF
End Enum

#If Win64 Then
  Private Const NULL_PTR = 0^
#Else
  Private Const NULL_PTR = 0&
#End If

'Clipboard Management
#If VBA7 Then
  Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Integer) As Long
  Private Declare PtrSafe Function OpenClipboard Lib "user32" (Optional ByVal hWnd As LongPtr = NULL_PTR) As Long
  Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
  Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
  Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Integer) As LongPtr

  Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As LongPtr
  Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
  Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
  Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As Long
  Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
  
  'Obtain image bytes from clipboard
  Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
  Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As LongPtr
  Private Declare PtrSafe Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As LongPtr, ByVal nCount As Long, lpObject As Any) As Long
  Private Declare PtrSafe Function CreateDIBSection Lib "gdi32" (ByVal hDC As LongPtr, pbmi As BITMAPINFO, ByVal iUsage As Long, ByVal ppvBits As LongPtr, ByVal hSection As LongPtr, ByVal dwOffset As Long) As LongPtr
  Private Declare PtrSafe Function GetDIBits Lib "gdi32" (ByVal aHDC As LongPtr, ByVal hBitmap As LongPtr, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
  Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
  Private Declare PtrSafe Function CreateIconFromResourceEx Lib "user32" (ByVal pbIconBits As LongPtr, ByVal cbIconBits As Long, ByVal fIcon As Long, ByVal dwVersion As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal uFlags As Long) As LongPtr
  Private Declare PtrSafe Function CreateIconIndirect Lib "user32" (ByRef piconinfo As ICONINFO) As LongPtr
  Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
  Private Declare PtrSafe Function DestroyIcon Lib "user32" (ByVal hIcon As LongPtr) As Long
  Private Declare PtrSafe Function DrawIconEx Lib "user32" (ByVal hdc As LongPtr, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As LongPtr, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
  Private Declare PtrSafe Function GetIconInfo Lib "user32" (ByVal hIcon As LongPtr, ByRef piconinfo As ICONINFO) As Long

  Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As Long)
  Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long

  Private Declare PtrSafe Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As LongPtr, ByVal offsetinVft As LongPtr, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As VbVarType, ByRef paValues As LongPtr, ByRef retVAR As Variant) As Long

  Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" (token As LongPtr, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As LongPtr = 0) As Long
  Private Declare PtrSafe Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal FILENAME As LongPtr, BITMAP As LongPtr) As Long
  Private Declare PtrSafe Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal BITMAP As LongPtr, hbmReturn As LongPtr, ByVal background As LongPtr) As Long
  Private Declare PtrSafe Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hBitmap As LongPtr, ByVal hPalette As LongPtr, gdipBitmap As LongPtr) As Long
  Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal Image As LongPtr) As Long
  Private Declare PtrSafe Sub GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr)
  Private Declare PtrSafe Function GdipSaveImageToFile Lib "gdiplus" (ByVal Image As LongPtr, ByVal Filename As LongPtr, ByRef clsidEncoder As GUID, ByRef encoderParams As Any) As Long
  Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long
  Private Declare PtrSafe Function OleTranslateColor Lib "oleAut32.dll" (ByVal clr As OLE_COLOR, ByVal palet As LongPtr, Col As Long) As Long
  Private Declare PtrSafe Function IIDFromString Lib "ole32" (ByVal lpsz As LongPtr, ByRef lpiid As GUID) As Long

  'Screenshot
  Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
  Private Declare PtrSafe Function GetWindowDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
  Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
  Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
  Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr
  Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
  Private Declare PtrSafe Function BitBlt Lib "gdi32" (ByVal hDestDC As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
  Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hdc As LongPtr) As Long
#Else
  Enum LongPtr
    [_]
  End Enum

  Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Integer) As Long
  Private Declare Function OpenClipboard Lib "user32" (Optional ByVal hwnd As LongPtr = NULL_PTR) As Long
  Private Declare Function CloseClipboard Lib "user32" () As Long
  Private Declare Function EmptyClipboard Lib "user32" () As Long
  Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Integer) As LongPtr

  Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As LongPtr
  Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
  Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
  Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As Long
  Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
  
  'Obtain image bytes from clipboard
  Private Declare Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
  Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As LongPtr
  Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As LongPtr, ByVal nCount As Long, lpObject As Any) As Long
  Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As LongPtr, pbmi As BITMAPINFO, ByVal iUsage As Long, ByVal ppvBits As LongPtr, ByVal hSection As LongPtr, ByVal dwOffset As Long) As LongPtr
  Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As LongPtr, ByVal hBitmap As LongPtr, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
  Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
  Private Declare Function CreateIconFromResourceEx Lib "user32" (ByVal pbIconBits As LongPtr, ByVal cbIconBits As Long, ByVal fIcon As Long, ByVal dwVersion As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal uFlags As Long) As LongPtr
  Private Declare Function CreateIconIndirect Lib "user32" (ByRef piconinfo As ICONINFO) As LongPtr
  Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
  Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As LongPtr) As Long
  Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As LongPtr, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As LongPtr, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
  Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As LongPtr, ByRef piconinfo As ICONINFO) As Long

  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As Long)
  Private Declare Function GetTickCount Lib "kernel32" () As Long

  'FIX: The use of `VbVarType` for the type of `paTypes` on 32 bit causes Invalid Callee error. Workaround is to use `Integer` instead as below.
  Private Declare Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long

  Private Declare Function GdiplusStartup Lib "gdiplus" (token As LongPtr, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As LongPtr = 0) As Long
  Private Declare Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal FILENAME As LongPtr, BITMAP As LongPtr) As Long
  Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal BITMAP As LongPtr, hbmReturn As LongPtr, ByVal background As LongPtr) As Long
  Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hBitmap as LongPtr, ByVal hPalette as LongPtr, gdipBitmap as LongPtr) As Long
  Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As LongPtr) As Long
  Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr)
  Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal Image As LongPtr, ByVal Filename As LongPtr, ByRef clsidEncoder As GUID, ByRef encoderParams As Any) As Long
  Private Declare Function OleCreatePictureIndirect Lib "oleaut32" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long
  Private Declare Function OleTranslateColor Lib "oleAut32.dll" (ByVal clr As OLE_COLOR, ByVal palet As LongPtr, Col As Long) As Long
  Private Declare Function IIDFromString Lib "ole32" (ByVal lpsz As LongPtr, ByRef lpiid As GUID) As Long
  Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

  'Screenshot
  Private Declare Function GetDesktopWindow Lib "user32" () As LongPtr
  Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
  Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
  Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
  Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr
  Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
  Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
  Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As LongPtr) As Long
#End If
Private Const CF_BITMAP = 2



Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type BITMAP
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As LongPtr
End Type

Private Type BITMAPFILEHEADER
  bfType As String * 2&
  bfSize As Long
  bfReserved1 As Integer
  bfReserved2 As Integer
  bfOffBits As Long
End Type

Private Type BITMAPINFOHEADER
  biSize As Long
  biWidth As Long
  biHeight As Long
  biPlanes As Integer
  biBitCount As Integer
  biCompression As Long
  biSizeImage As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed As Long
  biClrImportant As Long
End Type

Private Type BITMAPINFO
  bmiheader As BITMAPINFOHEADER
End Type

'Declare a UDT to store the bitmap information
Private Type PICTDESC
    size As Long
    Type As Long
    hPic As LongPtr
    hPal As LongPtr
End Type

'Declare a UDT to store the GDI+ Startup information
Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As LongPtr
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type GUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(0 To 7) As Byte
End Type

Private Type EncoderParameter
  GUID As GUID
  NumberOfValues As Long
  type As Long
  Value As LongPtr
End Type

Private Type EncoderParameters
  count As Long
  Parameter() As EncoderParameter
End Type

'Before:
'Private Type EncoderParameter
'  GUID As GUID
'  NumberOfValues As Long
'  type As Long
'  Value As Long
'End Type
'
'Private Type EncoderParameters
'  count As Long
'  Parameter As EncoderParameter
'End Type

Private Type ICONINFO
  fIcon As Long       ' Specifies whether this structure defines an icon or a cursor. A value of 1 specifies an icon; 0 specifies a cursor.
  xHotspot As Long
  yHotspot As Long
  hbmMask As LongPtr  'HBITMAP
  hbmColor As LongPtr 'HBITMAP
End Type


'***********
'*  State  *
'***********
Private Type TThisCache
  HICON As LongPtr
End Type
Private Type TThis
  HBITMAP As LongPtr
  HBITMAPTYPE As BITMAP
  Cache As TThisCache
End Type
Private This As TThis


'****************
'* Constructors *
'****************

'Creates a stdImage object from a byte array
'@constructor
'@TODO: Awaiting implementation
'@param b - Bytes of Bitmap/Jpeg/Png/... image
'@param format - The format of the image
'@returns - Picture object representing data
'@example ```vb
'  Dim b() as Byte: Redim b(1 to 100)
'  Call stdImage.CreateFromBinary(b).ToFile("H:\test1.png")
'```
Public Function CreateFromBinary(ByRef b() As Byte, ByVal format As stdImgFormat) As stdImage
  'TODO: Create Stream, Write to stream, CreateBitmapFromStream
  Err.Raise 1, , "Not implemented"
End Function

'@constructor
'@TODO: Awaiting implementation
Public Function CreateFromDataURL(ByVal sURL As String) As stdImage
  'TODO: Parse Base64, CreateFromBinary
  Err.Raise 1, , "Not implemented"
End Function

'Creates a stdImage object from a file
'@constructor
'@param sPath - Path of file
'@param format - The format of the image
'@returns - Picture object representing data
'@example ```vb
'stdImage.CreateFromFile("H:\test1.bmp").ToFile("H:\test1.png")
'```
Public Function CreateFromFile(ByVal sPath As String, Optional ByVal format As stdImgFormat = stdImgFormatDefault, Optional ByVal transparentColor As Long = 0) As stdImage
  If format = stdImgFormatDefault Then
    Dim pathParts: pathParts = Split(sPath, ".")
    Dim ext As String: ext = pathParts(UBound(pathParts))
    Select Case LCase(ext)
      Case "png": format = stdImgFormat.stdImgFormatPNG
      Case "bmp": format = stdImgFormat.stdImgFormatBMP
      Case "gif": format = stdImgFormat.stdImgFormatGIF
      Case "jpeg", "jpg": format = stdImgFormat.stdImgFormatJPEG
      Case "tiff": format = stdImgFormat.stdImgFormatTIFF
      Case "webp": Err.Raise 5, "stdImage::CreateFromFile", "WEBP format is not currently supported"
      Case "svg": Err.Raise 5, "stdImage::CreateFromFile", "SVG format is not currently supported"
      Case Else
        Err.Raise 1, "stdImage::CreateFromFile", "Could not determine format from file extension"
    End Select
  End If
  Select Case format
    'TODO: Case webp
    Case Else 'BMP, GIF, JPEG, PNG, TIFF, Exif, WMF, and EMF - Load using GDI+
      Dim uGdiInput As GdiplusStartupInput
      uGdiInput.GdiplusVersion = 1
      Dim hGDIPlus As LongPtr
      Dim lResult As Long: lResult = GdiplusStartup(hGDIPlus, uGdiInput)
      If lResult = 0 Then
        Dim hGDIImage As LongPtr
        lResult = GdipCreateBitmapFromFile(StrPtr(sPath), hGDIImage)
        If lResult = 0 Then
          Dim hBitmap As LongPtr
          lResult = GdipCreateHBITMAPFromBitmap(hGDIImage, hBitmap, BGR(transparentColor))
          If lResult = 0 Then
            Set CreateFromFile = CreateFromHBitmap(hBitmap)
          Else
            Err.Raise 1, "stdImage::CreateFromFile", "Could not convert GDI+ image to HBITMAP (Error Result: " & lResult & ")"
          End If

          Call GdipDisposeImage(hGDIImage)
        Else
          Err.Raise 1, "stdImage::CreateFromFile", "Could not load image (Error Result: " & lResult & ")"
        End If

        Call GdiplusShutdown(hGDIPlus)
      Else
        Err.Raise 1, "stdImage::CreateFromFile", "Could not start GDI+ (Error Result: " & lResult & ")"
      End If
  End Select
End Function

'Creates a stdImage object from a url
'@constructor
'@param sUrl - Url of file
'@param format - The format of the image
'@returns - Picture object representing data
'@example ```vb
'stdImage.CreateFromURL("https://raw.githubusercontent.com/sancarn/awesome-vba/main/resources/VBALogo.png").ToFile("H:\test.png", stdImgFormatPNG)
'```
'@remark - TODO: consider removing
Public Function CreateFromUrl(ByVal sURL As String, Optional ByVal format As stdImgFormat = stdImgFormatDefault) As stdImage
  Set CreateFromUrl = New stdImage
  Dim HTTP As Object: Set HTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
  HTTP.Open "GET", sURL, False
  'TODO: AutoProxy?
  HTTP.send
  If HTTP.status = 200 Then
    Dim body() As Byte: body = HTTP.responseBody
    If format = stdImgFormatDefault Then
      'TODO: Detect format
    End If
    Set CreateFromUrl = CreateFromBinary(body, format)
  Else
    Err.Raise 1, "", "Could not get file (HTTP Status: " & HTTP.status & ")"
  End If
End Function

'Creates a stdImage object from a file
'@constructor
'@param hBitmap - hBitmap (handle to Bitmap)
'@returns - Picture object representing data
'@example ```vb
'  stdImage.CreateFromHBitmap(Image1.Picture.Handle).ToFile("H:\test.png", stdImgFormatPNG)
'```
Public Function CreateFromHBitmap(ByVal hBitmap As LongPtr) As stdImage
  Dim bmp As BITMAP
  If GetObjectAPI(hBitmap, LenB(bmp), bmp) <> 0 Then
    Set CreateFromHBitmap = New stdImage
    Call CreateFromHBitmap.protInit(hBitmap)
  Else
    Err.Raise 1, "stdImage::CreateFromHBitmap", "The provided hBITMAP is not valid"
  End If
End Function

'Creates a stdImage from the image in the clipboard
'@constructor
'@returns - Picture object representing data
'@example ```vb
'stdImage.CreateFromClipboard().ToFile("H:\test.png", stdImgFormatPNG)
'```
Public Function CreateFromClipboard() As stdImage
  'Open clipboard
  If Not CBool(OpenClipboardTimeout(getAppHwnd())) Then Call Err.Raise(1, "Value [GET]", "Unable to open clipboard (DllError: " & Err.LastDllError & ")")
  
  'Check image in clipboard
  If CBool(IsClipboardFormatAvailable(CF_BITMAP)) Then
    Dim hBitmap As LongPtr: hBitmap = GetClipboardData(CF_BITMAP)  'Obtain image handle
    Set CreateFromClipboard = CreateFromHBitmap(hBitmap)
  Else
    Dim sErrorText As String: sErrorText = "No bitmap format in clipboard"
  End If
  
  'Close clipboard
  If Not CBool(CloseClipboard()) Then Call Err.Raise(1, "Value [GET]", "Unable to close clipboard (DllError: " & Err.LastDllError & ")")
  If sErrorText <> "" Then Call Err.Raise(1, "Value [GET]", sErrorText)
End Function

'Creates a stdImage object from a stdole.IPictureDisp object
'@constructor
'@param picture - Object to create picture from
'@returns - Picture object representing data
'@example TODO:
Public Function CreateFromIPictureDisp(ByVal picture As stdole.IPictureDisp) As stdImage
  Set CreateFromIPictureDisp = CreateFromHBitmap(picture.Handle)
End Function

'Creates a stdImage object from a stdole.stdPicture object
'@constructor
'@param picture - Object to create picture from
'@returns - Picture object representing data
'@example ```vb
'stdImage.CreateFromStdPicture(Userform1.Picture).ToFile("H:\test.png", stdImgFormatPNG)
'```
Public Function CreateFromStdPicture(ByVal picture As stdole.StdPicture) As stdImage
  Set CreateFromStdPicture = CreateFromHBitmap(picture.Handle)
End Function

'Creates a stdImage object from a stdole.IPicture object
'@constructor
'@param picture - Object to create picture from
'@returns - Picture object representing data
'@example TODO:
Public Function CreateFromIPicture(ByVal picture As stdole.IPicture) As stdImage
  Set CreateFromIPicture = CreateFromHBitmap(picture.Handle)
End Function

'Creates a stdImage object from a Excel.IPicture object
'@constructor
'@param picture as Object<Excel.IPicture> - Object to create picture from
'@returns - Picture object representing data
'@example TODO:
Public Function CreateFromExcelIPicture(ByVal picture As Object) As stdImage
  Call picture.CopyPicture(, xlBitmap)
  Set CreateFromExcelIPicture = CreateFromClipboard()
End Function

'Creates a stdImage object from a Shape object
'@constructor
'@param shp as Object<(Excel|Word|PowerPoint).Shape> - Object to create picture from
'@returns - Picture object representing data
'@example ```vb
'stdImage.CreateFromShape(Sheet1.Shapes("Picture 2")).ToFile("H:\test.png")
'```
Public Function CreateFromShape(ByVal shp As Object) As stdImage
  Select Case Application.Name
    Case "Microsoft Excel"
      Dim xlShp As Excel.Shape: Set xlShp = shp
      Call xlShp.CopyPicture(, XlCopyPictureFormat.xlBitmap)
      Set CreateFromShape = stdImage.CreateFromClipboard()
  End Select
End Function

'Creates a stdImage object from a Chart object
'@constructor
'@param chart as Object<(Excel|Word|PowerPoint).Chart> - Object to create picture from
'@returns - Picture object representing data
'@example ```vb
'stdImage.CreateFromChart(Sheet1.ChartObjects(1).Chart).ToFile("H:\test.png")
'```
Public Function CreateFromChart(ByVal chart As Object) As stdImage
  Select Case Application.Name
    Case "Microsoft Excel"
      Dim xlChart As Excel.Chart: Set xlChart = chart
      Call xlChart.CopyPicture(, XlCopyPictureFormat.xlBitmap)
      Set CreateFromChart = stdImage.CreateFromClipboard()
  End Select
End Function

'Creates a stdImage object from the desktop window (screenshot)
'@constructor
'@returns - Picture object containing screenshot data
'@example ```vb
'stdImage.CreateFromScreen().ToFile("H:\test.png")
'```
'@TODO: Capture both monitors https://stackoverflow.com/a/34445431/6302131
Public Function CreateFromScreen() As stdImage
  Dim hDesktop As LongPtr: hDesktop = GetDesktopWindow()
  Set CreateFromScreen = CreateFromWindow(hDesktop)
End Function

'Creates a stdImage object from a window
'@constructor
'@param hwnd - Handle to window
'@param bClientArea - If true, only the client area of the window is captured. If false, the entire window is captured.
'@returns - Picture object representing data
'@TODO: Option to include cursor https://stackoverflow.com/a/48925443/6302131
#If VBA7 Then
  Public Function CreateFromWindow(ByVal hwnd As LongPtr, Optional ByVal bClientArea As Boolean = False) As stdImage
#Else
  Public Function CreateFromWindow(ByVal hwnd as Long, Optional ByVal bClientArea as Boolean = false) as stdImage
#End If
  Const SRCCOPY As Long = &HCC0020
  
  'Get the proper Device Context (DC) depending on the value of bClient.
  'Additionally get the rect of the window/client.
  Dim hDCSource As LongPtr, tRect As RECT
  If bClientArea Then
    hDCSource = GetDC(hwnd)
    Call GetClientRect(hwnd, tRect)
  Else
    hDCSource = GetWindowDC(hwnd)
    Call GetWindowRect(hwnd, tRect)
  End If

  'Create an in-memory DC for the copy process
  Dim hCaptureDC As LongPtr: hCaptureDC = CreateCompatibleDC(hDCSource)

  'Create a bitmap to capture the window
  Dim hCaptureBitmap As LongPtr: hCaptureBitmap = CreateCompatibleBitmap(hDCSource, tRect.Right - tRect.Left, tRect.Bottom - tRect.Top)
  
  'Load bitmap into in-memory DC
  Dim hOldBitmap As LongPtr: hOldBitmap = SelectObject(hCaptureDC, hCaptureBitmap)

  'Copy the on-screen image into the memory DC.
  Dim r As Long: r = BitBlt(hCaptureDC, 0, 0, tRect.Right - tRect.Left, tRect.Bottom - tRect.Top, hDCSource, 0, 0, SRCCOPY)

  'Move old bitmap back into capture DC
  Call SelectObject(hCaptureDC, hOldBitmap)
  Call ReleaseDC(hwnd, hDCSource)
  Call DeleteDC(hCaptureDC)

  'Create stdImage from bitmap
  Set CreateFromWindow = CreateFromHBitmap(hCaptureBitmap)
End Function

'Creates a stdImage object from a HICON
'@constructor
'@param hIcon - Handle to icon
'@returns - Picture object representing data
'@example ```vb
'Dim hICON as LongPtr: hICON = stdWindow.CreateFromHwnd(Application.hWnd).HICON
'stdImage.CreateFromHICON(hICON).ToFile("H:\test.png", stdImgFormatPNG)
'```
#If VBA7 Then
  Public Function CreateFromHICON(ByVal hIcon As LongPtr) As stdImage
#Else
  Public Function CreateFromHICON(ByVal hIcon as Long) as stdImage
#End If
  'Get icon information from HICON
  Dim iconInf As ICONINFO: If GetIconInfo(hIcon, iconInf) = 0 Then Err.Raise 1, "stdImage::CreateFromHICON", "Could not get icon information"
  
  'Get bitmap from icon
  Dim bmp As BITMAP: Call GetObjectAPI(iconInf.hbmColor, LenB(bmp), bmp)

  'Get DC
  Dim hDC As LongPtr: hDC = GetDC(0)

  'Create an in-memory DC for the copy process
  Dim hCaptureDC As LongPtr: hCaptureDC = CreateCompatibleDC(hDC)

  'Create a bitmap to capture the window
  Dim hCaptureBitmap As LongPtr: hCaptureBitmap = CreateCompatibleBitmap(hDC, bmp.bmWidth, bmp.bmHeight)

  'Load bitmap into in-memory DC
  Dim hOldBitmap As LongPtr: hOldBitmap = SelectObject(hCaptureDC, hCaptureBitmap)

  'Copy the on-screen image into the memory DC.
  Const DI_NORMAL As Long = &H3
  Dim r As Long: r = DrawIconEx(hCaptureDC, 0, 0, hIcon, bmp.bmWidth, bmp.bmHeight, 0, 0, DI_NORMAL)

  'Move old bitmap back into capture DC
  Call SelectObject(hCaptureDC, hOldBitmap)

  'Create stdImage from bitmap
  Set CreateFromHICON = CreateFromHBitmap(hCaptureBitmap)

  'Clean up
  Call DeleteObject(iconInf.hbmColor)
  Call DeleteObject(iconInf.hbmMask)
  Call DeleteDC(hCaptureDC)
  Call ReleaseDC(0, hDC)
End Function

'Initialised stdImage
'@protected
'@param hBITMAP - Handle to bitmap
Public Sub protInit(ByVal hBITMAP As LongPtr)
  Call GetObjectAPI(hBITMAP, LenB(This.HBITMAPTYPE), This.HBITMAPTYPE)
  This.HBITMAP = hBITMAP
  This.Cache.HICON = 0
End Sub

'Obtain a HBitmap
'@returns stdole.StdPicture - Obtain the image as a stdole.StdPicture
'@remark You are in charge of freeing this pointer.
Public Property Get HBitmap() As LongPtr
  HBitmap = This.HBITMAP
End Property

'Obtain (and cache) an HICON
Public Property Get HICON() As LongPtr
  If This.Cache.HICON = 0 Then
    Dim iconInf As ICONINFO
    iconInf.fIcon = 1
    iconInf.hbmMask = This.HBITMAP
    iconInf.hbmColor = This.HBITMAP
    This.Cache.HICON = CreateIconIndirect(iconInf)
  End If
  HICON = This.Cache.HICON
End Property

'Obtain the colors of the image
'@returns ARGB() - 2D array of colors. Colors ordered from (1,1) = bottom-left, (1,height) = top-left, across to (width, height) = top-right.
Public Property Get Colors() As Long()
  With This.HBITMAPTYPE
    'Obtain colors
    Dim ret() As Long
    ReDim ret(1 To .bmWidth, 1 To .bmHeight)
    Call CopyMemory(VarPtr(ret(1, 1)), .bmBits, .bmWidth * .bmHeight * 4)
    Colors = ret
  End With
End Property

'Obtain the Image Data
Public Property Get Typee()      As Long    : Typee      = This.HBITMAPTYPE.bmType       : End Property 'Type cannot be used As a name, thats why its named Typee
Public Property Get Width()      As Long    : Width      = This.HBITMAPTYPE.bmWidth      : End Property
Public Property Get Height()     As Long    : Height     = This.HBITMAPTYPE.bmHeight     : End Property
Public Property Get WidthBytes() As Long    : WidthBytes = This.HBITMAPTYPE.bmWidthBytes : End Property
Public Property Get Planes()     As Integer : Planes     = This.HBITMAPTYPE.bmPlanes     : End Property
Public Property Get BitsPixel()  As Integer : BitsPixel  = This.HBITMAPTYPE.bmBitsPixel  : End Property
Public Property Get Bits()       As LongPtr : Bits       = This.HBITMAPTYPE.bmBits       : End Property


'Obtain binary data of the image in any format required
'@param eFormat - The format the data should be in.
'@returns - Binary data of image, ready to load into a file
Public Function ToBinary(Optional ByVal eFormat As stdImgFormat = stdImgFormatDefault) As Byte()
  'HACK: This is a hack to get the image data as a byte array. It is not the most efficient way to do this.
  'TODO: Implement a more efficient way to do this.
  Dim sPath As String: sPath = ToTempFile()
  Dim ff As Long: ff = FreeFile()
  Open sPath For Binary As #ff
    ToBinary = InputB(LOF(ff), 1)
  Close #ff
End Function

'Save image as file
'@param sSaveAsFile - Path to new file
'@param eFormat - The required format of the file. Default is the current assosciated format.
Public Sub ToFile(ByVal FileName As String, Optional format As stdImgFormat = stdImgFormatDefault, Optional ByVal Quality As Byte = 80, Optional ByVal TIFF_ColorDepth As Long = 24, Optional ByVal TIFF_Compression As Long = 6)
  If format = stdImgFormatDefault Then
    Dim pathParts: pathParts = Split(FileName, ".")
    Dim ext As String: ext = pathParts(UBound(pathParts))
    Select Case LCase(ext)
      Case "png": format = stdImgFormat.stdImgFormatPNG
      Case "bmp": format = stdImgFormat.stdImgFormatBMP
      Case "gif": format = stdImgFormat.stdImgFormatGIF
      Case "jpeg", "jpg": format = stdImgFormat.stdImgFormatJPEG
      Case "tiff": format = stdImgFormat.stdImgFormatTIFF
      Case "webp": Err.Raise 1, "stdImage::CreateFromFile", "WEBP format is not currently supported"
      Case Else
        Err.Raise 1, "stdImage::CreateFromFile", "Could not determine format from file extension"
    End Select
  End If
  
  Dim tSI As GdiplusStartupInput
  tSI.GdiplusVersion = 1
  
  Dim lRes As Long, lGDIP As LongPtr
  lRes = GdiplusStartup(lGDIP, tSI)
  If lRes = 0 Then
    Dim gdipBitmap As LongPtr
    lRes = GdipCreateBitmapFromHBITMAP(This.HBITMAP, 0, gdipBitmap)
    If lRes = 0 Then
      Dim EncoderGUID As GUID, tParams As EncoderParameters
      Select Case format
        Case stdImgFormatJPEG
          Call IIDFromString(StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), EncoderGUID)
          tParams.count = 1
          ReDim tParams.Parameter(1 To 1)
          With tParams.Parameter(1)
            Call IIDFromString(StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID)
            .NumberOfValues = 1
            .type = 4
            .Value = VarPtr(Quality)
          End With
          

        Case stdImgFormatPNG
          Call IIDFromString(StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), EncoderGUID)

        Case stdImgFormatGIF
          Call IIDFromString(StrPtr("{557CF402-1A04-11D3-9A73-0000F81EF32E}"), EncoderGUID)

        Case stdImgFormatTIFF
          Call IIDFromString(StrPtr("{557CF405-1A04-11D3-9A73-0000F81EF32E}"), EncoderGUID)
          tParams.count = 2
          ReDim tParams.Parameter(1 To 2)
          With tParams.Parameter(1)
            .NumberOfValues = 1
            .type = 4
            Call IIDFromString(StrPtr("{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"), .GUID)
            .Value = VarPtr(TIFF_Compression)
          End With
          With tParams.Parameter(2)
            .NumberOfValues = 1
            .type = 4
            Call IIDFromString(StrPtr("{66087055-AD66-4C7C-9A18-38A2310B8337}"), .GUID)
            .Value = VarPtr(TIFF_ColorDepth)
          End With

        Case stdImgFormatBMP
          Call SavePicture(ToStdPicture(), FileName)
          Exit Sub
      End Select

      'Save the image. It seems that encoderParams should be NULL if there are no parameters.
      If tParams.count = 0 Then
        lRes = GdipSaveImageToFile(gdipBitmap, StrPtr(FileName), EncoderGUID, ByVal 0)
      Else
        lRes = GdipSaveImageToFile(gdipBitmap, StrPtr(FileName), EncoderGUID, tParams)
      End If
      If lRes <> 0 Then Err.Raise 1, "stdImage::ToFile", "Could not save image to file (Error Result: " & lRes & ")"

      GdipDisposeImage gdipBitmap
    Else
      Err.Raise 1, "stdImage::ToFile", "Could not convert HBITMAP to GDI+ image (Error Result: " & lRes & ")"
    End If
    GdiplusShutdown lGDIP
  End If
End Sub

'Copy the image to the clipboard
'@example `stdImage.CreateFromFile("H:\test1.bmp").ToClipboard()
Public Sub ToClipboard()
  Dim oHtml As Object
  Set oHtml = CreateObject("htmlfile")
  If VarType(Clipboard) = vbObject Then
    Call oHtml.parentWindow.clipboardData.SetData(ToStdPicture())
  Else
    'Hack: This is a hack; better approach would be to get the image data as a byte array and use clipboard directly.
    'TODO: A less hacky method can be found here:  http://access.mvps.org/access/api/api0042.htm
    With ActiveSheet.OLEObjects.Add(ClassType:="Forms.Image.1", Link:=False, DisplayAsIcon:=False, Left:=330, Top:=215)
      .Activate
      Set .Picture = ToStdPicture()
      .AutoSize = True
      .CopyPicture
      .Delete
    End With
  End If
End Sub


'Save a copy in a temporary file and return the path of this file
'@param eFormat - The format the data should be in.
'@returns - The path to the temporary file.
Public Function ToTempFile(Optional ByVal eFormat As stdImgFormat = stdImgFormatDefault) As String
  Dim sExt As String
  Select Case eFormat
    Case stdImgFormatBMP:  sExt = "bmp"
    Case stdImgFormatPNG:  sExt = "png"
    Case stdImgFormatGIF:  sExt = "gif"
    Case stdImgFormatJPEG: sExt = "jpeg"
    Case stdImgFormatTIFF: sExt = "tiff"
  End Select
  Dim sPath As String: sPath = Environ("Temp") & "\" & getGUID() & "." & sExt
  Call ToFile(sPath, eFormat)
  ToTempFile = sPath
End Function

'Adds the picture to the sheet
'@param ws - Worksheet to add image to
'@param Top - Top coord of the newly placed image
'@param Left - Left coord of the newly placed image
'@param compress - Whether to compress the image on import.
'@returns Object<Picture> - Picture object
Public Function ToSheet(ByVal ws As Worksheet, Optional ByVal Top As Double = 0, Optional ByVal Left As Double = 0, Optional ByVal compress As MsoPictureCompress) As Object
  Set ToSheet = ws.Shapes.AddPicture2(ToTempFile(), False, 0, Top, Left, -1, -1, compress)
End Function

'Update a UI control's background picture
'@param oControl as Object<Control> - The control to update the picture of
'@example ```vb
'stdImage.CreateFromFile("H:\test1.bmp").ToUIControl(UserForm1.Image1)
'```
Public Sub ToUIControl(ByVal oControl As Object)
  Set oControl.Picture = ToStdPicture()
End Sub

'Obtain a data URL of the image
'@param eFormat - The format of the image
'@returns - Data URL of the image
Public Function ToDataUrl(Optional ByVal eFormat As stdImgFormat = stdImgFormatDefault) As String
  'TODO:
  Err.Raise 1, , "Not implemented"
End Function

'Output the image to a shape's fill
'@param shp as Object<Excel.Shape> - The shape to fill
'@example ```vb
'stdImage.CreateFromFile("H:\test1.bmp").ToShapeFill(ActiveSheet.Shapes(1))
'```
Public Sub ToShapeFill(ByVal shp As Object)
  Select Case Application.Name
    Case "Microsoft Excel"
      Dim xlShp As Excel.Shape: Set xlShp = shp
      Call xlShp.Fill.UserPicture(ToTempFile())
  End Select
End Sub

'Obtain a stdole.StdPicture
'@returns - Obtain the image as a stdole.StdPicture
Public Function ToStdPicture() As stdole.StdPicture
  Const PICTYPE_BITMAP = 1

  Dim IID_IDispatch As GUID
  If IIDFromString(StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), IID_IDispatch) <> 0 Then Err.Raise 1, , "Could not convert IID_IDispatch to GUID"

  Dim uPicinfo As PICTDESC
  With uPicinfo
    .size = Len(uPicinfo)
    .Type = PICTYPE_BITMAP
    .hPic = This.HBITMAP
    .hPal = 0
  End With

  Dim lResult As Long: lResult = OleCreatePictureIndirect(uPicinfo, IID_IDispatch, 0, ToStdPicture)
  If lResult <> 0 Then Call Err.Raise(1, "Value [GET]", "Could not create StdPicture (Error Result: " & lResult & ")")
End Function

Private Sub Class_Terminate()
  'Free HICON
  If This.Cache.HICON <> 0 Then Call DestroyIcon(This.Cache.HICON)
  
  'Free HBITMAP
  If This.HBITMAP <> 0 Then Call DeleteObject(This.HBITMAP)
End Sub

'@devRemark source - https://www.mrexcel.com/board/threads/working-with-images-in-vba-displaying-png-files.1224206/#post-5991836
'@devRemark - Not sure if this is needed, could be useful for ToBinary() method.
Private Function getHBitmapBytes(ByVal hBitmap As LongPtr) As Byte()
    Const DIB_RGB_COLORS = 0&
    
    Dim tBmpInf As BITMAPINFO, tBmpFileHearder As BITMAPFILEHEADER, tBitmap As BITMAP
    Dim bDIBBits() As Byte
    Dim hDC As LongPtr, hDib As LongPtr
    
    If GetObjectAPI(hBitmap, LenB(tBitmap), tBitmap) = NULL_PTR Then
        MsgBox "Failed to retrieve info for the bitmap.": GoTo ReleaseHandles
    End If
    With tBmpInf.bmiheader
        .biSize = LenB(tBmpInf.bmiheader)
        .biWidth = tBitmap.bmWidth
        .biHeight = tBitmap.bmHeight
        .biPlanes = 1&
        .biBitCount = 32&
        .biSizeImage = .biWidth * 4& * .biHeight
        hDib = CreateDIBSection(NULL_PTR, tBmpInf, 0&, NULL_PTR, NULL_PTR, 0&)
        If hDib = NULL_PTR Then
            MsgBox "Failed to create a DIB.": GoTo ReleaseHandles
        End If
        'OleLoadPicture expects the graphic byte array to include 54 bytes [file header + Inf header].
        ReDim bDIBBits(0& To .biSizeImage + 53&)
        'Fill bmp file header
        Call CopyMemory(VarPtr(bDIBBits(0&)), VarPtr(&H4D42&), 2&)
        Call CopyMemory(VarPtr(bDIBBits(2&)), VarPtr(54& + .biSizeImage), 4&)
        Call CopyMemory(VarPtr(bDIBBits(10&)), VarPtr(54&), 4&)
        'Fill bmp info header
        Call CopyMemory(VarPtr(bDIBBits(14&)), VarPtr(tBmpInf), 40&)
        hDC = GetDC(NULL_PTR)
        If GetDIBits(hDC, hBitmap, 0&, .biHeight, bDIBBits(54&), tBmpInf, DIB_RGB_COLORS) = NULL_PTR Then
            MsgBox "Failed to retrieve the bits of the bitmap.": GoTo ReleaseHandles
        End If
    End With
    getHBitmapBytes = bDIBBits
ReleaseHandles:
    Call ReleaseDC(NULL_PTR, hDC)
    Call DeleteObject(hDib)
End Function

'Returns a byte array stored at a pointer using RtlMoveMemory
'@param pointer - The pointer to copy bytes at
'@param size - The size of the buffer to copy
'@returns - The bytes stired at the pointer
'@devRemark - Not sure if this is needed, could be useful for ToBinary() method.
Private Function BytesFromPointer(ByVal pointer As LongPtr, ByVal size As Long) As Byte()
  Dim buff() As Byte
  If size > 0 Then
    ReDim buff(1 To size)
    Call CopyMemory(VarPtr(buff(1)), pointer, size)
  Else
    Call Err.Raise(1, "BytesFromPointer [PRIVATE]", "No size supplied.")
  End If
  BytesFromPointer = buff
End Function

'Tries to open the clipbaord with a timeout
'@param hWnd - hWnd to open clipboard in
'@param dwTimeoutMilliseconds - number of milliseconds to wait
'@returns - true if clipboard is open, false if clipboard couldn't be opened
Private Function OpenClipboardTimeout(ByVal hWnd As LongPtr, Optional ByVal dwTimeoutMilliseconds As Long = 5000) As Boolean
  Dim iStart As Long: iStart = GetTickCount()
  Dim bOpen As Boolean, bExpired As Boolean
  While Not (bOpen Or bExpired)
    'Open clipboard and only continue if open
    bOpen = CBool(OpenClipboard(hWnd))
    OpenClipboardTimeout = bOpen
    If bOpen Then Exit Function
    
    'Check for expirey
    bExpired = Abs(GetTickCount() - iStart) >= dwTimeoutMilliseconds
    
    'Do Excel events
    DoEvents
  Wend
End Function

'Generates a new GUID
'@returns - A new GUID
Private Function getGUID() As String
  Call Randomize 'Ensure random GUID generated
  getGUID = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx"
  getGUID = Replace(getGUID, "y", Hex(Rnd() And &H3 Or &H8))
  Dim i As Long: For i = 1 To 30
    getGUID = Replace(getGUID, "x", Hex$(Int(Rnd() * 16)), 1, 1)
  Next
End Function

'Converts a long color to a BGR color
'@param longColor - The long color to convert
'@returns - The BGR color
Private Function BGR(ByVal longColor As Long) As Long
  Call OleTranslateColor(longColor, 0, longColor)
  BGR = RGB((longColor \ 65536) Mod 256, (longColor \ 256) Mod 256, (longColor Mod 256))
End Function

#If VBA7 Then
Private Function getAppHwnd() As LongPtr
#Else
Private Function getAppHwnd() As Long
#End If
Static hWnd As LongPtr
If hWnd = 0 Then
    Dim oApp As Object: Set oApp = Application
    Select Case oApp.Name
    Case "Microsoft Excel"
        hWnd = oApp.hwnd
    Case "Microsoft Word"
        hWnd = oApp.ActiveWindow.Hwnd
    Case "Microsoft Access"
        hWnd = oApp.hWndAccessApp.Hwnd
    Case "Microsoft PowerPoint"
        hWnd = oApp.CommandBars("Status Bar").Hwnd
    Case "Outlook"
        hWnd = oApp.ActiveWindow.CommandBars("Status Bar").Hwnd
    Case "Microsoft Publisher"
        hWnd = oApp.CommandBars("Status Bar").Hwnd
    Case Else
        Err.Raise 1, "stdAcc::CreateFromApplication()", "No implementation for getting application window of " & oApp.Name
    End Select
End If
getAppHwnd = hWnd
End Function


'Searches for a template image within a larger image (i.e., this image)
'@param template - The image to search for
'@param tolerance - A value of 0 means that the pixels must be an exact match; greater values signify slight (1-5), moderate (6 - 20) and signifanct (20+) variations (default 5)
'@param TopToBottom - If True, search from top-to-bottom; if False, search from bottom-to-top (default True)
'@param LeftToRight - If True, search from left-to-right; if False, search from right-to-left (default True)
'@returns Variant<Array<Long>|Empty> - Array [x, y] of the top-left match (standard image coordinates), or Empty if not found
'@example
'  Dim haystack As stdImage, needle As stdImage, result As Variant
'  Set haystack = stdImage.CreateFromFile("C:\YourImageDir\big.png")
'  Set needle = stdImage.CreateFromFile("C:\YourImageDir\small.png")
'  result = haystack.FindImage(needle, 10)
'  If Not IsEmpty(result) Then Debug.Print "Found at:", result(0), result(1)
Public Function FindImage(ByVal template As stdImage, Optional ByVal tolerance As Long = 5, Optional ByVal TopToBottom As Boolean = True, Optional ByVal LeftToRight As Boolean = True) As Variant
  Dim haystackColors() As Long, templateColors() As Long
  Dim hayW As Long, hayH As Long, tplW As Long, tplH As Long
  Dim startX As Long, endX As Long, stepX As Long
  Dim startY As Long, endY As Long, stepY As Long
  
  haystackColors = Me.Colors
  templateColors = template.Colors
  hayW = UBound(haystackColors, 1): hayH = UBound(haystackColors, 2)
  tplW = UBound(templateColors, 1): tplH = UBound(templateColors, 2)
    
  If TopToBottom Then
    startY = 0: endY = hayH - tplH: stepY = 1
  Else
    startY = hayH - tplH: endY = 0: stepY = -1
  End If
  
  If LeftToRight Then
    startX = 0: endX = hayW - tplW: stepX = 1
  Else
    startX = hayW - tplW: endX = 0: stepX = -1
  End If
  
  Dim x As Long, y As Long, i As Long, j As Long, match As Boolean
  For y = startY To endY Step stepY
    For x = startX To endX Step stepX
      match = True
      For j = 1 To tplH
        For i = 1 To tplW
          If Not ColorsMatch(haystackColors(x + i, y + j), templateColors(i, j), tolerance) Then
            match = False
            Exit For
          End If
        Next
        If Not match Then Exit For
      Next
      If match Then
        FindImage = Array(x, hayH - y - tplH + 1)
        Exit Function
      End If
    Next
  Next
  FindImage = Empty
End Function

'Compares two ARGB color values within a given tolerance per RGB channel
'@param c1 - First color value
'@param c2 - Second color value
'@param tolerance - Maximum allowed difference per channel (0-255).
'@returns - True if the colors match within the given tolerance, otherwise False
Private Function ColorsMatch(ByVal c1 As Long, ByVal c2 As Long, ByVal tolerance As Long) As Boolean
  Dim r1 As Long, g1 As Long, b1 As Long, r2 As Long, g2 As Long, b2 As Long
  r1 = (c1 And &HFF): g1 = (c1 \ &H100) And &HFF: b1 = (c1 \ &H10000) And &HFF
  r2 = (c2 And &HFF): g2 = (c2 \ &H100) And &HFF: b2 = (c2 \ &H10000) And &HFF
  ColorsMatch = Abs(r1 - r2) <= tolerance And Abs(g1 - g2) <= tolerance And Abs(b1 - b2) <= tolerance
End Function