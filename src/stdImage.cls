VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


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

  Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As Long)
  Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long

  Private Declare PtrSafe Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As LongPtr, ByVal offsetinVft As LongPtr, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As VbVarType, ByRef paValues As LongPtr, ByRef retVAR As Variant) As Long

  Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" (token As LongPtr, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As LongPtr = 0) As Long
  Private Declare PtrSafe Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal FILENAME As LongPtr, BITMAP As LongPtr) As Long
  Private Declare PtrSafe Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal BITMAP As LongPtr, hbmReturn As LongPtr, ByVal background As LongPtr) As Long
  Private Declare PtrSafe Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hBitmap as LongPtr, ByVal hPalette as LongPtr, gdipBitmap as LongPtr) As Long
  Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal Image As LongPtr) As Long
  Private Declare PtrSafe Sub GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr)
  Private Declare PtrSafe Function GdipSaveImageToFile Lib "gdiplus" (ByVal Image As LongPtr, ByVal Filename As LongPtr, ByRef clsidEncoder As GUID, ByRef encoderParams As Any) As Long
  Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long
  Private Declare PtrSafe Function OleTranslateColor Lib "oleAut32.dll" (ByVal clr As OLE_COLOR, ByVal palet As LongPtr, Col As Long) As Long
  Private Declare PtrSafe Function IIDFromString Lib "ole32" (ByVal lpsz As LongPtr, ByRef lpiid As GUID) As Long
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
#End If
Private Const CF_BITMAP = 2

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
  HICON as LongPtr
End Type
Private Type TThis
  HBITMAP as LongPtr
  Cache as TThisCache
End Type
Private This as TThis


'****************
'* Constructors *
'****************

'Creates a stdImage object from a byte array
'@param b - Bytes of Bitmap/Jpeg/Png/... image
'@param format - The format of the image
'@returns {stdImage} Picture object representing data
'@example ```
'  Dim b() as Byte: Redim b(1 to 100)
'  Call stdImage.CreateFromBinary(b).saveAs("H:\test1.png", stdImgFormatPNG)
'```
Public Function CreateFromBinary(ByRef b() As Byte, ByVal format as stdImgFormat) As stdImage
  'TODO:
  Err.Raise 1, , "Not implemented"
End Function


Public Function CreateFromDataURL(ByVal sURL as string) as stdImage
  'TODO:
  Err.Raise 1, , "Not implemented"
End Function

'Creates a stdImage object from a file
'@param sPath - Path of file
'@param format - The format of the image
'@returns {stdImage} Picture object representing data
'@example `stdImage.CreateFromFile("H:\test1.bmp").saveAs("H:\test1.png", stdImgFormatPNG)`
Public Function CreateFromFile(ByVal sPath As String, Optional ByVal format as stdImgFormat = stdImgFormatDefault, optional byval transparentColor as Long = 0) As stdImage
  If format = stdImgFormatDefault Then
    Dim pathParts: pathParts = Split(sPath, ".")
    Dim ext As String: ext = pathParts(UBound(pathParts))
    Select Case LCase(ext)
      Case "png": format = stdImgFormat.stdImgFormatPNG
      Case "bmp": format = stdImgFormat.stdImgFormatBMP
      Case "gif": format = stdImgFormat.stdImgFormatGIF
      Case "jpeg", "jpg": format = stdImgFormat.stdImgFormatJPEG
      Case "tiff": format = stdImgFormat.stdImgFormatTIFF
      Case "webp": Err.Raise 1, "stdImage::CreateFromFile", "WEBP format is not currently supported"
      case else
        Err.Raise 1, "stdImage::CreateFromFile", "Could not determine format from file extension"
    End Select
  End If
  Select Case format
    'TODO: Case webp
    Case Else 'BMP, GIF, JPEG, PNG, TIFF, Exif, WMF, and EMF - Load using GDI+
      Dim uGdiInput As GdiplusStartupInput
      uGdiInput.GdiplusVersion = 1
      Dim lResult as Long: lResult = GdiplusStartup(hGdiPlus, uGdiInput)
      if lResult = 0 then
        Dim hGDIImage As LongPtr
        lResult = GdipCreateBitmapFromFile(StrPtr(sPath), hGDIImage)
        if lResult = 0 then
          Dim hBitmap as LongPtr
          lResult = GdipCreateHBITMAPFromBitmap(hGDIImage, hBitmap, BGR(transparentColor))
          if lResult = 0 then
            Set CreateFromFile = CreateFromHBitmap(hBitmap)
          else
            Err.Raise 1, "stdImage::CreateFromFile", "Could not convert GDI+ image to HBITMAP (Error Result: " & lResult & ")"
          end if

          Call GdipDisposeImage(hGDIImage)
        else
          Err.Raise 1, "stdImage::CreateFromFile", "Could not load image (Error Result: " & lResult & ")"
        end if

        Call GdiplusShutdown(hGdiPlus)
      else
        Err.Raise 1, "stdImage::CreateFromFile", "Could not start GDI+ (Error Result: " & lResult & ")"
      end if
  End Select
End Function

'Creates a stdImage object from a url
'@param sUrl - Url of file
'@returns - Picture object representing data
'@example `stdImage.CreateFromURL("https://raw.githubusercontent.com/sancarn/awesome-vba/main/resources/VBALogo.png").saveAs("H:\test.png", stdImgFormatPNG)`
'@remark - TODO: consider removing
Public Function CreateFromUrl(ByVal sURL As String, Optional ByVal format as stdImgFormat = stdImgFormatDefault) As stdImage
  Set CreateFromUrl = New stdImage
  Dim HTTP As Object: Set HTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
  HTTP.Open "GET", TargetURL, False
  'TODO: AutoProxy?
  HTTP.send
  If HTTP.status = 200 Then
    Dim body() As Byte: body = HTTP.responseBody
    if format = stdImgFormatDefault then
      'TODO: Detect format
    end if
    Set CreateFromUrl = CreateFromBinary(body, format)
  Else
    Err.Raise 1, "", "Could not get file (HTTP Status: " & HTTP.status & ")"
  End If
End Function

'Creates a stdImage object from a file
'@param {LongPtr} hBitmap (handle to Bitmap)
'@returns {stdImage} Picture object representing data
'@example ```
'  stdImage.CreateFromHBitmap(Image1.Picture.Handle).saveAs("H:\test.png", stdImgFormatPNG)
'```
Public Function CreateFromHBitmap(ByVal hBitmap As LongPtr) As stdImage
  Dim bmp As BITMAP
  if GetObjectAPI(hBitmap, LenB(bmp), bmp) <> 0 then
    set CreateFromHBitmap = New stdImage
    Call CreateFromHBitmap.protInit(hBitmap)
  else
    Err.Raise 1, "stdImage::CreateFromHBitmap" , "The provided hBITMAP is not valid"
  end if
End Function

'Creates a stdImage from the image in the clipboard
'@returns {stdImage} Picture object representing data
'@example stdImage.CreateFromClipboard().saveAs("H:\test.png", stdImgFormatPNG)
Public Function CreateFromClipboard() As stdImage
  'Open clipboard
  If Not CBool(OpenClipboardTimeout(Application.hWnd)) Then Call Err.Raise(1, "Value [GET]", "Unable to open clipboard (DllError: " & Err.LastDllError & ")")
  
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
'@param picture - Object to create picture from
'@returns - Picture object representing data
'@example TODO:
Public Function CreateFromIPictureDisp(ByVal picture As stdole.IPictureDisp) As stdImage
  Set CreateFromIPictureDisp = CreateFromHBitmap(picture.Handle)
End Function

'Creates a stdImage object from a stdole.stdPicture object
'@param picture - Object to create picture from
'@returns - Picture object representing data
'@example `stdImage.CreateFromStdPicture(Userform1.Picture).saveAs("H:\test.png", stdImgFormatPNG)`
Public Function CreateFromStdPicture(ByVal picture As stdole.stdPicture) As stdImage
  Set CreateFromStdPicture = CreateFromHBitmap(picture.Handle)
End Function

'Creates a stdImage object from a stdole.IPicture object
'@param picture - Object to create picture from
'@returns - Picture object representing data
'@example TODO:
Public Function CreateFromIPicture(ByVal picture As stdole.IPicture) As stdImage
  Set CreateFromIPicture = CreateFromHBitmap(picture.Handle)
End Function

'Creates a stdImage object from a Excel.IPicture object
'@param {Excel.IPicture} Object to create picture from
'@returns {stdImage} Picture object representing data
'@example TODO:
Public Function CreateFromExcelIPicture(ByVal picture As Excel.IPicture) As stdImage
  Call picture.CopyPicture(, xlBitmap)
  Set CreateFromExcelIPicture = CreateFromClipboard()
End Function

'Creates a stdImage object from a Shape object
'@param {Object<(Excel|Word|PowerPoint).Shape>} Object to create picture from
'@returns {stdImage} Picture object representing data
'@example `stdImage.CreateFromShape(Sheet1.Shapes("Picture 2")).saveAs("H:\test.png")`
Public Function CreateFromShape(ByVal shp As Object) As stdImage
  Select Case Application.Name
    Case "Microsoft Excel"
      Dim xlShp As Excel.Shape: Set xlShp = shp
      Call xlShp.CopyPicture(, XlCopyPictureFormat.xlBitmap)
      Set CreateFromShape = stdImage.CreateFromClipboard()
  End Select
End Function

'Creates a stdImage object from a Chart object
'@param {Object<(Excel|Word|PowerPoint).Chart>} Object to create picture from
'@returns {stdImage} Picture object representing data
'@example `stdImage.CreateFromChart(Sheet1.ChartObjects(1).Chart).saveAs("H:\test.png")`
Public Function CreateFromChart(ByVal chart As Object) As stdImage
  Select Case Application.Name
    Case "Microsoft Excel"
      Dim xlChart As Excel.chart: Set xlChart = chart
      Call xlChart.CopyPicture(, XlCopyPictureFormat.xlBitmap)
      Set CreateFromChart = stdImage.CreateFromClipboard()
  End Select
End Function

'Initialised stdImage from WIA object
'@param {Object<WIA.ImageFile>}
Public Sub protInit(ByVal hBITMAP As LongPtr)
  This.HBITMAP = hBITMAP
  This.Cache.HICON = 0
End Sub

'Obtain a HBitmap
'@returns {stdole.StdPicture} Obtain the image as a stdole.StdPicture
'@remark You are in charge of freeing this pointer.
Public Property Get HBitmap() as LongPtr
  HBitmap = This.HBITMAP
End Property

'Obtain (and cache) an HICON
Public Property Get HICON() as LongPtr
  if This.Cache.HICON = 0 then
    Dim iconInf as ICONINFO
    iconInf.fIcon = 1
    iconInf.hbmMask = This.HBITMAP
    iconInf.hbmColor = This.HBITMAP
    This.Cache.HICON = CreateIconIndirect(iconInf)
  end if
  HICON = This.Cache.HICON
End Property





'Obtain binary data of the image in any format required
'@param eFormat - The format the data should be in.
'@returns - Binary data of image, ready to load into a file
Public Function ToBinary(Optional ByVal eFormat As stdImgFormat = stdImgFormatDefault) As Byte()
  'HACK: This is a hack to get the image data as a byte array. It is not the most efficient way to do this.
  'TODO: Implement a more efficient way to do this.
  Dim sPath as string: sPath = ToTempFile()
  Dim ff as long: ff = FreeFile()
  Open sPath For Binary As #ff
    ToBinary = InputB(LOF(ff), 1)
  Close #ff
End Function

'Save image as file
'@param sSaveAsFile - Path to new file
'@param eFormat - The required format of the file. Default is the current assosciated format.
Public Sub ToFile(ByVal FileName As String, Optional format As stdImgFormat = stdImgFormatDefault, Optional ByVal Quality As Byte = 80, Optional ByVal TIFF_ColorDepth As Long = 24, Optional ByVal TIFF_Compression As Long = 6)
  if format = stdImgFormatDefault then
    Dim pathParts: pathParts = Split(FileName, ".")
    Dim ext As String: ext = pathParts(UBound(pathParts))
    Select Case LCase(ext)
      Case "png": format = stdImgFormat.stdImgFormatPNG
      Case "bmp": format = stdImgFormat.stdImgFormatBMP
      Case "gif": format = stdImgFormat.stdImgFormatGIF
      Case "jpeg", "jpg": format = stdImgFormat.stdImgFormatJPEG
      Case "tiff": format = stdImgFormat.stdImgFormatTIFF
      Case "webp": Err.Raise 1, "stdImage::CreateFromFile", "WEBP format is not currently supported"
      case else
        Err.Raise 1, "stdImage::CreateFromFile", "Could not determine format from file extension"
    End Select
  end if
  
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
    End If
    GdiplusShutdown lGDIP
  End If
End Sub

'Copy the image to the clipboard
'@example `stdImage.CreateFromFile("H:\test1.bmp").ToClipboard()
Public Sub ToClipboard()
  if VarType(Clipboard) = vbObject then
    Call Clipboard.SetData(ToStdPicture())
  else
    'Hack: This is a hack; better approach would be to get the image data as a byte array and use clipboard directly.
    'TODO: A less hacky method can be found here:  http://access.mvps.org/access/api/api0042.htm
    With ActiveSheet.OLEObjects.Add(ClassType:="Forms.Image.1", Link:=False, DisplayAsIcon:=False, Left:=330, Top:=215)
      .Activate
      set .Picture = ToStdPicture()
      .AutoSize = true
      .CopyPicture
      .Delete
    End with
  end if
End Sub

'Save a copy in a temporary file and return the path of this file
'@param eFormat - The format the data should be in.
'@returns - The path to the temporary file.
Public Function ToTempFile(Optional ByVal eFormat As stdImgFormat = stdImgFormatDefault) as string
  Dim sExt As String
  Select Case eFormat
    Case stdImgFormatBMP:  sExt = "bmp"
    Case stdImgFormatPNG:  sExt = "png"
    Case stdImgFormatGIF:  sExt = "gif"
    Case stdImgFormatJPEG: sExt = "jpeg"
    Case stdImgFormatTIFF: sExt = "tiff"
  End Select
  Dim sPath as string: sPath = Environ("Temp") & "\" & getGUID() & "." & sExt
  Call ToFile(sPath, eFormat)
  ToTempFile = sPath
End Function

'Adds the picture to the sheet
'@param ws - Worksheet to add image to
'@param Top - Top coord of the newly placed image
'@param Left - Left coord of the newly placed image
'@param compress - Whether to compress the image on import.
'@returns Object<Picture> - Picture object
Public Function ToSheet(ByVal ws as Worksheet, Optional ByVal Top as Double = 0 , Optional ByVal Left as Double = 0, optional ByVal compress as MsoPictureCompress) as Object
  set ToSheet = ws.Shapes.AddPicture2(toTempFile(), false, 0, Top, Left, -1, -1, compress)
End Function

'Update a UI control's background picture
'@param oControl as Object<Control> - The control to update the picture of
'@example `stdImage.CreateFromFile("H:\test1.bmp").ToUIControl(UserForm1.Image1)`
Public Sub ToUIControl(Byval oControl as object)
  set oControl.Picture = ToStdPicture()
End Sub

'Obtain a data URL of the image
'@param eFormat - The format of the image
'@returns - Data URL of the image
Public Function ToDataUrl(Optional ByVal eFormat As stdImgFormat = stdImgFormatDefault) as string
  'TODO:
  Err.raise 1, , "Not implemented"
End Function

'Output the image to a shape's fill
'@param shp as Object<Excel.Shape> - The shape to fill
'@example `stdImage.CreateFromFile("H:\test1.bmp").ToShapeFill(ActiveSheet.Shapes(1))`
Public Sub ToShapeFill(ByVal shp As Object)
  Select Case Application.Name
    Case "Microsoft Excel"
      Dim xlShp As Excel.Shape: Set xlShp = shp
      Call xlShp.Fill.UserPicture(ToTempFile())
  End Select
End Sub

'Obtain a stdole.StdPicture
'@returns - Obtain the image as a stdole.StdPicture
Public Function ToStdPicture() as stdole.StdPicture
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

Private Sub Class_Uninitialize()
  'Free HICON
  If This.Cache.HICON <> 0 Then Call DestroyIcon(This.Cache.HICON)
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
Private Function OpenClipboardTimeout(ByVal hWnd As Long, Optional ByVal dwTimeoutMilliseconds As Long = 5000) As Boolean
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