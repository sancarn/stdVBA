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
'@example Example of how to get screenshots of chromium windows. Unfortunately `stdImage.CreateFromWindow()` does not work on these types of window.
'So instead you have to take the screenshot from the desktop, and then crop the image to size using the window dimensions.
'```vb
'Sub tChrome()
'  Dim win As stdWindow, i As Long
'  For Each win In stdWindow.CreateFromDesktop().Children
'    If win.Exists Then
'      If win.Visible Then
'        If win.Class Like "Chrome_WidgetWin*" Then
'          i = i + 1
'          Call getChromiumScreenshot(win).ToFile("C:\Temp\chrome\" & i & ".png")
'        End If
'      End If
'    End If
'  Next
'  stdWindow.CreateFromApplication().Activate
'End Sub
'Private Function getChromiumScreenshot(ByVal win As stdWindow) As stdImage
'  win.State = Maximised
'  Call win.Activate
'  Application.Wait Now() + TimeValue("00:00:01")
'  Dim region: region = win.RectWindow()
'  Set getChromiumScreenshot = stdImage.CreateFromScreen().Crop(region(0), region(1), region(2), region(3))
'End Function
'```
'@example Finding the "Record a macro" button in the status window of Excel
'```vb
'Dim excel As stdImage: Set excel = stdImage.CreateFromWindow(Application.hwnd)
'Dim button As stdImage: Set button = stdImage.CreateFromDataURL("data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAA4AAAANCAIAAAAWvsgoAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAACUSURBVChTlZBRDcAgDEQnZgYwgAAUYAADyEAGLrCBA0Twv0dKWCDLxu7jcr2WK3Cc22ijdQN99BHGGGst3OvfqaIEMcZSSghBa00Jo3Hw11Rcljrncs6UMBoHf0qlQQY9GUopyQEc/DWVjbSBUooWLCX+ncp7RcDkIQRo8f+nDnzcdcB7//YDLWoGLhls5DCMxqm1XpNsMy38F0JSAAAAAElFTkSuQmCC")
'Dim pos: pos = excel.FindFirst(button)
'Debug.Print pos(0), pos(1)
'```


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
  Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
  Private Declare PtrSafe Function OpenClipboard Lib "user32" (Optional ByVal hwnd As LongPtr = NULL_PTR) As Long
  Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
  Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
  Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
  Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
  Private Declare PtrSafe Function CopyImage Lib "user32" (ByVal hImage As LongPtr, ByVal imageType As Long, ByVal cx As Long, ByVal cy As Long, ByVal flags As Long) As LongPtr
  Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr

  Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As LongPtr
  Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
  Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
  Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As Long
  Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
  
  'Obtain image bytes from clipboard
  Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
  Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As LongPtr
  Private Declare PtrSafe Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As LongPtr, ByVal nCount As Long, lpObject As Any) As Long
  Private Declare PtrSafe Function CreateDIBSection Lib "gdi32" (ByVal hDC As LongPtr, pbmi As BITMAPINFO, ByVal iUsage As Long, ByVal ppvBits As LongPtr, ByVal hSection As LongPtr, ByVal dwOffset As Long) As LongPtr
  Private Declare PtrSafe Function GetDIBits Lib "gdi32" (ByVal aHDC As LongPtr, ByVal hBitmap As LongPtr, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
  Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
  Private Declare PtrSafe Function CreateIconFromResourceEx Lib "user32" (ByVal pbIconBits As LongPtr, ByVal cbIconBits As Long, ByVal fIcon As Long, ByVal dwVersion As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal uFlags As Long) As LongPtr
  Private Declare PtrSafe Function CreateIconIndirect Lib "user32" (ByRef piconinfo As ICONINFO) As LongPtr
  Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
  Private Declare PtrSafe Function DestroyIcon Lib "user32" (ByVal HICON As LongPtr) As Long
  Private Declare PtrSafe Function DrawIconEx Lib "user32" (ByVal hDC As LongPtr, ByVal xLeft As Long, ByVal yTop As Long, ByVal HICON As LongPtr, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
  Private Declare PtrSafe Function GetIconInfo Lib "user32" (ByVal HICON As LongPtr, ByRef piconinfo As ICONINFO) As Long

  Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As Long)
  Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long

  Private Declare PtrSafe Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As LongPtr, ByVal offsetinVft As LongPtr, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As VbVarType, ByRef paValues As LongPtr, ByRef retVAR As Variant) As Long

  Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" (token As LongPtr, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As LongPtr = 0) As Long
  Private Declare PtrSafe Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal FileName As LongPtr, BITMAP As LongPtr) As Long
  Private Declare PtrSafe Function GdipLoadImageFromFile Lib "gdiplus" (ByVal FileName As LongPtr, ByRef image As LongPtr) As Long
  Private Declare PtrSafe Function GdipLoadImageFromStream Lib "gdiplus" (ByVal stream As stdole.IUnknown, ByRef image As LongPtr) As Long
  Private Declare PtrSafe Function GdipGetImageWidth Lib "gdiplus" (ByVal image As LongPtr, ByRef width As Long) As Long
  Private Declare PtrSafe Function GdipGetImageHeight Lib "gdiplus" (ByVal image As LongPtr, ByRef height As Long) As Long
  Private Declare PtrSafe Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal width As Long, ByVal height As Long, ByVal stride As Long, ByVal pixelFormat As Long, ByVal scan0 As LongPtr, ByRef bitmap As LongPtr) As Long
  Private Declare PtrSafe Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal image As LongPtr, ByRef graphics As LongPtr) As Long
  Private Declare PtrSafe Function GdipGraphicsClear Lib "gdiplus" (ByVal graphics As LongPtr, ByVal color As Long) As Long
  Private Declare PtrSafe Function GdipDrawImageRectI Lib "gdiplus" (ByVal graphics As LongPtr, ByVal image As LongPtr, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long) As Long
  Private Declare PtrSafe Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As LongPtr) As Long
  Private Declare PtrSafe Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal BITMAP As LongPtr, hbmReturn As LongPtr, ByVal background As LongPtr) As Long
  Private Declare PtrSafe Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hBitmap As LongPtr, ByVal hPalette As LongPtr, gdipBitmap As LongPtr) As Long
  Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal Image As LongPtr) As Long
  Private Declare PtrSafe Sub GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr)
  Private Declare PtrSafe Function GdipSaveImageToFile Lib "gdiplus" (ByVal Image As LongPtr, ByVal FileName As LongPtr, ByRef clsidEncoder As GUID, ByRef encoderParams As Any) As Long
  Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long
  Private Declare PtrSafe Function OleTranslateColor Lib "oleAut32.dll" (ByVal clr As OLE_COLOR, ByVal palet As LongPtr, col As Long) As Long
  Private Declare PtrSafe Function IIDFromString Lib "ole32" (ByVal lpsz As LongPtr, ByRef lpiid As GUID) As Long

  'Screenshot
  Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
  Private Declare PtrSafe Function GetWindowDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
  Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
  Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
  Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
  Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
  Private Declare PtrSafe Function BitBlt Lib "gdi32" (ByVal hDestDC As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
  Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long

  'Create from binary
  Private Declare PtrSafe Function GdipCreateBitmapFromStream Lib "gdiplus" (ByVal stream As stdole.IUnknown, ByRef BITMAP As LongPtr) As Long
  Private Declare PtrSafe Function CreateStreamOnHGlobal Lib "ole32.dll" (ByVal hGlobal As LongPtr, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any) As Long
  Private Declare PtrSafe Function GetHGlobalFromStream Lib "ole32" (ByVal stream As stdole.IUnknown, ByRef hGlobal As LongPtr) As Long
  Private Declare PtrSafe Function IStream_Size Lib "shlwapi" (ByVal stream As stdole.IUnknown, ByRef size As Currency) As Long
  Private Declare PtrSafe Function GdipSaveImageToStream Lib "gdiplus" (ByVal Image As LongPtr, ByVal stream As stdole.IUnknown, ByRef clsidEncoder As GUID, ByVal encoderParams As LongPtr) As Long
  Private Declare PtrSafe Function CryptBinaryToStringA Lib "crypt32" (ByRef pbBinary As Byte, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As String, ByRef pcchString As Long) As Long

  'Create from DataURL
  Private Declare PtrSafe Function CryptStringToBinaryA Lib "crypt32.dll" (ByVal pszString As String, ByVal cchString As Long, ByVal dwFlags As Long, ByRef pbBinary As Byte, ByRef pcbBinary As Long, ByVal pdwSkip As LongPtr, ByVal pdwFlags As LongPtr) As Long
  
  Private Declare PtrSafe Function PrintWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal hdcBlt As LongPtr, ByVal nFlags As Long) As Boolean
  Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#Else
  Enum LongPtr
    [_]
  End Enum

  Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
  Private Declare Function OpenClipboard Lib "user32" (Optional ByVal hwnd As LongPtr = NULL_PTR) As Long
  Private Declare Function CloseClipboard Lib "user32" () As Long
  Private Declare Function EmptyClipboard Lib "user32" () As Long
  Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
  Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
  Private Declare Function CopyImage Lib "user32" (ByVal hImage As LongPtr, ByVal imageType As Long, ByVal cx As Long, ByVal cy As Long, ByVal flags As Long) As LongPtr
  Private Declare Function GetActiveWindow Lib "user32" () As LongPtr

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
  Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal FileName As LongPtr, ByRef image As LongPtr) As Long
  Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal stream As stdole.IUnknown, ByRef image As LongPtr) As Long
  Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal image As LongPtr, ByRef width As Long) As Long
  Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal image As LongPtr, ByRef height As Long) As Long
  Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal width As Long, ByVal height As Long, ByVal stride As Long, ByVal pixelFormat As Long, ByVal scan0 As LongPtr, ByRef bitmap As LongPtr) As Long
  Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal image As LongPtr, ByRef graphics As LongPtr) As Long
  Private Declare Function GdipGraphicsClear Lib "gdiplus" (ByVal graphics As LongPtr, ByVal color As Long) As Long
  Private Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal graphics As LongPtr, ByVal image As LongPtr, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long) As Long
  Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As LongPtr) As Long
  Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal BITMAP As LongPtr, hbmReturn As LongPtr, ByVal background As LongPtr) As Long
  Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hBitmap as LongPtr, ByVal hPalette as LongPtr, gdipBitmap as LongPtr) As Long
  Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As LongPtr) As Long
  Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr)
  Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal Image As LongPtr, ByVal Filename As LongPtr, ByRef clsidEncoder As GUID, ByRef encoderParams As Any) As Long
  Private Declare Function OleCreatePictureIndirect Lib "oleaut32" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long
  Private Declare Function OleTranslateColor Lib "oleAut32.dll" (ByVal clr As OLE_COLOR, ByVal palet As LongPtr, Col As Long) As Long
  Private Declare Function IIDFromString Lib "ole32" (ByVal lpsz As LongPtr, ByRef lpiid As GUID) As Long

  'Screenshot
  Private Declare Function GetDesktopWindow Lib "user32" () As LongPtr
  Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
  Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
  Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
  Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr
  Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
  Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
  Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As LongPtr) As Long

  'Create from binary
  Private Declare Function GdipCreateBitmapFromStream Lib "gdiplus" (ByVal stream As stdole.IUnknown, ByRef BITMAP As LongPtr) As Long
  Private Declare Function CreateStreamOnHGlobal Lib "ole32.dll" (ByVal hGlobal As LongPtr, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any) As Long
  Private Declare Function GetHGlobalFromStream Lib "ole32" (ByVal stream As stdole.IUnknown, ByRef hGlobal As LongPtr) As Long
  Private Declare Function IStream_Size Lib "shlwapi" (ByVal stream As stdole.IUnknown, ByRef size As Currency) As Long
  Private Declare Function GdipSaveImageToStream Lib "gdiplus" (ByVal Image As LongPtr, ByVal stream As stdole.IUnknown, ByRef clsidEncoder As GUID, ByVal encoderParams As LongPtr) As Long
  Private Declare Function CryptBinaryToStringA Lib "crypt32" (ByRef pbBinary As Byte, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As String, ByRef pcchString As Long) As Long

  'Create from DataURL
  Private Declare Function CryptStringToBinaryA Lib "crypt32.dll" (ByVal pszString As String, ByVal cchString As Long, ByVal dwFlags As Long, ByRef pbBinary As Byte, ByRef pcbBinary As Long, ByVal pdwSkip As LongPtr, ByVal pdwFlags As LongPtr) As Long
  
  Private Declare Function PrintWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal hdcBlt As LongPtr, ByVal nFlags As Long) As Boolean
  Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
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
    Size As Long
    type As Long
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
#If Win64 Then
  Reserved As Long
#End If
  Parameter(0 To 1) As EncoderParameter
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

Private Type TDIBSECTION
    dsBm As BITMAP            ' Contains general bitmap info
    dsBmih As BITMAPINFOHEADER ' Contains DIB-specific header info
    dsDwOffset As Long        ' Offset to the bitmap bits (if in file/section)
    dsHSection As LongPtr     ' Handle to the file mapping object (if memory-mapped)
    dsDwFlags As Long         ' Reserved
    dsBmBits As LongPtr       ' **Pointer to the actual pixel data**
End Type


'***********
'*  State  *
'***********
Private Type TThisCache
  HICON As LongPtr
End Type
Private Type TThis
  hBitmap As LongPtr
  HBITMAPTYPE As BITMAP
  Cache As TThisCache
End Type
Private This As TThis


'****************
'* Constructors *
'****************

'Creates a stdImage object from a byte array
'@constructor
'@param b - Non-empty encoded image bytes; any array lower bound is supported.
'@remark The first frame is rasterised against black; animation, vector data, metadata, and alpha are not retained.
'@param format - Supported format hint retained for compatibility. GDI+ detects the actual encoded format from the data.
'@returns - Picture object representing data
'@example ```vb
'  Dim b() As Byte: b = stdImage.CreateFromFile("H:\source.png").ToBinary(stdImgFormatPNG)
'  Call stdImage.CreateFromBinary(b).ToFile("H:\copy.png")
'```
Public Function CreateFromBinary(ByRef b() As Byte, Optional ByVal format As stdImgFormat = stdImgFormatDefault) As stdImage
  On Error GoTo Catch
  Call validateInputFormat(format)
  Dim byteCount As Long
  byteCount = UBound(b) - LBound(b) + 1
  If byteCount <= 0 Then Err.Raise 5, "stdImage::CreateFromBinary", "Byte array is empty"
  Dim hMem As LongPtr, pMem As LongPtr, stream As stdole.IUnknown
  hMem = GlobalAlloc(&H2, byteCount)
  If hMem = 0 Then Err.Raise 7, "stdImage::CreateFromBinary", "Could not allocate stream memory"
  pMem = GlobalLock(hMem)
  If pMem = 0 Then Err.Raise 7, "stdImage::CreateFromBinary", "Could not lock stream memory"
  Call CopyMemory(pMem, VarPtr(b(LBound(b))), byteCount)
  Call GlobalUnlock(hMem)
  pMem = 0
  If CreateStreamOnHGlobal(hMem, 1, stream) <> 0 Then Err.Raise 7, "stdImage::CreateFromBinary", "Could not create input stream"
  hMem = 0 'The stream owns the allocation from this point.
  Set CreateFromBinary = loadImage(vbNullString, stream, 0)
  Exit Function
Catch:
  Dim errorNumber As Long, errorSource As String, errorDescription As String
  errorNumber = Err.Number: errorSource = Err.Source: errorDescription = Err.Description
  If pMem <> 0 Then Call GlobalUnlock(hMem)
  If hMem <> 0 Then Call GlobalFree(hMem)
  Set stream = Nothing
  If errorNumber = 9 Then Err.Raise 5, "stdImage::CreateFromBinary", "Byte array is empty"
  Err.Raise errorNumber, errorSource, errorDescription
End Function

'Creates a stdImage object from a Data URL string
'@constructor
'@param sDataURL - A supported image MIME type with a Base64 payload (e.g., "data:image/png;base64,iVBORw0KGgo...")
'@returns - A new stdImage object with transparency flattened against black.
'@remark Non-Base64 Data URLs and unsupported MIME types are rejected.
'@example ```vb
'  Dim dataUrl as String
'  dataUrl = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg=="
'  stdImage.CreateFromDataURL(dataUrl).ToFile "C:\Temp\FromDataUrl.png"
'```
Public Function CreateFromDataURL(ByVal sDataURL As String) As stdImage
  If LCase$(Left$(sDataURL, 5)) <> "data:" Then Err.Raise 5, "stdImage::CreateFromDataURL", "Expected a data: URL"
  Dim comma As Long: comma = InStr(1, sDataURL, ",")
  If comma = 0 Then Err.Raise 5, "stdImage::CreateFromDataURL", "Missing comma separator"
  Dim header As String: header = LCase$(Mid$(sDataURL, 6, comma - 6))
  If Right$(header, 7) <> ";base64" Then Err.Raise 5, "stdImage::CreateFromDataURL", "Only Base64 image Data URLs are supported"
  Dim headerParts As Variant: headerParts = Split(header, ";")
  Dim mime As String: mime = headerParts(0)
  Select Case mime
    Case "image/bmp", "image/x-ms-bmp", "image/png", "image/gif", "image/jpeg", "image/tiff", "image/wmf", "image/x-wmf", "image/emf", "image/x-emf", "image/x-icon", "image/vnd.microsoft.icon"
    Case "image/webp": Err.Raise 5, "stdImage::CreateFromDataURL", "WEBP format is not currently supported (see issue #178)"
    Case Else: Err.Raise 5, "stdImage::CreateFromDataURL", "Unsupported image MIME type"
  End Select
  Dim encoded As String: encoded = Mid$(sDataURL, comma + 1)
  'Allow folded Base64, but reject empty payloads and invalid padding.
  encoded = Replace(Replace(Replace(Replace(encoded, vbCr, ""), vbLf, ""), vbTab, ""), " ", "")
  If Len(encoded) = 0 Or Len(encoded) Mod 4 <> 0 Then Err.Raise 5, "stdImage::CreateFromDataURL", "Invalid Base64 payload"
  Const CRYPT_STRING_BASE64_STRICT As Long = &H20000001
  Dim byteCount As Long, data() As Byte
  If CryptStringToBinaryA(encoded, Len(encoded), CRYPT_STRING_BASE64_STRICT, ByVal 0&, byteCount, 0, 0) = 0 Then Err.Raise 5, "stdImage::CreateFromDataURL", "Invalid Base64 payload"
  If byteCount <= 0 Then Err.Raise 5, "stdImage::CreateFromDataURL", "Empty Base64 payload"
  ReDim data(0 To byteCount - 1)
  If CryptStringToBinaryA(encoded, Len(encoded), CRYPT_STRING_BASE64_STRICT, data(0), byteCount, 0, 0) = 0 Then Err.Raise 5, "stdImage::CreateFromDataURL", "Invalid Base64 payload"
  Set CreateFromDataURL = CreateFromBinary(data)
End Function

'Creates a stdImage object from a file
'@constructor
'@param sPath - Path of a BMP, PNG, GIF, JPEG, TIFF, WMF, EMF, or ICO file.
'@param format - Supported format hint retained for compatibility. GDI+ detects the actual encoded format from the data.
'@param transparentColor - Background colour used to flatten transparency. The default is black.
'@remark The first frame is rasterised; animation, vector data, metadata, and alpha are not retained.
'@returns - Picture object representing data
'@example ```vb
'stdImage.CreateFromFile("H:\test1.bmp").ToFile("H:\test1.png")
'```
Public Function CreateFromFile(ByVal sPath As String, Optional ByVal format As stdImgFormat = stdImgFormatDefault, Optional ByVal transparentColor As Long = 0) As stdImage
  Call validateInputFormat(format)
  If format = stdImgFormatDefault Then
    Select Case LCase$(Mid$(sPath, InStrRev(sPath, ".") + 1))
      Case "wmf", "emf", "ico" 'These input-only formats are rasterised by GDI+.
      Case Else
        Dim inferred As stdImgFormat: inferred = formatFromExtension(sPath)
    End Select
  End If
  Set CreateFromFile = loadImage(sPath, Nothing, transparentColor)
End Function

'Creates a stdImage object from a url
'@constructor
'@param sUrl - Url of file
'@param format - Supported format hint retained for compatibility. GDI+ detects the actual encoded format from the data.
'@returns - Picture object representing data
'@example ```vb
'stdImage.CreateFromURL("https://raw.githubusercontent.com/sancarn/awesome-vba/main/resources/VBALogo.png").ToFile("H:\test.png", stdImgFormatPNG)
'```
'@remark Uses the default WinHTTP proxy configuration. For custom proxy, authentication, or timeout handling, use stdHTTP and pass its response bytes to CreateFromBinary.
Public Function CreateFromUrl(ByVal sURL As String, Optional ByVal format As stdImgFormat = stdImgFormatDefault) As stdImage
  Call validateInputFormat(format)
  Dim HTTP As Object: Set HTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
  HTTP.Open "GET", sURL, False
  HTTP.send
  If HTTP.status <> 200 Then Err.Raise 1, "stdImage::CreateFromUrl", "Could not get image (HTTP status: " & HTTP.status & ")"
  Dim data() As Byte: data = HTTP.responseBody
  Set CreateFromUrl = CreateFromBinary(data, format)
End Function

'Creates a stdImage object from a file
'@constructor
'@param hBitmap - Valid bitmap handle. Ownership transfers to the returned stdImage; do not free it or pass a borrowed handle.
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
  On Error GoTo Catch
  Dim isOpen As Boolean, copy As LongPtr
  If Not OpenClipboardTimeout(getAppHwnd()) Then Err.Raise 1, "stdImage::CreateFromClipboard", "Unable to open clipboard"
  isOpen = True
  If IsClipboardFormatAvailable(CF_BITMAP) = 0 Then Err.Raise 5, "stdImage::CreateFromClipboard", "No bitmap format in clipboard"
  copy = copyBitmap(GetClipboardData(CF_BITMAP))
  Set CreateFromClipboard = CreateFromHBitmap(copy)
  copy = 0
  If CloseClipboard() = 0 Then Err.Raise 1, "stdImage::CreateFromClipboard", "Unable to close clipboard"
  isOpen = False
  Exit Function
Catch:
  Dim errorNumber As Long, errorSource As String, errorDescription As String
  errorNumber = Err.Number: errorSource = Err.Source: errorDescription = Err.Description
  If isOpen Then Call CloseClipboard
  If copy <> 0 Then Call DeleteObject(copy)
  Err.Raise errorNumber, errorSource, errorDescription
End Function

'Creates a stdImage object from a stdole.IPictureDisp object
'@constructor
'@param picture - Object to create picture from
'@returns - Picture object representing data
'@example `Dim picture As stdole.IPictureDisp: Set picture = LoadPicture("C:\Temp\photo.bmp"): Dim img As stdImage: Set img = stdImage.CreateFromIPictureDisp(picture)`
Public Function CreateFromIPictureDisp(ByVal picture As stdole.IPictureDisp) As stdImage
  If picture Is Nothing Then Err.Raise 5, "stdImage::CreateFromIPictureDisp", "Picture is Nothing"
  Select Case picture.Type
    Case 1
      Set CreateFromIPictureDisp = imageFromBitmapCopy(picture.handle)
    Case 3
      Set CreateFromIPictureDisp = CreateFromHICON(picture.handle)
    Case Else
      Err.Raise 5, "stdImage::CreateFromIPictureDisp", "Only bitmap and icon pictures are supported"
  End Select
End Function

'Creates a stdImage object from a stdole.stdPicture object
'@constructor
'@param picture - Object to create picture from
'@returns - Picture object representing data
'@example ```vb
'stdImage.CreateFromStdPicture(Userform1.Picture).ToFile("H:\test.png", stdImgFormatPNG)
'```
Public Function CreateFromStdPicture(ByVal picture As stdole.StdPicture) As stdImage
  If picture Is Nothing Then Err.Raise 5, "stdImage::CreateFromStdPicture", "Picture is Nothing"
  Select Case picture.Type
    Case 1
      Set CreateFromStdPicture = imageFromBitmapCopy(picture.handle)
    Case 3
      Set CreateFromStdPicture = CreateFromHICON(picture.handle)
    Case Else
      Err.Raise 5, "stdImage::CreateFromStdPicture", "Only bitmap and icon pictures are supported"
  End Select
End Function

'Creates a stdImage object from a stdole.IPicture object
'@constructor
'@param picture - Object to create picture from
'@returns - Picture object representing data
'@example `Dim picture As stdole.IPicture: Set picture = LoadPicture("C:\Temp\photo.bmp"): Dim img As stdImage: Set img = stdImage.CreateFromIPicture(picture)`
Public Function CreateFromIPicture(ByVal picture As stdole.IPicture) As stdImage
  If picture Is Nothing Then Err.Raise 5, "stdImage::CreateFromIPicture", "Picture is Nothing"
  Select Case picture.Type
    Case 1
      Set CreateFromIPicture = imageFromBitmapCopy(picture.handle)
    Case 3
      Set CreateFromIPicture = CreateFromHICON(picture.handle)
    Case Else
      Err.Raise 5, "stdImage::CreateFromIPicture", "Only bitmap and icon pictures are supported"
  End Select
End Function

'Creates a stdImage object from a Excel.IPicture object
'@constructor
'@param picture as Object<Excel.Picture> - Object to create picture from
'@returns - Picture object representing data
'@example `Dim img As stdImage: Set img = stdImage.CreateFromExcelIPicture(ActiveSheet.Pictures(1))`
Public Function CreateFromExcelIPicture(ByVal picture As Object) As stdImage
  Call picture.CopyPicture(, 2)
  Set CreateFromExcelIPicture = CreateFromClipboard()
End Function

'Creates a stdImage object from a Shape object
'@constructor
'@param shp as Object<Excel.Shape> - Object to create picture from
'@returns - Picture object representing data
'@example ```vb
'stdImage.CreateFromShape(Sheet1.Shapes("Picture 2")).ToFile("H:\test.png")
'```
Public Function CreateFromShape(ByVal shp As Object) As stdImage
  Select Case Application.Name
    Case "Microsoft Excel"
      Dim xlShp As Object: Set xlShp = shp
      Call xlShp.CopyPicture(, 2)
      Set CreateFromShape = stdImage.CreateFromClipboard()
  End Select
End Function

'Creates a stdImage object from a Chart object
'@constructor
'@param chart as Object<Excel.Chart> - Object to create picture from
'@returns - Picture object representing data
'@example ```vb
'stdImage.CreateFromChart(Sheet1.ChartObjects(1).Chart).ToFile("H:\test.png")
'```
Public Function CreateFromChart(ByVal chart As Object) As stdImage
  Select Case Application.Name
    Case "Microsoft Excel"
      Dim xlChart As Object: Set xlChart = chart
      Call xlChart.CopyPicture(, 2)
      Set CreateFromChart = stdImage.CreateFromClipboard()
  End Select
End Function

'Creates a stdImage object from the screen/desktop window (screenshot)
'@constructor
'@returns - Picture object containing screenshot data
'@example ```vb
'stdImage.CreateFromScreen().ToFile("H:\test.png")
'```
'@TODO: Capture both monitors https://stackoverflow.com/a/34445431/6302131
Public Function CreateFromScreen() As stdImage
  Dim hDesktop As LongPtr: hDesktop = GetDesktopWindow()
  
  Const SM_XVIRTUALSCREEN = 76
  Const SM_YVIRTUALSCREEN = 77
  Const SM_CXVIRTUALSCREEN = 78
  Const SM_CYVIRTUALSCREEN = 79
  Dim ScreenX As Long: ScreenX = GetSystemMetrics(SM_XVIRTUALSCREEN)
  Dim ScreenY As Long: ScreenY = GetSystemMetrics(SM_YVIRTUALSCREEN)
  Dim ScreenWidth As Long: ScreenWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN)
  Dim ScreenHeight As Long: ScreenHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN)
  
  Set CreateFromScreen = CreateFromWindow(hDesktop, region:=Array(ScreenX, ScreenY, ScreenX + ScreenWidth, ScreenY + ScreenHeight))
End Function

'Creates a stdImage object from a window
'@constructor
'@param hwnd - Handle to window
'@param bClientArea - If true, only the client area of the window is captured. If false, the entire window is captured.
'@param region - If empty, the region is obtained from the window dimensions. Else the region should be an array of longs Array(Left, Top, Right, Bottom)
'@returns - Picture object representing data
'@remark - This method still doesn't always work on all windows. Unfortunately for instance chrome windows render with `DirectX` or `OpenGL`, whos buffers cannot be accessed
'natively with GDI. An alternative is provided in the examples section of this module.
'@TODO: Investigate usage of `DirectX Duplication API` - can we get screenshots in a more direct manor? [Example](https://github.com/smasherprog/screen_capture_lite/blob/master/src_cpp/windows/DXFrameProcessor.cpp)
'@TODO: Option to include cursor https://stackoverflow.com/a/48925443/6302131
#If VBA7 Then
  Public Function CreateFromWindow(ByVal hwnd As LongPtr, Optional ByVal bClientArea As Boolean = False, Optional ByVal region As Variant = Empty) As stdImage
#Else
  Public Function CreateFromWindow(ByVal hwnd As Long, Optional ByVal bClientArea As Boolean = False, Optional ByVal region As Variant = Empty) As stdImage
#End If
  Const SRCCOPY As Long = &HCC0020
  
  'Get the proper Device Context (DC) depending on the value of bClient.
  'Additionally get the rect of the window/client.
  Dim hDCSource As LongPtr
  If bClientArea Then
    hDCSource = GetDC(hwnd)
  Else
    hDCSource = GetWindowDC(hwnd)
  End If
    
  Dim tRect As RECT
  If IsEmpty(region) Then
    If bClientArea Then
      Call GetClientRect(hwnd, tRect)
    Else
      Call GetWindowRect(hwnd, tRect)
    End If
  Else
    'region => Array(Left, Top, Right, Bottom)
    With tRect
      .Left = region(0)
      .Top = region(1)
      .Right = region(2)
      .Bottom = region(3)
    End With
  End If
  
  'Create an in-memory DC for the copy process
  Dim hCaptureDC As LongPtr: hCaptureDC = CreateCompatibleDC(hDCSource)

  'Create a bitmap to capture the window
  Dim hCaptureBitmap As LongPtr: hCaptureBitmap = CreateCompatibleBitmap(hDCSource, tRect.Right - tRect.Left, tRect.Bottom - tRect.Top)
  
  'Load bitmap into in-memory DC
  Dim hOldBitmap As LongPtr: hOldBitmap = SelectObject(hCaptureDC, hCaptureBitmap)

  'Copy the on-screen image into the memory DC.
  Dim bSuccess As Long: bSuccess = PrintWindow(hwnd, hCaptureDC, 0&)
  If bSuccess = 0 Then
    'Sometimes `PrintWindow` fails... Try again with bitblt here
    bSuccess = BitBlt(hCaptureDC, 0, 0, tRect.Right - tRect.Left, tRect.Bottom - tRect.Top, hDCSource, 0, 0, SRCCOPY)
    If bSuccess = 0 Then
      Err.Raise 1, "", "Cannot capture image. Error: " & Hex(Err.LastDllError)
    End If
  End If
  
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
  Public Function CreateFromHICON(ByVal HICON As LongPtr) As stdImage
#Else
  Public Function CreateFromHICON(ByVal HICON As Long) As stdImage
#End If
  'Get icon information from HICON
  Dim iconInf As ICONINFO: If GetIconInfo(HICON, iconInf) = 0 Then Err.Raise 1, "stdImage::CreateFromHICON", "Could not get icon information"
  
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
  Dim r As Long: r = DrawIconEx(hCaptureDC, 0, 0, HICON, bmp.bmWidth, bmp.bmHeight, 0, 0, DI_NORMAL)

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

'@param x - The x coord to start the crop at
'@param y - The y coord to start the crop at
'@param width - The width of the new cropped image
'@param height - The height of the new cropped image
Public Function Crop(ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long) As stdImage
  Dim hDCScreen As LongPtr: hDCScreen = GetDC(0) ' Get screen DC to create compatible ones
  Dim hDCOriginal As LongPtr: hDCOriginal = CreateCompatibleDC(hDCScreen)
  Dim hDCCropped As LongPtr: hDCCropped = CreateCompatibleDC(hDCScreen)

  Dim hOldBitmapOriginal As LongPtr
  Dim hOldBitmapCropped As LongPtr
  Dim hBitmapCropped As LongPtr

  ' Select the original bitmap into a memory DC
  hOldBitmapOriginal = SelectObject(hDCOriginal, This.hBitmap)

  ' Create a new bitmap for the cropped image
  hBitmapCropped = CreateCompatibleBitmap(hDCScreen, width, height)
  hOldBitmapCropped = SelectObject(hDCCropped, hBitmapCropped)

  ' Copy the specified rectangle from the original bitmap to the new cropped bitmap
  ' Source (original) coordinates start at x, y
  ' Destination (cropped) coordinates start at 0, 0
  Const SRCCOPY As Long = &HCC0020 ' Standard copy flag
  Dim r As Long: r = BitBlt(hDCCropped, 0, 0, width, height, hDCOriginal, x, y, SRCCOPY)

  ' Deselect bitmaps from DCs
  Call SelectObject(hDCOriginal, hOldBitmapOriginal)
  Call SelectObject(hDCCropped, hOldBitmapCropped)

  ' Clean up DCs
  Call DeleteDC(hDCCropped)
  Call DeleteDC(hDCOriginal)
  Call ReleaseDC(0, hDCScreen)

  If r = 0 Then
    ' BitBlt failed, clean up the created bitmap and raise an error
    Call DeleteObject(hBitmapCropped)
    Err.Raise 1, "stdImage::Crop", "Failed to crop image using BitBlt."
    Set Crop = Nothing
  Else
    ' Create a new stdImage object from the cropped HBITMAP
    Set Crop = New stdImage ' Assuming stdImage has a default constructor
    ' You'll need a private/friend method in stdImage to set its hBitmap from an external handle
    ' Or, if CreateFromHBitmap is public static/class method, use it:
    Set Crop = CreateFromHBitmap(hBitmapCropped) ' Assuming CreateFromHBitmap is accessible
                                                 ' and handles ownership/deletion of hBitmap
  End If
End Function

'Initialised stdImage
'@protected
'@param hBITMAP - Handle to bitmap
Public Sub protInit(ByVal hBitmap As LongPtr)
  Call GetObjectAPI(hBitmap, LenB(This.HBITMAPTYPE), This.HBITMAPTYPE)
  This.hBitmap = hBitmap
  This.Cache.HICON = 0
End Sub

'Obtain a HBitmap
'@returns - Borrowed bitmap handle, valid for the lifetime of this image.
'@remark Do not delete this handle. stdImage owns it and frees it on termination.
Public Property Get hBitmap() As LongPtr
  hBitmap = This.hBitmap
End Property

'Obtain (and cache) an HICON
Public Property Get HICON() As LongPtr
  If This.Cache.HICON = 0 Then
    Dim iconInf As ICONINFO
    iconInf.fIcon = 1
    iconInf.hbmMask = This.hBitmap
    iconInf.hbmColor = This.hBitmap
    This.Cache.HICON = CreateIconIndirect(iconInf)
  End If
  HICON = This.Cache.HICON
End Property

'Obtain the colors of the image
'@returns ARGB() - 2D array of colors. Colors ordered from (1,1) = bottom-left, (1,height) = top-left, across to (width, height) = top-right.
Public Property Get Colors() As Long()
  Const BI_RGB = 0
  With This.HBITMAPTYPE
    'Obtain colors
    Dim ret() As Long
    ReDim ret(1 To .bmWidth, 1 To .bmHeight)
    'If .bmBits = NULL_PTR Then Err.Raise 1, "", "Image is empty?"
    'Call CopyMemory(VarPtr(ret(1, 1)), .bmBits, .bmWidth * .bmHeight * 4)
    
    Dim hDC As LongPtr: hDC = GetDC(NULL_PTR)
    Dim lBitsPerPixel As Long: lBitsPerPixel = 32
    Dim lWidthBytes As Long: lWidthBytes = (((.bmWidth * lBitsPerPixel) + 31) \ 32) * 4
    Dim lImageSize As Long: lImageSize = lWidthBytes * Abs(.bmHeight)
    
    Dim tBmpInf As BITMAPINFO
    With tBmpInf.bmiheader
        .biSize = LenB(tBmpInf.bmiheader)
        .biWidth = This.HBITMAPTYPE.bmWidth
        .biHeight = This.HBITMAPTYPE.bmHeight
        .biPlanes = 1
        .biBitCount = lBitsPerPixel
        .biCompression = BI_RGB ' No compression
        .biSizeImage = lImageSize
    End With
    
    If GetDIBits(hDC, hBitmap, 0, .bmHeight, ret(1, 1), tBmpInf, 0&) = 0 Then
      Call ReleaseDC(NULL_PTR, hDC)
      Err.Raise 1, "", "Could not get pixels?"
    End If
    
    Call ReleaseDC(NULL_PTR, hDC)
    
    Colors = ret
  End With
End Property

'Obtain the Image Data
Public Property Get Typee() As Long: Typee = This.HBITMAPTYPE.bmType: End Property                      'Type cannot be used As a name, thats why its named Typee
Public Property Get width() As Long: width = This.HBITMAPTYPE.bmWidth: End Property
Public Property Get height() As Long: height = This.HBITMAPTYPE.bmHeight: End Property
Public Property Get WidthBytes() As Long: WidthBytes = This.HBITMAPTYPE.bmWidthBytes: End Property
Public Property Get Planes() As Integer: Planes = This.HBITMAPTYPE.bmPlanes: End Property
Public Property Get BitsPixel() As Integer: BitsPixel = This.HBITMAPTYPE.bmBitsPixel: End Property
Public Property Get Bits() As LongPtr: Bits = This.HBITMAPTYPE.bmBits: End Property


'Obtain binary data of the image in any format required
'@param eFormat - Output format. The default is PNG.
'@returns - Binary data of image, ready to load into a file
Public Function ToBinary(Optional ByVal eFormat As stdImgFormat = stdImgFormatDefault) As Byte()
  On Error GoTo Catch
  eFormat = resolveOutputFormat(eFormat)
  Dim stream As stdole.IUnknown
  If CreateStreamOnHGlobal(0, 1, stream) <> 0 Then Err.Raise 7, "stdImage::ToBinary", "Could not create output stream"
  Call saveImage(eFormat, vbNullString, stream, 80, 24, 6)

  'GlobalSize can include allocation padding; use the logical stream length.
  Dim streamSize As Currency, byteCount As Long
  If IStream_Size(stream, streamSize) <> 0 Then Err.Raise 1, "stdImage::ToBinary", "Could not obtain stream size"
  If streamSize <= 0 Or streamSize > 214748.3647@ Then Err.Raise 6, "stdImage::ToBinary", "Image data exceeds supported array size"
  byteCount = CLng(streamSize * 10000@)
  Dim hMem As LongPtr, pMem As LongPtr
  If GetHGlobalFromStream(stream, hMem) <> 0 Then Err.Raise 1, "stdImage::ToBinary", "Could not obtain stream memory"
  pMem = GlobalLock(hMem)
  If pMem = 0 Then Err.Raise 7, "stdImage::ToBinary", "Could not lock stream memory"
  Dim result() As Byte: ReDim result(0 To byteCount - 1)
  Call CopyMemory(VarPtr(result(0)), pMem, byteCount)
  Call GlobalUnlock(hMem)
  pMem = 0
  ToBinary = result
  Exit Function
Catch:
  Dim errorNumber As Long, errorSource As String, errorDescription As String
  errorNumber = Err.Number: errorSource = Err.Source: errorDescription = Err.Description
  If pMem <> 0 Then Call GlobalUnlock(hMem)
  Set stream = Nothing
  Err.Raise errorNumber, errorSource, errorDescription
End Function

'Save image as file
'@param FileName - Path of the output file.
'@param format - Output format. The default infers BMP, PNG, GIF, JPEG, or TIFF from the filename extension.
'@param Quality - JPEG quality from 0 to 100.
'@param TIFF_ColorDepth - TIFF colour depth passed to the encoder.
'@param TIFF_Compression - TIFF compression value passed to the encoder.
Public Sub ToFile(ByVal FileName As String, Optional format As stdImgFormat = stdImgFormatDefault, Optional ByVal Quality As Byte = 80, Optional ByVal TIFF_ColorDepth As Long = 24, Optional ByVal TIFF_Compression As Long = 6)
  If format = stdImgFormatDefault Then format = formatFromExtension(FileName)
  Call saveImage(resolveOutputFormat(format), FileName, Nothing, CLng(Quality), TIFF_ColorDepth, TIFF_Compression)
End Sub

'Copy the image to the clipboard
'@example `stdImage.CreateFromFile("H:\test1.bmp").ToClipboard()
Public Sub ToClipboard()
  On Error GoTo Catch
  Dim copy As LongPtr, isOpen As Boolean
  copy = copyBitmap(This.hBitmap)
  If Not OpenClipboardTimeout(getAppHwnd()) Then Err.Raise 1, "stdImage::ToClipboard", "Unable to open clipboard"
  isOpen = True
  If EmptyClipboard() = 0 Then Err.Raise 1, "stdImage::ToClipboard", "Unable to empty clipboard"
  If SetClipboardData(CF_BITMAP, copy) = 0 Then Err.Raise 1, "stdImage::ToClipboard", "Unable to set clipboard bitmap"
  copy = 0 'The clipboard owns this copy after SetClipboardData succeeds.
  If CloseClipboard() = 0 Then Err.Raise 1, "stdImage::ToClipboard", "Unable to close clipboard"
  isOpen = False
  Exit Sub
Catch:
  Dim errorNumber As Long, errorSource As String, errorDescription As String
  errorNumber = Err.Number: errorSource = Err.Source: errorDescription = Err.Description
  If isOpen Then Call CloseClipboard
  If copy <> 0 Then Call DeleteObject(copy)
  Err.Raise errorNumber, errorSource, errorDescription
End Sub


'Save a copy in a temporary file and return the path of this file
'@param eFormat - Output format. The default is PNG.
'@returns - The path to the temporary file. The caller is responsible for deleting it.
Public Function ToTempFile(Optional ByVal eFormat As stdImgFormat = stdImgFormatDefault) As String
  eFormat = resolveOutputFormat(eFormat)
  Dim extensions As Variant: extensions = Array("", "bmp", "png", "gif", "jpeg", "tiff")
  Dim sPath As String: sPath = Environ("Temp") & "\" & getGUID() & "." & extensions(eFormat)
  Call ToFile(sPath, eFormat)
  ToTempFile = sPath
End Function

'Adds the picture to the sheet
'@param ws as Object<Excel.Worksheet> - Worksheet to add image to. Late bound so other Office hosts do not require an Excel reference.
'@param Top - Top coord of the newly placed image
'@param Left - Left coord of the newly placed image
'@param compress - Excel MsoPictureCompress value; 0 disables compression (msoPictureCompressFalse).
'@returns Object<Picture> - Picture object
Public Function ToSheet(ByVal ws As Object, Optional ByVal Top As Double = 0, Optional ByVal Left As Double = 0, Optional ByVal compress As Long = 0) As Object
  On Error GoTo Catch
  Dim path As String: path = ToTempFile()
  Set ToSheet = ws.Shapes.AddPicture2(path, False, True, Left, Top, -1, -1, compress)
  Kill path
  Exit Function
Catch:
  Dim errorNumber As Long, errorSource As String, errorDescription As String
  errorNumber = Err.Number: errorSource = Err.Source: errorDescription = Err.Description
  On Error Resume Next
  If Len(path) > 0 Then Kill path
  On Error GoTo 0
  Err.Raise errorNumber, errorSource, errorDescription
End Function

'Update a UI control's background picture
'@param oControl as Object<Control> - The control to update the picture of
'@example ```vb
'stdImage.CreateFromFile("H:\test1.bmp").ToUIControl(UserForm1.Image1)
'```
Public Sub ToUIControl(ByVal oControl As Object)
  Set oControl.picture = ToStdPicture()
End Sub

'Obtain a data URL of the image
'@param eFormat - Output format. The default is PNG.
'@returns - A data:image/...;base64 URL without line breaks.
'@example `Debug.Print stdImage.CreateFromFile("C:\Temp\photo.jpg").ToDataUrl()`
Public Function ToDataUrl(Optional ByVal eFormat As stdImgFormat = stdImgFormatDefault) As String
  eFormat = resolveOutputFormat(eFormat)
  Dim mimeTypes As Variant: mimeTypes = Array("", "image/bmp", "image/png", "image/gif", "image/jpeg", "image/tiff")
  Dim data() As Byte: data = ToBinary(eFormat)
  Const CRYPT_STRING_BASE64_NOCRLF As Long = &H40000001
  Dim charCount As Long
  If CryptBinaryToStringA(data(0), UBound(data) + 1, CRYPT_STRING_BASE64_NOCRLF, vbNullString, charCount) = 0 Then Err.Raise 1, "stdImage::ToDataUrl", "Could not size Base64 output"
  Dim encoded As String: encoded = String$(charCount, vbNullChar)
  If CryptBinaryToStringA(data(0), UBound(data) + 1, CRYPT_STRING_BASE64_NOCRLF, encoded, charCount) = 0 Then Err.Raise 1, "stdImage::ToDataUrl", "Could not encode Base64 data"
  ToDataUrl = "data:" & mimeTypes(eFormat) & ";base64," & Left$(encoded, charCount)
End Function

'Output the image to a shape's fill
'@param shp as Object<Excel.Shape> - The shape to fill
'@example ```vb
'stdImage.CreateFromFile("H:\test1.bmp").ToShapeFill(ActiveSheet.Shapes(1))
'```
Public Sub ToShapeFill(ByVal shp As Object)
  On Error GoTo Catch
  If Application.Name <> "Microsoft Excel" Then Err.Raise 5, "stdImage::ToShapeFill", "Only Excel shapes are supported"
  Dim path As String: path = ToTempFile()
  Call shp.Fill.UserPicture(path)
  Kill path
  Exit Sub
Catch:
  Dim errorNumber As Long, errorSource As String, errorDescription As String
  errorNumber = Err.Number: errorSource = Err.Source: errorDescription = Err.Description
  On Error Resume Next
  If Len(path) > 0 Then Kill path
  On Error GoTo 0
  Err.Raise errorNumber, errorSource, errorDescription
End Sub

'Obtain a stdole.StdPicture
'@returns - An independent stdole.StdPicture which remains valid after this image is released.
Public Function ToStdPicture() As stdole.StdPicture
  On Error GoTo Catch
  Dim iid As GUID, descriptor As PICTDESC, copy As LongPtr
  If IIDFromString(StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), iid) <> 0 Then Err.Raise 1, "stdImage::ToStdPicture", "Invalid picture interface identifier"
  copy = copyBitmap(This.hBitmap)
  descriptor.Size = LenB(descriptor)
  descriptor.type = 1
  descriptor.hPic = copy
  If OleCreatePictureIndirect(descriptor, iid, 1, ToStdPicture) <> 0 Then Err.Raise 1, "stdImage::ToStdPicture", "Could not create StdPicture"
  copy = 0 'The returned picture owns its independent bitmap.
  Exit Function
Catch:
  Dim errorNumber As Long, errorSource As String, errorDescription As String
  errorNumber = Err.Number: errorSource = Err.Source: errorDescription = Err.Description
  If copy <> 0 Then Call DeleteObject(copy)
  Err.Raise errorNumber, errorSource, errorDescription
End Function

Private Sub Class_Terminate()
  'Free HICON
  If This.Cache.HICON <> 0 Then Call DestroyIcon(This.Cache.HICON)
  
  'Free HBITMAP
  If This.hBitmap <> 0 Then Call DeleteObject(This.hBitmap)
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
Private Function BytesFromPointer(ByVal pointer As LongPtr, ByVal Size As Long) As Byte()
  Dim buff() As Byte
  If Size > 0 Then
    ReDim buff(1 To Size)
    Call CopyMemory(VarPtr(buff(1)), pointer, Size)
  Else
    Call Err.Raise(1, "BytesFromPointer [PRIVATE]", "No size supplied.")
  End If
  BytesFromPointer = buff
End Function

'Tries to open the clipbaord with a timeout
'@param hWnd - hWnd to open clipboard in
'@param dwTimeoutMilliseconds - number of milliseconds to wait
'@returns - true if clipboard is open, false if clipboard couldn't be opened
Private Function OpenClipboardTimeout(ByVal hwnd As LongPtr, Optional ByVal dwTimeoutMilliseconds As Long = 5000) As Boolean
  Dim iStart As Long: iStart = GetTickCount()
  Dim bOpen As Boolean, bExpired As Boolean
  While Not (bOpen Or bExpired)
    'Open clipboard and only continue if open
    bOpen = CBool(OpenClipboard(hwnd))
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
  'Late bind host-specific members so the class imports without an Excel reference.
  Dim oApp As Object: Set oApp = Application
  Dim hwnd As LongPtr
  On Error Resume Next
  hwnd = oApp.hwnd
  If hwnd = 0 Then hwnd = oApp.ActiveWindow.hwnd
  If hwnd = 0 Then hwnd = oApp.hWndAccessApp
  If hwnd = 0 Then hwnd = GetActiveWindow()
  On Error GoTo 0
  If hwnd = 0 Then Err.Raise 1, "stdImage", "Could not obtain an Office window for clipboard ownership"
  getAppHwnd = hwnd
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
Public Function FindFirst(ByVal template As stdImage, Optional ByVal tolerance As Long = 5, Optional ByVal TopToBottom As Boolean = True, Optional ByVal LeftToRight As Boolean = True) As Variant
  Dim haystackColors() As Long: haystackColors = Me.Colors
  Dim templateColors() As Long: templateColors = template.Colors
  Dim hayW As Long: hayW = UBound(haystackColors, 1)
  Dim hayH As Long: hayH = UBound(haystackColors, 2)
  Dim tplW As Long: tplW = UBound(templateColors, 1)
  Dim tplH As Long: tplH = UBound(templateColors, 2)

  Dim startY As Long, endY As Long, stepY As Long
  If TopToBottom Then
    startY = 0: endY = hayH - tplH: stepY = 1
  Else
    startY = hayH - tplH: endY = 0: stWepY = -1
  End If
  
  Dim startX As Long, endX As Long, stepX As Long
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
        FindFirst = Array(x, hayH - y - tplH + 1)
        Exit Function
      End If
    Next
  Next
  FindFirst = Empty
End Function

'Searches for a template image within a larger image (i.e., this image)
'@param template - The image to search for
'@param tolerance - A value of 0 means that the pixels must be an exact match; greater values signify slight (1-5), moderate (6 - 20) and signifanct (20+) variations (default 5)
'@param TopToBottom - If True, search from top-to-bottom; if False, search from bottom-to-top (default True)
'@param LeftToRight - If True, search from left-to-right; if False, search from right-to-left (default True)
'@returns Collection<Array<Long>> - Collection of arrays [x, y] of the top-left matches (standard image coordinates), or empty collection if not found
'@example
'  Dim haystack as stdImage: set haystack = stdImage.CreateFromWindow(Application.hWnd)
'  Dim needle as stdImage: set needle = stdImage.CreateFrom
Public Function FindAll(ByVal template As stdImage, Optional ByVal tolerance As Long = 5, Optional ByVal TopToBottom As Boolean = True, Optional ByVal LeftToRight As Boolean = True) As Collection
  Set FindAll = New Collection
  
  Dim haystackColors() As Long: haystackColors = Me.Colors
  Dim templateColors() As Long: templateColors = template.Colors
  Dim hayW As Long: hayW = UBound(haystackColors, 1)
  Dim hayH As Long: hayH = UBound(haystackColors, 2)
  Dim tplW As Long: tplW = UBound(templateColors, 1)
  Dim tplH As Long: tplH = UBound(templateColors, 2)

  Dim startY As Long, endY As Long, stepY As Long
  If TopToBottom Then
    startY = 0: endY = hayH - tplH: stepY = 1
  Else
    startY = hayH - tplH: endY = 0: stepY = -1
  End If
  
  Dim startX As Long, endX As Long, stepX As Long
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
        Next i
        If Not match Then Exit For
      Next j
      
      If match Then
        Call FindAll.add(Array(x, hayH - y - tplH + 1))
      End If
    Next x
  Next y
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
'Validate an output format, resolving the default to PNG.
Private Function resolveOutputFormat(ByVal format As stdImgFormat) As stdImgFormat
  If format = stdImgFormatDefault Then format = stdImgFormatPNG
  If format < stdImgFormatBMP Or format > stdImgFormatTIFF Then Err.Raise 5, "stdImage", "Unsupported output format"
  resolveOutputFormat = format
End Function

'Infer a format from a file extension; preserve the public enum values.
Private Function formatFromExtension(ByVal FileName As String) As stdImgFormat
  Select Case LCase$(Mid$(FileName, InStrRev(FileName, ".") + 1))
    Case "bmp": formatFromExtension = stdImgFormatBMP
    Case "png": formatFromExtension = stdImgFormatPNG
    Case "gif": formatFromExtension = stdImgFormatGIF
    Case "jpg", "jpeg": formatFromExtension = stdImgFormatJPEG
    Case "tif", "tiff": formatFromExtension = stdImgFormatTIFF
    Case "webp": Err.Raise 5, "stdImage", "WEBP format is not currently supported (see issue #178)"
    Case Else: Err.Raise 5, "stdImage", "Could not determine format from file extension"
  End Select
End Function

'Encode a bitmap to either a file or an independent memory stream.
Private Sub saveImage(ByVal format As stdImgFormat, ByVal FileName As String, ByVal stream As stdole.IUnknown, ByVal Quality As Long, ByVal TIFF_ColorDepth As Long, ByVal TIFF_Compression As Long)
  On Error GoTo Catch
  format = resolveOutputFormat(format)
  If This.hBitmap = 0 Then Err.Raise 5, "stdImage::saveImage", "Image is not initialised"
  If format = stdImgFormatJPEG And (Quality < 0 Or Quality > 100) Then Err.Raise 5, "stdImage::saveImage", "JPEG quality must be between 0 and 100"
  Dim startup As GdiplusStartupInput, token As LongPtr, image As LongPtr
  startup.GdiplusVersion = 1
  Call checkGdip(GdiplusStartup(token, startup), "start GDI+")
  Call checkGdip(GdipCreateBitmapFromHBITMAP(This.hBitmap, 0, image), "create encoder bitmap")
  Dim encoder As GUID, parameters As EncoderParameters, pParameters As LongPtr
  Dim encoderIDs As Variant
  encoderIDs = Array("", "{557CF400-1A04-11D3-9A73-0000F81EF32E}", "{557CF406-1A04-11D3-9A73-0000F81EF32E}", "{557CF402-1A04-11D3-9A73-0000F81EF32E}", "{557CF401-1A04-11D3-9A73-0000F81EF32E}", "{557CF405-1A04-11D3-9A73-0000F81EF32E}")
  Dim encoderID As String: encoderID = encoderIDs(format)
  If IIDFromString(StrPtr(encoderID), encoder) <> 0 Then Err.Raise 1, "stdImage::saveImage", "Invalid encoder identifier"
  Select Case format
    Case stdImgFormatJPEG
      parameters.count = 1
      With parameters.Parameter(0)
        Call IIDFromString(StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID)
        .NumberOfValues = 1: .type = 4: .Value = VarPtr(Quality)
      End With
    Case stdImgFormatTIFF
      parameters.count = 2
      With parameters.Parameter(0)
        Call IIDFromString(StrPtr("{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"), .GUID)
        .NumberOfValues = 1: .type = 4: .Value = VarPtr(TIFF_Compression)
      End With
      With parameters.Parameter(1)
        Call IIDFromString(StrPtr("{66087055-AD66-4C7C-9A18-38A2310B8337}"), .GUID)
        .NumberOfValues = 1: .type = 4: .Value = VarPtr(TIFF_ColorDepth)
      End With
  End Select
  If parameters.count > 0 Then pParameters = VarPtr(parameters)
  If stream Is Nothing Then
    Call checkGdip(GdipSaveImageToFile(image, StrPtr(FileName), encoder, ByVal pParameters), "save image to file")
  Else
    Call checkGdip(GdipSaveImageToStream(image, stream, encoder, pParameters), "save image to stream")
  End If
  Call GdipDisposeImage(image)
  Call GdiplusShutdown(token)
  Exit Sub
Catch:
  Dim errorNumber As Long, errorSource As String, errorDescription As String
  errorNumber = Err.Number: errorSource = Err.Source: errorDescription = Err.Description
  If image <> 0 Then Call GdipDisposeImage(image)
  If token <> 0 Then Call GdiplusShutdown(token)
  Err.Raise errorNumber, errorSource, errorDescription
End Sub

'Raise on a GDI+ failure so callers can share their cleanup path.
Private Sub checkGdip(ByVal status As Long, ByVal operation As String)
  If status <> 0 Then Err.Raise 1, "stdImage", "Could not " & operation & " (GDI+ status: " & status & ")"
End Sub

'Copy a borrowed bitmap into independently owned storage.
Private Function copyBitmap(ByVal hBitmap As LongPtr) As LongPtr
  If hBitmap = 0 Then Err.Raise 5, "stdImage", "Bitmap handle is empty"
  Const IMAGE_BITMAP As Long = 0, LR_CREATEDIBSECTION As Long = &H2000
  copyBitmap = CopyImage(hBitmap, IMAGE_BITMAP, 0, 0, LR_CREATEDIBSECTION)
  If copyBitmap = 0 Then Err.Raise 1, "stdImage", "Could not copy bitmap"
End Function

'Validate compatibility format hints; decoding itself uses image signatures.
Private Sub validateInputFormat(ByVal format As stdImgFormat)
  If format < stdImgFormatDefault Or format > stdImgFormatTIFF Then Err.Raise 5, "stdImage", "Unsupported input format"
End Sub

'Decode and rasterise the first frame, keeping the source stream alive until disposal.
Private Function loadImage(ByVal FileName As String, ByVal stream As stdole.IUnknown, ByVal background As Long) As stdImage
  On Error GoTo Catch
  Dim startup As GdiplusStartupInput, token As LongPtr
  Dim image As LongPtr, bitmap As LongPtr, graphics As LongPtr, hBitmap As LongPtr
  startup.GdiplusVersion = 1
  Call checkGdip(GdiplusStartup(token, startup), "start GDI+")
  If stream Is Nothing Then
    Call checkGdip(GdipLoadImageFromFile(StrPtr(FileName), image), "load image file")
  Else
    Call checkGdip(GdipLoadImageFromStream(stream, image), "load image stream")
  End If
  Dim width As Long, height As Long
  Call checkGdip(GdipGetImageWidth(image, width), "get image width")
  Call checkGdip(GdipGetImageHeight(image, height), "get image height")
  If width <= 0 Or height <= 0 Then Err.Raise 5, "stdImage", "Image dimensions are empty"
  Const PixelFormat32bppARGB As Long = &H26200A
  Call checkGdip(GdipCreateBitmapFromScan0(width, height, 0, PixelFormat32bppARGB, 0, bitmap), "create raster bitmap")
  Call checkGdip(GdipGetImageGraphicsContext(bitmap, graphics), "create raster graphics")
  Dim argbBackground As Long: argbBackground = &HFF000000 Or BGR(background)
  Call checkGdip(GdipGraphicsClear(graphics, argbBackground), "set image background")
  Call checkGdip(GdipDrawImageRectI(graphics, image, 0, 0, width, height), "rasterise image")
  Call GdipDeleteGraphics(graphics)
  graphics = 0
  Call checkGdip(GdipCreateHBITMAPFromBitmap(bitmap, hBitmap, argbBackground), "create bitmap handle")
  Set loadImage = CreateFromHBitmap(hBitmap)
  hBitmap = 0
  Call GdipDisposeImage(bitmap)
  Call GdipDisposeImage(image)
  Call GdiplusShutdown(token)
  Exit Function
Catch:
  Dim errorNumber As Long, errorSource As String, errorDescription As String
  errorNumber = Err.Number: errorSource = Err.Source: errorDescription = Err.Description
  If graphics <> 0 Then Call GdipDeleteGraphics(graphics)
  If hBitmap <> 0 Then Call DeleteObject(hBitmap)
  If bitmap <> 0 Then Call GdipDisposeImage(bitmap)
  If image <> 0 Then Call GdipDisposeImage(image)
  If token <> 0 Then Call GdiplusShutdown(token)
  Err.Raise errorNumber, errorSource, errorDescription
End Function

'Wrap a copied bitmap, including cleanup if initialisation fails.
Private Function imageFromBitmapCopy(ByVal hBitmap As LongPtr) As stdImage
  On Error GoTo Catch
  Dim copy As LongPtr: copy = copyBitmap(hBitmap)
  Set imageFromBitmapCopy = CreateFromHBitmap(copy)
  Exit Function
Catch:
  Dim errorNumber As Long, errorSource As String, errorDescription As String
  errorNumber = Err.Number: errorSource = Err.Source: errorDescription = Err.Description
  If copy <> 0 Then Call DeleteObject(copy)
  Err.Raise errorNumber, errorSource, errorDescription
End Function
