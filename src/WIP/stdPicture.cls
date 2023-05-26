VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdPicture"
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
' * TODO: Convert to GDIP instead of WIA e.g: https://www.mrexcel.com/board/threads/using-listbox-value-to-display-png-in-image-box.1103057/
'   GDIP is faster. WIA is apparently intentionally slow...






'Mostly Wraps GDI+ and WIA APIs
'Also allows Copy(), Export(), Picture (which can be used for PictureBox or Image control), PictureDisp (can be used with charts)

'Scope:
'stdPicture::CreateFromFile(".../test.png")
'stdPicture::CreateFromFile(".../test.jpg")
'stdPicture::CreateFromFile(".../test.tif")
'stdPicture::CreateFromFile(...)
'stdPicture::CreateFromShape(shp as Shape)
'stdPicture::Create()                                                                                             'Create blank canvas - can be used for charts
'stdPicture::CreateFromPicture(...)                                                                               'Create from StdOle.StdPicture
'stdPicture::CreateFromPictureDisp(...)                                                                           'Create from StdOle.StdOle.IPictureDisp
'stdPicture::CreateFromClipboard()                                                                                'Create image from clipboard
'<#stdPicture>.CreateFromSVG()                                                                                     ???
'<#stdPicture>.ToClipboard()                                                                                      'Copies to clipboard
'<#stdPicture>.ToFile(sFilePath as string)                                                                        'Exports to file
'<#stdPicture>.ToUIComponent()                                                                                    'Converts to UserForm Component
'<#stdPicture>.ToDataURL()                                                                                        'Converts picture to data url as string
'<#stdPicture>.Picture                                                                                            'Returns StdOle.StdPicture
'<#stdPicture>.PictureDisp                                                                                        'Returns StdOle.IPictureDisp
'<#stdPicture>.Draw(x as long, y as long, rgbColor as long)                                                       'Draw pixel
'<#stdPicture>.DrawLine(x1 as long, y1 as long, x2 as long, y2 as long, rgbColor as long)                         'Draw Line
'<#stdPicture>.DrawEllipse(x as long, y as long, w as long, h as long, rgbColor as long)                          'Draw Ellipse
'<#stdPicture>.DrawPolyline(points() as Point, rgbLineColor as long, iLineThickness as long)                      'Draw a polyline
'<#stdPicture>.DrawPolygon(points() as Point, rgbLineColor as long, iLineThickness as long, rgbFillColor as long) 'Draw a polygon
'<#stdPicture>.DrawRect(x,y,w,h)                                                                                  'Draw a rectangle
'<#stdPicture>.DrawRoundRect(x,y,w,h,rw,rh)                                                                       'Draw a rectangle with rounded corners
'<#stdPicture>.DrawArc(...)                                                                                       'Draw an arc
'<#stdPicture>.DrawChord(...)                                                                                     'Draw a chord
'<#stdPicture>.DrawPie(...)                                                                                       'Draw a pie
'<#stdPicture>.DrawPolyBezier(...)                                                                                'Draws cubic Bézier curves by using the endpoints and control points specified by the lppt parameter.
'<#stdPicture>.Resize()
'<#stdPicture>.Crop()

Public Enum stdPicFormat
  stdPicFormatDefault
  stdPicFormatBMP
  stdPicFormatPNG
  stdPicFormatGIF
  stdPicFormatJPEG
  stdPicFormatTIFF
End Enum

#If Win64 Then
  Private Const NULL_PTR = 0^
#Else
  Private Const NULL_PTR = 0&
#End If

'Clipboard Management
#If VBA7 Then
  Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Integer) As Long
  Private Declare PtrSafe Function OpenClipboard Lib "user32" (Optional ByVal hWnd As LongPtr = 0^) As Long
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

  Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As Long)
  Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long

  Private Declare PtrSafe Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As LongPtr, ByVal offsetinVft As LongPtr, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As VbVarType, ByRef paValues As LongPtr, ByRef retVAR As Variant) As Long
#Else
  Enum LongPtr
    [_]
  End Enum

  Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Integer) As Long
  Private Declare Function OpenClipboard Lib "user32" (Optional ByVal hwnd As LongPtr = 0^) As Long
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
  
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As Long)
  Private Declare Function GetTickCount Lib "kernel32" () As Long

  'FIX: The use of `VbVarType` for the type of `paTypes` on 32 bit causes Invalid Callee error. Workaround is to use `Integer` instead as below.
  Private Declare Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
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



'****************
'* Private Vars *
'****************
Private pImageFile As Object
Private pCurrentFormat As stdPicFormat

'****************
'* Constructors *
'****************

'Creates a stdPicture object from a byte array
'@param {ByRef Array<Byte>} Bytes of Bitmap/Jpeg/Png/... image
'@returns {stdPicture} Picture object representing data
'@example ```
'  Dim b() as Byte: Redim b(1 to 100)
'  Call stdPicture.CreateFromBinary(b).saveAs("H:\test1.png", stdPicFormatPNG)
'```
Public Function CreateFromBinary(ByRef b() As Byte) As stdPicture
  Set CreateFromBinary = New stdPicture
  With CreateObject("WIA.Vector")
    .BinaryData = b
    Call CreateFromBinary.protInit(.imageFile)
  End With
End Function


'Creates a stdPicture object from a file
'@param {String} Path of file
'@returns {stdPicture} Picture object representing data
'@example stdPicture.CreateFromFile("H:\test1.bmp").saveAs("H:\test1.png", stdPicFormatPNG)
Public Function CreateFromFile(ByVal sPath As String) As stdPicture
  Set CreateFromFile = New stdPicture
  Set wia = CreateObject("WIA.ImageFile")
  Call wia.LoadFile(sPath)
  Call CreateFromFile.protInit(wia)
End Function

'Creates a stdPicture object from a url
'@param {String} Url of file
'@returns {stdPicture} Picture object representing data
'@example stdPicture.CreateFromURL("https://raw.githubusercontent.com/sancarn/awesome-vba/main/resources/VBALogo.png").saveAs("H:\test.png", stdPicFormatPNG)
Public Function CreateFromUrl(ByVal sURL As String) As stdPicture
  Set CreateFromUrl = New stdPicture
  Dim HTTP As Object: Set HTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
  HTTP.Open "GET", TargetURL, False
  'TODO: AutoProxy?
  HTTP.send
  If HTTP.status = 200 Then
    Dim body() As Byte: body = HTTP.responseBody
    Set CreateFromUrl = CreateFromBinary(body)
  Else
    Err.Raise 1, "", "Could not get file (HTTP Status: " & HTTP.status & ")"
  End If
End Function

'Creates a stdPicture object from a file
'@param {LongPtr} hBitmap (handle to Bitmap)
'@returns {stdPicture} Picture object representing data
'@example ```
'  Dim bmp as LongPtr: bmp = TODO:
'  stdPicture.CreateFromHBitmap(bmp).saveAs("H:\test.png", stdPicFormatPNG)
'```
Public Function CreateFromHBitmap(ByVal hBitmap As LongPtr) As stdPicture
  Dim b() As Byte: b = getHBitmapBytes(hBitmap) ' Obtain byte array
  Set CreateFromHBitmap = CreateFromBinary(b) 'Initialise from binary
End Function

'Creates a stdPicture from the image in the clipboard
'@returns {stdPicture} Picture object representing data
'@example stdPicture.CreateFromClipboard().saveAs("H:\test.png", stdPicFormatPNG)
Public Function CreateFromClipboard() As stdPicture
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

'Creates a stdPicture object from a stdole.IPictureDisp object
'@param {stdole.IPictureDisp} Object to create picture from
'@returns {stdPicture} Picture object representing data
'@example TODO:
Public Function CreateFromSTDOLEIPictureDisp(ByVal picture As stdole.IPictureDisp) As stdPicture
  Set CreateFromSTDOLEIPictureDisp = CreateFromHBitmap(picture.Handle)
End Function

'Creates a stdPicture object from a stdole.stdPicture object
'@param {stdole.stdPicture} Object to create picture from
'@returns {stdPicture} Picture object representing data
'@example stdPicture.CreateFromSTDOLEStdPicture(Userform1.Picture).saveAs("H:\test.png", stdPicFormatPNG)
Public Function CreateFromSTDOLEStdPicture(ByVal picture As stdole.stdPicture) As stdPicture
  Set CreateFromSTDOLEStdPicture = CreateFromHBitmap(picture.Handle)
End Function

'Creates a stdPicture object from a stdole.IPicture object
'@param {stdole.IPicture} Object to create picture from
'@returns {stdPicture} Picture object representing data
'@example TODO:
Public Function CreateFromSTDOLEIPicture(ByVal picture As stdole.IPicture) As stdPicture
  Set CreateFromSTDOLEIPicture = CreateFromHBitmap(picture.Handle)
End Function

'Creates a stdPicture object from a Excel.IPicture object
'@param {Excel.IPicture} Object to create picture from
'@returns {stdPicture} Picture object representing data
'@example TODO:
Public Function CreateFromExcelIPicture(ByVal picture As Excel.IPicture) As stdPicture
  Call picture.CopyPicture(, xlBitmap)
  Set CreateFromExcelIPicture = CreateFromClipboard()
End Function

'Creates a stdPicture object from a Shape object
'@param {Object<(Excel|Word|PowerPoint).Shape>} Object to create picture from
'@returns {stdPicture} Picture object representing data
'@example TODO:
Public Function CreateFromShape(ByVal shp As Object) As stdPicture
  Select Case Application.Name
    Case "Microsoft Excel"
      Dim xlShp As Excel.Shape: Set xlShp = shp
      Call xlShp.CopyPicture(, XlCopyPictureFormat.xlBitmap)
      Set CreateFromShape = stdPicture.CreateFromClipboard()
  End Select
End Function

'Creates a stdPicture object from a Chart object
'@param {Object<(Excel|Word|PowerPoint).Chart>} Object to create picture from
'@returns {stdPicture} Picture object representing data
Public Function CreateFromChart(ByVal chart As Object) As stdPicture
  Select Case Application.Name
    Case "Microsoft Excel"
      Dim xlChart As Excel.chart: Set xlChart = chart
      Call xlChart.CopyPicture(, XlCopyPictureFormat.xlBitmap)
      Set CreateFromChart = stdPicture.CreateFromClipboard()
  End Select
End Function

'Initialised stdPicture from WIA object
'@param {Object<WIA.ImageFile>}
Public Sub protInit(ByVal imageFile As Object)
  Set pImageFile = imageFile
End Sub

'Obtain binary data of the image in any format required
'@param {eFormat=Default} The format the data should be in.
'@returns {Byte()} Binary data of image
Public Property Get Binary(Optional ByVal eFormat As stdPicFormat = stdPicFormatDefault) As Byte()
  'If default, leave in current format
  If eFormat = stdPicFormatDefault Then
    Binary = pImageFile.FileData.BinaryData
  Else
    'Obtain required format
    Dim sFormatUID As String
    Select Case eFormat
      Case stdPicFormatBMP:  sFormatUID = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
      Case stdPicFormatPNG:  sFormatUID = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
      Case stdPicFormatGIF:  sFormatUID = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}"
      Case stdPicFormatJPEG: sFormatUID = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
      Case stdPicFormatTIFF: sFormatUID = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}"
    End Select

    With CreateObject("WIA.ImageProcess")
      .Filters.Add .FilterInfos("Convert").FilterID
      .Filters(.Filters.Count).Properties("FormatID").Value = sFormatUID
      Binary = .Apply(pImageFile).FileData.BinaryData
    End With
  End If
End Property


'Set binary data of the image in any format required
'@param {eFormat} The format the data should be in.
'@param {ByRef Byte()} Binary data of image
'@planned TODO:
Public Property Let Binary(ByVal eFormat As stdPicFormat, ByRef b() As Byte)
  With CreateObject("WIA.Vector")
    .BinaryData = b
    Call protInit(.imageFile)
  End With
End Property

'Save image as file
'@param {String} Path to new file
'@param {eFormat=Default} The required format of the file. Default is the current assosciated format.
Public Sub saveAs(ByVal sSaveAsFile As String, Optional ByVal eFormat As stdPicFormat = stdPicFormatDefault)
  Dim ff As Long: ff = FreeFile()
  Open sSaveAsFile For Binary As #ff
    Put #ff, 1, Binary(eFormat)
  Close #ff
End Sub

'
'@planned TODO:
Public Sub toClipboard()
  Err.Raise 1, , "Not implemented"
End Sub

'Save a copy in a temporary file and return the path of this file
'@param {eFormat} The format the data should be in.
'@returns {String} The path to the temporary file.
Public Function toTempFile(Optional ByVal eFormat As stdPicFormat = stdPicFormatDefault) as string
  Dim sExt As String
  Select Case eFormat
    Case stdPicFormatBMP:  sExt = "bmp"
    Case stdPicFormatPNG:  sExt = "png"
    Case stdPicFormatGIF:  sExt = "gif"
    Case stdPicFormatJPEG: sExt = "jpeg"
    Case stdPicFormatTIFF: sExt = "tiff"
  End Select
  Dim sPath as string: sPath = Environ("Temp") & "\" & getGUID() & "." & sExt
  Call saveAs(sPath, eFormat)
  toTempFile = sPath
End Function

'Adds the picture to the sheet
'@param {Worksheet} Worksheet to add image to
'@param {Double=0} Top coord of the newly placed image
'@param {Double=0} Left coord of the newly placed image
'@param {MsoPictureCompress} Whether to compress the image on import.
'@returns {Object<Picture>} Picture object
Public Function toSheet(ByVal ws as Worksheet, Optional ByVal Top as Double = 0 , Optional ByVal Left as Double = 0, optional ByVal compress as MsoPictureCompress) as Object
  set toSheet = ws.Shapes.AddPicture2(toTempFile(), false, 0, Top, Left, -1, -1, compress)
End Function

'Update a UI control's background picture
'@param {Object<Control>}
Public Sub toUIControl(Byval oControl as object)
  set oControl.Picture = STDOLE_StdPicture
End Sub

'@planned TODO:
Public Function toDataUrl(Optional ByVal eFormat As stdPicFormat = stdPicFormatDefault) as string
  Err.raise 1, , "Not implemented"
End Function

Public Sub toShapeFill(ByVal shp As Object)
  Select Case Application.Name
    Case "Microsoft Excel"
      Dim xlShp As Excel.Shape: Set xlShp = shp
      Call xlShp.Fill.UserPicture(toTempFile())
  End Select
End Sub

'Obtain a stdole.StdPicture
'@returns {stdole.StdPicture} Obtain the image as a stdole.StdPicture
Public Property Get STDOLE_StdPicture() as stdole.StdPicture
  set toSTDOLEStdPicture = pImageFile.FileData.Picture
End Property

'Obtain a HBitmap
'@returns {stdole.StdPicture} Obtain the image as a stdole.StdPicture
'@remark You are in charge of freeing this pointer.
Public Property Get HBitmap() as LongPtr
  HBitmap = STDOLE_StdPicture.Handle
End Property



















' Option Explicit

' Private Type GUID
'     Data1 As Long
'     Data2 As Integer
'     Data3 As Integer
'     Data4(0 To 7) As Byte
' End Type
' Private Type GdiplusStartupInput
'     GdiplusVersion As Long
'     DebugEventCallback As Long
'     SuppressBackgroundThread As Long
'     SuppressExternalCodecs As Long
' End Type
' Private Type EncoderParameter
'     GUID As GUID
'     NumberOfValues As Long
'     type As Long
'     Value As Long
' End Type
' Private Type EncoderParameters
'     count As Long
'     Parameter As EncoderParameter
' End Type


' Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
' Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long
' Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As Long, ByVal hPal As Long, BITMAP As Long) As Long
' Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As Long
' Private Declare Function GdipSaveImageToFile Lib "GDIPlus" (ByVal Image As Long, ByVal FileName As Long, clsidEncoder As GUID, encoderParams As Any) As Long
' Private Declare Function CLSIDFromString Lib "ole32" (ByVal Str As Long, id As GUID) As Long
' Private Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb As Long) As Long

' Private pPicture as StdOle.StdPicture

' Public Event PictureBeforeChanged(OldPicture as StdOle.StdPicture, NewPicture as StdOle.StdPicture)
' Public Event PictureAfterChanged(NewPicture as StdOle.StdPicture)

' Public Property Get Picture() as StdOle.StdPicture
'   set Picture = pPicture
' End Sub
' Public Property Set Picture(pic as StdOle.StdPicture)
'   RaiseEvent PictureBeforeChanged(pPicture,pic)
'   set pPicture = pic
'   RaiseEvent PictureAfterChanged(pPicture)
' End Sub

' Public Sub ToClipboard()
'   if vartype(Clipboard) = vbObject then
'     Clipboard.SetData pPicture
'   else
'     'TODO: Think this method is very hacky.
'     'TODO: A less hacky method can be found here:  http://access.mvps.org/access/api/api0042.htm
'     With ActiveSheet.OLEObjects.Add(ClassType:="Forms.Image.1", Link:=False, DisplayAsIcon:=False, Left:=330, Top:=215)
'       .Activate
'       .Picture = pPicture
'       .AutoSize = true
'       .CopyPicture
'       .Delete
'     End with
'   end if
' End Sub

' Public Sub ToFile(ByVal FileName As String, Optional PicType As String="INFER", Optional ByVal Quality As Byte = 80, Optional ByVal TIFF_ColorDepth As Long = 24, Optional ByVal TIFF_Compression As Long = 6)
'     Screen.MousePointer = vbHourglass
'     Dim tSI As GdiplusStartupInput
'     Dim lRes As Long
'     Dim lGDIP As Long
'     Dim lBitmap As Long
'     Dim aEncParams() As Byte
'     On Error GoTo ErrHandle:
'     tSI.GdiplusVersion = 1
'     lRes = GdiplusStartup(lGDIP, tSI)
'     If lRes = 0 Then
'         lRes = GdipCreateBitmapFromHBITMAP(pPicture.Handle, 0, lBitmap)
'         If lRes = 0 Then
'             Dim tJpgEncoder As GUID
'             Dim tParams As EncoderParameters
'             Select Case PicType
'             Case ".jpg"
'                 CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
'                 tParams.count = 1
'                 With tParams.Parameter
'                     CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID
'                     .NumberOfValues = 1
'                     .type = 4
'                     .Value = VarPtr(Quality)
'                 End With
'                 ReDim aEncParams(1 To Len(tParams))
'                 Call CopyMemory(aEncParams(1), tParams, Len(tParams))
'             Case ".png"
'                 CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
'                 ReDim aEncParams(1 To Len(tParams))
'             Case ".gif"
'                 CLSIDFromString StrPtr("{557CF402-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
'                 ReDim aEncParams(1 To Len(tParams))
'             Case ".tiff"
'                 CLSIDFromString StrPtr("{557CF405-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
'                 tParams.count = 2
'                 ReDim aEncParams(1 To Len(tParams) + Len(tParams.Parameter))
'                 With tParams.Parameter
'                     .NumberOfValues = 1
'                     .type = 4
'                     CLSIDFromString StrPtr("{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"), .GUID
'                     .Value = VarPtr(TIFF_Compression)
'                 End With
'                 Call CopyMemory(aEncParams(1), tParams, Len(tParams))
'                 With tParams.Parameter
'                     .NumberOfValues = 1
'                     .type = 4
'                     CLSIDFromString StrPtr("{66087055-AD66-4C7C-9A18-38A2310B8337}"), .GUID
'                     .Value = VarPtr(TIFF_ColorDepth)
'                 End With
'                 Call CopyMemory(aEncParams(Len(tParams) + 1), tParams.Parameter, Len(tParams.Parameter))
'             Case ".bmp"
'                 SavePicture pPicture, FileName
'                 Screen.MousePointer = vbDefault
'                 Exit Sub
'             End Select
'             lRes = GdipSaveImageToFile(lBitmap, StrPtr(FileName), tJpgEncoder, aEncParams(1))
'             GdipDisposeImage lBitmap
'         End If
'         GdiplusShutdown lGDIP
'     End If
'     Screen.MousePointer = vbDefault
'     Erase aEncParams
'     Exit Sub
'   ErrHandle:
'     Screen.MousePointer = vbDefault
'     MsgBox "Error" & vbCrLf & vbCrLf & "Error No. " & Err.Number & vbCrLf & " Error .Description:  " & Err.Description, vbInformation Or vbOKOnly
' End Sub

'@param {IUnknown} Object to query the interface of
'@param {string} IID to query
'@returns {IUnknown} New interface
Private Function unkQueryInterface(ByVal obj As IUnknown, ByVal sIID As String) As IUnknown
    Dim GUID() As Byte: GUID = IIDFromString(sIID)
    Call CallVT(ObjPtr(obj), 0, vbLong, VarPtr(GUID(0)), VarPtr(unkQueryInterface))
End Function

'Obtain a UUID as bytes from a string
'@param {string} UUID, may contain {} or -s
'@return {Byte()} Bytes representing the UUID
Private Function IIDFromString(ByVal sUUID As String) As Byte()
    Dim ret() As Byte: ReDim ret(0 To 15)
    Dim i As Long: i = 1
    Dim j As Long: j = 0
    While i < Len(sUUID)
        Select Case Mid(sUUID, i, 1)
            Case "{", "-", "}"
                i = i + 1
            Case Else
                ret(j) = CLng("&H" & Mid(sUUID, i, 2))
                j = j + 1
                i = i + 2
        End Select
    Wend
    IIDFromString = ret
End Function

'Call a function in the VTable of the object at a specified offset
'@param {Long}      The offset of the method to call. IUnknown::QueryInterface=0; IUnknown::AddRef=1; IUnknown::Release=2; IDispatch::Invoke=6
'@param {VbVarType} Return value of function, typically Long (HResult) for COM Object calls
'@paramArray Arguments to send to function.
Public Function CallVT(ByVal pObjPtr As LongPtr, ByVal VTableOffset As Long, ByVal ReturnType As VbVarType, ParamArray FunctionParameters() As Variant) As Variant
    'Create array from params
    Dim vParam() As Variant: vParam = FunctionParameters
    Dim iParamCount As Long: iParamCount = UBound(vParam) - LBound(vParam) + 1
    
    'Create array of types and pointers to vars, for function call
    #If Win64 Then
      Dim vPtr() As LongLong
      Dim vType() As VbVarType
    #Else
      Dim vPtr() As Long
      Dim vType() As Integer 'FIX: On 32-bit use of vbVarType for the type parameter will fail
    #End If
    
    'Populate array of pointers to params, and array of types for DispCallFunc
    If iParamCount = 0 Then
      ReDim vType(0 To 0)
      ReDim vPtr(0 To 0)
    Else
      ReDim vType(0 To iParamCount - 1)
      ReDim vPtr(0 To iParamCount - 1)
    
      Dim lIdx As Long
      For lIdx = 0 To iParamCount - 1
        vType(lIdx) = VarType(vParam(lIdx))
        vPtr(lIdx) = VarPtr(vParam(lIdx))
      Next
    End If
    
    'Call function
    Dim hResult As Long
    hResult = DispCallFunc(pObjPtr, VTableOffset * PTRSIZE, CC_STDCALL, ReturnType, iParamCount, vType(0), vPtr(0), vv)
    
    If hResult < 0 Then
        Err.Raise hResult, "DispCallFunc"
    End If
End Function




'@source: https://www.mrexcel.com/board/threads/working-with-images-in-vba-displaying-png-files.1224206/#post-5991836
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
'@param {ByVal LongPtr/Long} The pointer to copy bytes at
'@param {ByVal Long} The size of the buffer to copy
'@returns {Byte()} The bytes stired at the pointer
#If VBA7 Then
Private Function BytesFromPointer(ByVal pointer As LongPtr, ByVal size As Long) As Byte()
#Else
Private Function BytesFromPointer(ByVal pointer As Long, ByVal size As Long) As Byte()
#End If
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
'@param {long} hWnd to open clipboard in
'@param {long} number of milliseconds to wait
'@returns {boolean} true if clipboard is open, false if clipboard couldn't be opened
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
Private Function getGUID() As String
  Call Randomize 'Ensure random GUID generated
  getGUID = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx"
  getGUID = Replace(getGUID, "y", Hex(Rnd() And &H3 Or &H8))
  Dim i As Long: For i = 1 To 30
    getGUID = Replace(getGUID, "x", Hex$(Int(Rnd() * 16)), 1, 1)
  Next
End Function