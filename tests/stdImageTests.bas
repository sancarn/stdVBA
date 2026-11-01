Attribute VB_Name = "stdImageTests"
Option Explicit

#If VBA7 Then
  Private Declare PtrSafe Function GetCurrentProcess Lib "kernel32" () As LongPtr
  Private Declare PtrSafe Function GetGuiResources Lib "user32" (ByVal process As LongPtr, ByVal flags As Long) As Long
  Private Declare PtrSafe Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal handle As LongPtr, ByVal count As Long, ByRef buffer As Any) As Long
#Else
  Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
  Private Declare Function GetGuiResources Lib "user32" (ByVal process As Long, ByVal flags As Long) As Long
  Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal handle As Long, ByVal count As Long, ByRef buffer As Any) As Long
#End If
Private checks As Long

'Run using Tools/Testing/run.py --suite stdImage, which supplies disposable fixtures and hosts.
Public Function RunTests(ByVal folder As String, ByVal baseURL As String) As String
  On Error GoTo Failed
  checks = 0
  Dim img As stdImage, copy As stdImage, data() As Byte, shifted() As Byte
  Set img = stdImage.CreateFromFile(folder & "\fixture.png")
  Check img.width = 3 And img.height = 2, "PNG dimensions"
  CheckPixels img, "PNG input"
  Dim format As Long, url As String, path As String, i As Long
  Dim mimeTypes As Variant: mimeTypes = Array("", "image/bmp", "image/png", "image/gif", "image/jpeg", "image/tiff")
  For format = 1 To 5
    data = img.ToBinary(format)
    Check signature(data, format), "Binary signature " & format
    Set copy = stdImage.CreateFromBinary(data)
    Check copy.width = 3 And copy.height = 2, "Binary dimensions " & format
    If format <> 4 Then CheckPixels copy, "Binary pixels " & format
    url = img.ToDataUrl(format)
    Check Left$(url, Len("data:" & mimeTypes(format) & ";base64,")) = "data:" & mimeTypes(format) & ";base64,", "MIME type " & format
    Check InStr(url, vbCr) = 0 And InStr(url, vbLf) = 0 And InStr(url, vbNullChar) = 0, "Base64 has no CR/LF/null"
    Set copy = stdImage.CreateFromDataURL(url)
    Check copy.width = 3 And copy.height = 2, "Data URL dimensions " & format
    If format <> 4 Then CheckPixels copy, "Data URL pixels " & format
    path = img.ToTempFile(format)
    Check Len(Dir$(path)) > 0, "Explicit temporary file"
    Set copy = stdImage.CreateFromFile(path)
    Check copy.width = 3, "Temporary file loads"
    Kill path
  Next
  data = img.ToBinary()
  Check signature(data, 2), "Default binary PNG"
  Check Left$(img.ToDataUrl(), 22) = "data:image/png;base64,", "Default Data URL PNG"
  path = img.ToTempFile()
  Check Right$(path, 4) = ".png", "Default temporary extension"
  Kill path

  'No file-handle assumption: reserve handle #1 while serialising.
  Dim fileNumber As Long: fileNumber = FreeFile
  Open folder & "\open-handle.bin" For Binary As #fileNumber
  data = img.ToBinary(stdImgFormatBMP)
  Close #fileNumber
  Check signature(data, 1), "Export with another file open"
  ReDim shifted(7 To 7 + UBound(data))
  For i = 0 To UBound(data): shifted(i + 7) = data(i): Next
  Set copy = stdImage.CreateFromBinary(shifted)
  CheckPixels copy, "Nonzero lower bound"

  Dim extension As Variant
  For Each extension In Array("bmp", "gif", "jpg", "tif", "wmf", "emf", "ico")
    Set copy = stdImage.CreateFromFile(folder & "\fixture." & extension)
    Check copy.width > 0 And copy.height > 0, "Input extension " & extension
  Next
  Set copy = stdImage.CreateFromFile(folder & "\transparent.png", transparentColor:=vbBlue)
  Dim pixels() As Long: pixels = copy.Colors
  Check CBool((pixels(1, 2) And &HFFFFFF) = &HFF), "Transparent pixel flattened against blue"
  Set copy = stdImage.CreateFromFile(folder & "\transparent.png")
  pixels = copy.Colors
  Check CBool((pixels(1, 2) And &HFFFFFF) = 0), "Default transparency background black"

  Set copy = stdImage.CreateFromFile(folder & "\quality.png")
  copy.ToFile folder & "\low.jpg", Quality:=5
  copy.ToFile folder & "\high.jpg", Quality:=95
  Check FileLen(folder & "\high.jpg") > FileLen(folder & "\low.jpg"), "JPEG quality parameters"
  img.ToFile folder & "\output.tif", TIFF_Compression:=2
  Set copy = stdImage.CreateFromFile(folder & "\output.tif")
  CheckPixels copy, "TIFF encoder parameters"

  'Pictures and images must remain valid when their source wrapper is destroyed.
  Dim picture As stdole.StdPicture, disp As stdole.IPictureDisp, ip As stdole.IPicture
  Set picture = img.ToStdPicture()
  Set img = Nothing
  Set img = stdImage.CreateFromStdPicture(picture)
  Set disp = picture
  Set copy = stdImage.CreateFromIPictureDisp(disp)
  Set copy = Nothing
  Set ip = picture
  Set copy = stdImage.CreateFromIPicture(ip)
  Set ip = Nothing: Set disp = Nothing: Set picture = Nothing
  CheckPixels copy, "IPicture input owns copy"
  CheckPixels img, "StdPicture input owns copy"
  Set picture = LoadPicture(folder & "\fixture.bmp")
  Set copy = stdImage.CreateFromStdPicture(picture)
  Set copy = Nothing
  Dim info(0 To 7) As Long
  Check GetObjectAPI(picture.handle, 32, info(0)) <> 0, "Borrowed picture not deleted"
  Set picture = LoadPicture(folder & "\fixture.ico")
  Set copy = stdImage.CreateFromStdPicture(picture)
  Set picture = Nothing
  Check copy.width = 16 And copy.height = 16, "Icon picture supported"

  img.ToClipboard
  Set copy = stdImage.CreateFromClipboard()
  Set copy = Nothing
  Set copy = stdImage.CreateFromClipboard()
  CheckPixels copy, "Clipboard survives input disposal"
  Set img = Nothing
  Set img = stdImage.CreateFromClipboard()
  CheckPixels img, "Clipboard survives source disposal"
  Set copy = stdImage.CreateFromFile(folder & "\quality.png")
  copy.ToClipboard
  CheckPixels img, "Image survives clipboard replacement"
  Set copy = Nothing

  Set copy = stdImage.CreateFromUrl(baseURL & "/image")
  CheckPixels copy, "URL input"

  Dim invalid As Variant
  For Each invalid In Array("data:image/png;base64,", "data:image/png;base64,====", "http:image/png;base64,AAAA", "data:text/plain;base64,AAAA", "data:image/png,abc", "data:image/png;base64,AA!A", "data:image/webp;base64,AAAA", "data:image/png;base64")
    ExpectFailure 1, CStr(invalid), folder, baseURL
  Next
  For i = 2 To 10: ExpectFailure i, "", folder, baseURL: Next

  'Stress successful BMP saves and invalid input/output cleanup after warming up GDI+.
  Set copy = img
  copy.ToFile folder & "\stress.bmp"
  Dim before As Long: before = GetGuiResources(GetCurrentProcess(), 0)
  For i = 1 To 40
    data = img.ToBinary(stdImgFormatBMP)
    Set copy = stdImage.CreateFromBinary(data)
    Set picture = copy.ToStdPicture()
    Set picture = Nothing: Set copy = Nothing
    img.ToFile folder & "\stress.bmp"
    ExpectFailure 3, "", folder, baseURL
    ExpectFailure 5, "", folder, baseURL
  Next
  Dim after As Long: after = GetGuiResources(GetCurrentProcess(), 0)
  Check after <= before + 3, "GDI handles stable: " & before & " -> " & after

  If Application.Name = "Microsoft Excel" Then
    Dim app As Object: Set app = Application
    Dim sheet As Object: Set sheet = app.ActiveWorkbook.Worksheets(1)
    Dim shape As Object: Set shape = img.ToSheet(sheet, Top:=20, Left:=40)
    Check Abs(shape.Top - 20) < 1 And Abs(shape.Left - 40) < 1, "ToSheet position"
    Set copy = stdImage.CreateFromExcelIPicture(sheet.Pictures(1))
    Check copy.width > 0, "Excel picture example"
    img.ToShapeFill shape
    shape.Delete
  End If
  RunTests = "PASS: " & checks & " stdImage assertions"
  Exit Function
Failed:
  RunTests = "FAIL after " & checks & ": " & Err.Number & " " & Err.Source & " " & Err.Description
End Function

Private Sub Check(ByVal condition As Boolean, ByVal label As String)
  checks = checks + 1
  If Not condition Then Err.Raise vbObjectError + 98, "stdImageTests", label
End Sub

Private Sub CheckPixels(ByVal image As stdImage, ByVal label As String)
  Dim pixels() As Long: pixels = image.Colors
  Check image.width = 3 And image.height = 2, label & " dimensions"
  Check CBool((pixels(1, 2) And &HFFFFFF) = &HFF0000), label & " red"
  Check CBool((pixels(2, 2) And &HFFFFFF) = &HFF00&), label & " green"
  Check CBool((pixels(3, 2) And &HFFFFFF) = &HFF), label & " blue"
  Check CBool((pixels(1, 1) And &HFFFFFF) = &HFFFFFF), label & " white"
End Sub

Private Function signature(ByRef data() As Byte, ByVal format As Long) As Boolean
  Select Case format
    Case 1: signature = data(0) = 66 And data(1) = 77
    Case 2: signature = data(0) = 137 And data(1) = 80 And data(2) = 78 And data(3) = 71
    Case 3: signature = data(0) = 71 And data(1) = 73 And data(2) = 70
    Case 4: signature = data(0) = 255 And data(1) = 216
    Case 5: signature = (data(0) = 73 And data(1) = 73) Or (data(0) = 77 And data(1) = 77)
  End Select
End Function

Private Sub ExpectFailure(ByVal kind As Long, ByVal text As String, ByVal folder As String, ByVal baseURL As String)
  Dim img As stdImage, data() As Byte, caught As Long
  On Error Resume Next
  Select Case kind
    Case 1: Set img = stdImage.CreateFromDataURL(text)
    Case 2: Set img = stdImage.CreateFromBinary(data)
    Case 3
      ReDim data(0 To 3): Set img = stdImage.CreateFromBinary(data)
    Case 4: Set img = stdImage.CreateFromFile(folder & "\missing.png")
    Case 5
      Set img = stdImage.CreateFromFile(folder & "\fixture.png")
      img.ToFile folder & "\missing\out.bmp"
    Case 6
      Set img = stdImage.CreateFromFile(folder & "\fixture.png")
      data = img.ToBinary(99)
    Case 7
      Set img = stdImage.CreateFromFile(folder & "\fixture.png")
      img.ToFile folder & "\bad.jpg", Quality:=101
    Case 8: Set img = stdImage.CreateFromUrl(baseURL & "/missing")
    Case 9: Set img = stdImage.CreateFromFile(folder & "\unsupported.webp")
    Case 10
      Set img = New stdImage
      data = img.ToBinary()
  End Select
  caught = Err.Number
  Err.Clear
  On Error GoTo 0
  Check caught <> 0, "Expected failure case " & kind & " " & text
End Sub
