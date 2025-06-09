VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdClipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'CREDITS:
'----------
'Initial pictures implementation by Stephen Bullen (Office Automation Ltd) on 30/10/1998
'Other clipboard API examples and docs which have helped this project:
'* https://docs.microsoft.com/en-us/windows/win32/dataxchg/clipboard-overviews
'* https://docs.microsoft.com/en-us/windows/win32/dataxchg/clipboard-functions
'* https://codekabinett.com/rdumps.php?Lang=2&targetDoc=vba-clipboard-file-content
'* https://stackoverflow.com/a/35512118/6302131
'----------

'TODO: File is not fully complete, even though it has reached a stable state. Remaining tasks:
'* This needs a new home
  ''Save a picture from a shape
  ''@param {ByVal Shape} - Shape to copy as picture
  ''@param {ByVal String} - File path to save shape as picture too
  'Private Sub VbSavePicture(ByVal shp As Shape, ByVal file As String)
  '  Call shp.CopyPicture(Appearance:=xlScreen, format:=xlBitmap)
  '  StdOle.SavePicture GetPictureFromClipboard(xlBitmap), file
  'End Sub
'* Implement Picture [SET]

'TODO: Elsewhere not here:
'IPictureFromFile()
'IPictureFromHBitmap()
'IPictureFromBlob()


Public Enum CLIPFORMAT
  CF_NOFORMAT = 0
  CF_TEXT = 1
  CF_BITMAP = 2
  CF_METAFILEPICT = 3
  CF_SYLK = 4
  CF_DIF = 5
  CF_TIFF = 6
  CF_OEMTEXT = 7
  CF_DIB = 8
  CF_PALETTE = 9
  CF_PENDATA = 10
  CF_RIFF = 11
  CF_WAVE = 12
  CF_UNICODETEXT = 13
  CF_ENHMETAFILE = 14
  CF_HDROP = 15
  CF_LOCALE = 16
  CF_DIBV5 = 17

  CF_OWNERDISPLAY = &H80
  CF_DSPTEXT = &H81
  CF_DSPBITMAP = &H82
  CF_DSPMETAFILEPICT = &H83
  CF_DSPENHMETAFILE = &H8E
  CF_PRIVATEFIRST = &H200
  CF_PRIVATELAST = &H2FF
  CF_GDIOBJFIRST = &H300
  CF_GDIOBJLAST = &H3FF
End Enum

'API Declarations:
#If VBA7 Then
  'https://docs.microsoft.com/en-us/windows/win32/api/olectl/ns-olectl-pictdesc
  
  'DROPFiLES for CopyFiles()
  Private Type PICTDESC
      size As Long
      Type As Long
      hPic As LongPtr
      hPal As LongPtr
  End Type
  Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
  Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As CLIPFORMAT, ByVal hMem As LongPtr) As Long
  Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As LongPtr
  Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
  Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
  Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As Long
  Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
  Private Declare PtrSafe Function lstrcpyA Lib "kernel32" (ByVal lpString1 As LongPtr, ByVal lpString2 As LongPtr) As Long
  Private Declare PtrSafe Function lstrcpyW Lib "kernel32" (ByVal lpString1 As LongPtr, ByVal lpString2 As LongPtr) As Long
  Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As CLIPFORMAT) As Long
  Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
  Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As CLIPFORMAT) As LongPtr
  Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
  Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As LongPtr, IPic As stdole.IPicture) As Long
  Private Declare PtrSafe Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As LongPtr, ByVal lpszFile As String) As LongPtr
  Private Declare PtrSafe Function CopyImage Lib "user32" (ByVal Handle As LongPtr, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As LongPtr
  Private Declare PtrSafe Function EnumClipboardFormats Lib "user32" (ByVal Format As Long) As Long
  Private Declare PtrSafe Function GetClipboardFormatNameW Lib "user32" (ByVal Format As Long, ByVal lpszFormatName As LongPtr, ByVal cchMaxCount As Long) As Long
  Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As Long)
  Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
  Private Declare PtrSafe Function IIDFromString Lib "ole32" (ByVal lpszIID As LongPtr, ByRef iid As Any) As Long
  Private Declare PtrSafe Function SendMessage Lib "user32" (ByVal hwnd As LongPtr, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
  Private Declare PtrSafe Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As LongPtr, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
  Private Declare PtrSafe Function lstrlenW Lib "kernel32.dll" (ByVal lpString As LongPtr) As Long
  Private Declare PtrSafe Function lstrlenA Lib "kernel32.dll" (ByVal lpString As LongPtr) As Long
  Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
  Private Declare PtrSafe Function RegisterClipboardFormatA Lib "user32" (ByVal lpString As String) As Long
  Private Declare PtrSafe Function GetClipboardSequenceNumber Lib "user32" () As Long
  Private Declare PtrSafe Function WindowFromAccessibleObject Lib "oleacc" (ByVal pacc As IAccessible, ByRef phwnd As LongPtr) As Long
#Else
  Private Type PICTDESC
      size As Long
      Type As Long
      hPic As Long
      hPal As Long
  End Type
  Private Declare Function EmptyClipboard Lib "user32" () As Long
  Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As CLIPFORMAT, ByVal hMem As Long) As Long
  Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
  Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
  Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
  Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
  Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
  Private Declare Function lstrcpyA Lib "kernel32" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
  Private Declare Function lstrcpyW Lib "kernel32" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
  Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As CLIPFORMAT) As Long
  Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As CLIPFORMAT) As Long
  Private Declare Function CloseClipboard Lib "user32" () As Long
  Private Declare Function OleCreatePictureIndirect Lib "olepro32" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As stdole.IPicture) As Long
  Private Declare Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As Long, ByVal lpszFile As String) As Long
  Private Declare Function CopyImage Lib "user32" (ByVal Handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
  Private Declare Function EnumClipboardFormats Lib "user32" (ByVal Format As CLIPFORMAT) As CLIPFORMAT
  Private Declare Function GetClipboardFormatNameW Lib "user32" (ByVal Format As CLIPFORMAT, ByVal lpszFormatName As Long, ByVal cchMaxCount As Long) As Long
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
  Private Declare Function GetLastError Lib "kernel32" () As Long
  Private Declare Function IIDFromString Lib "ole32" (ByVal lpszIID As Long, ByRef iid As Any) As Long
  Private Declare Function SendMessage Lib "user32" (ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
  Private Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
  Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal lpString As Long) As Long
  Private Declare Function lstrlenA Lib "kernel32.dll" (ByVal lpString As Long) As Long
  Private Declare Function GetTickCount Lib "kernel32" () As Long
  Private Declare Function RegisterClipboardFormatA Lib "user32" (ByVal lpString As String) As Long
  Private Declare Function GetClipboardSequenceNumber Lib "user32" () as Long
  Private Declare Function WindowFromAccessibleObject Lib "oleacc" (ByVal pacc As IAccessible, ByRef phwnd As LongPtr) As Long
#End If


'POINTAPI struct is used in DROPFILES struct:
Private Type POINTAPI
  x As Long
  y As Long
End Type

'DROPFILES struct is used by CopyFiles() API:
Private Type DROPFILES
  pFiles As Long
  pt As POINTAPI
  fNC As Long
  fWide As Long
End Type

Private Enum GAlloc
  GMEM_FIXED = &H0
  GMEM_MOVEABLE = &H2
  GMEM_NOCOMPACT = &H10
  GMEM_NODISCARD = &H20
  GMEM_ZEROINIT = &H40
  GMEM_MODIFY = &H80
  GMEM_DISCARDABLE = &H100
  GMEM_NOT_BANKED = &H1000
  GMEM_SHARE = &H2000
  GMEM_DDESHARE = &H2000
  GMEM_NOTIFY = &H4000
  GMEM_LOWER = GMEM_NOT_BANKED
  GMEM_VALID_FLAGS = &H7F72
  GMEM_INVALID_HANDLE = &H8000
  GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
  GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
End Enum

'Declare a UDT to store a GUID for the stdole.IPicture OLE Interface
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

'************************************************************************************************
'* PUBLIC INTERFACE
'************************************************************************************************

'Get the ID of the current clipboard value
'@returns - The latest clipboard ID / sequence number
'@remark - This is incremented every time the clipboard changes. It can be used to track and identify when the clipboard has changed.
Public Property Get ClipboardID() As Long
  ClipboardID = GetClipboardSequenceNumber()
End Property

'Get or set the value of the current
'@param iFormat - The format to get/set the value of
'@returns Variant - The value stored in the clipboard
Public Property Get value(ByVal iFormat As CLIPFORMAT) As Variant
Attribute value.VB_UserMemId = 0
  #If VBA7 Then
    Dim hClipMemory     As LongPtr
    Dim lpClipMemory    As LongPtr
  #Else
    Dim hClipMemory     As Long
    Dim lpClipMemory    As Long
  #End If
  
  'Get pointer to clipboard data
  If Not CBool(OpenClipboardTimeout(getAppHwnd())) Then Call CriticalRaise("Value [GET]", "Unable to open clipboard (DllError: " & Err.LastDllError & ")")
  If Not CBool(IsClipboardFormatAvailable(iFormat)) Then Call CriticalRaise("Value [GET]", "Specified clipboard format is not available")
  hClipMemory = GetClipboardData(iFormat)
  If Not CBool(hClipMemory) Then Call CriticalRaise("Value [GET]", "Unable to obtain pointer to clipboard data (DllError: " & Err.LastDllError & ")")
  
  'Get the data from the supplied pointer
  Dim size As Long
  size = GlobalSize(hClipMemory)
  lpClipMemory = GlobalLock(hClipMemory)
  If CBool(lpClipMemory) Then
    Dim retVal As Variant
    Select Case iFormat
      Case CLIPFORMAT.CF_UNICODETEXT
        retVal = StringFromPointerW(lpClipMemory)
      Case CLIPFORMAT.CF_TEXT
        retVal = StringFromPointerA(lpClipMemory)
      Case Else
        retVal = BytesFromPointer(lpClipMemory, size)
    End Select
    Call CopyVariant(value, retVal)
  Else
    Call CriticalRaise("Value [GET]", "Unable to lock global memory (DllError: " & Err.LastDllError & ")")
  End If
  If CBool(GlobalUnlock(hClipMemory)) Then If Err.LastDllError <> 0 Then Call CriticalRaise("Value [GET]", "Unable to unlock global mem (DllError: " & Err.LastDllError & ")")
  If Not CBool(CloseClipboard()) Then Call CriticalRaise("Value [GET]", "Unable to close clipboard (DllError: " & Err.LastDllError & ")")
End Property
Public Property Let value(ByVal iFormat As CLIPFORMAT, ByVal v As Variant)
  #If VBA7 Then
    Dim hClipMemory     As LongPtr
    Dim lpClipMemory    As LongPtr
  #Else
    Dim hClipMemory     As Long
    Dim lpClipMemory    As Long
  #End If

  'Get the size of the block of memory
  Dim size As Long
  Select Case iFormat
    Case CLIPFORMAT.CF_TEXT
      size = LenB(v) + 1 '+1 to ensure room for ending NullChar
    Case CLIPFORMAT.CF_UNICODETEXT
      size = LenB(v) + 2 '+2 to ensure room for ending NullChar
    Case Else
      Dim data() As Byte
      data = v
      size = UBound(data) - LBound(data) + 1
  End Select
  
  'Allocate blob to global moveable memory
  hClipMemory = GlobalAlloc(GAlloc.GHND, size)
  If Not CBool(hClipMemory) Then Call CriticalRaise("Value [LET]", "Unable to allocate global memory (DllError: " & Err.LastDllError & ")")
  
  'Lock global memory
  lpClipMemory = GlobalLock(hClipMemory)
  If Not CBool(lpClipMemory) Then Call CriticalRaise("Value [LET]", "Unable to lock global memory (DllError: " & Err.LastDllError & ")")
  
  Select Case iFormat
    Case CLIPFORMAT.CF_TEXT
      lpClipMemory = lstrcpyA(lpClipMemory, StrPtr(v))
      If Not CBool(lpClipMemory) Then Call CriticalRaise("Value [LET]", "Unable to copy string to global memory (DllError: " & Err.LastDllError & ")")
    Case CLIPFORMAT.CF_UNICODETEXT
      lpClipMemory = lstrcpyW(lpClipMemory, StrPtr(v))
      If Not CBool(lpClipMemory) Then Call CriticalRaise("Value [LET]", "Unable to copy string to global memory (DllError: " & Err.LastDllError & ")")
    Case Else
      Call CopyMemory(lpClipMemory, VarPtr(data(LBound(data))), size)
  End Select
  If CBool(GlobalUnlock(hClipMemory)) Then If Err.LastDllError <> 0 Then Call CriticalRaise("Value [LET]", "Unable to unlock global mem. (DllError: " & Err.LastDllError & ")")
  
  'Try to set clipboard data:
  If Not CBool(OpenClipboardTimeout(getAppHwnd())) Then Call CriticalRaise("Value [LET]", "Unable to open clipboard (DllError: " & Err.LastDllError & ")")
  If Not CBool(EmptyClipboard()) Then Call CriticalRaise("Value [LET]", "Cannot empty clipboard (DllError: " & Err.LastDllError & ")")
  If Not CBool(SetClipboardData(iFormat, hClipMemory)) Then Call CriticalRaise("Value [LET]", "Unable to set clipboard data (DllError: " & Err.LastDllError & ")")
  If Not CBool(CloseClipboard()) Then Call CriticalRaise("Value [LET]", "Unable to close clipboard (DllError: " & Err.LastDllError & ")")
  
  'No need to call `GlobalFree` as system does this for us re: https://stackoverflow.com/questions/1264137/how-to-copy-string-to-clipboard-in-c#comment48533963_1264179
  'If CBool(GlobalFree(hClipMemory)) Then If Err.LastDllError <> 0 Then Call CriticalRaise("Value [GET]", "Cannot free Global mem (DllError: " & Err.LastDllError & ")")
End Property

'Waits for the clipboard to be available
Public Sub Await()
  While Not CBool(OpenClipboard(getAppHwnd()))
    DoEvents
  Wend
  Call CloseClipboard
End Sub

'Clears the clipboard
Public Sub Clear()
  If Not CBool(OpenClipboardTimeout(getAppHwnd())) Then Call CriticalRaise("Value [LET]", "Unable to open clipboard (DllError: " & Err.LastDllError & ")")
  If Not CBool(EmptyClipboard()) Then Call CriticalRaise("Value [LET]", "Cannot empty clipboard (DllError: " & Err.LastDllError & ")")
  If Not CBool(CloseClipboard()) Then Call CriticalRaise("Value [LET]", "Unable to close clipboard (DllError: " & Err.LastDllError & ")")
End Sub

'Get or set the clipboard as/to a string
'@returns String - The string stored on the clipboard
'@example ```vb
'stdClipboard.text = "Hello World"
'```
Public Property Get text() As String
  text = value(CF_UNICODETEXT)
End Property
Public Property Let text(ByVal s As String)
  value(CF_UNICODETEXT) = s
End Property

'Check whether a given format is available based on the clipboards current data.
'@param iFormat - The format to check support for.
'@returns - `True` if the format is allowed, `false` otherwise
'@remark - Formats which are technically "available" but have 0 bytes in size will return false.
Public Function IsFormatAvailable(ByVal iFormat As CLIPFORMAT) As Boolean
  Call OpenClipboardTimeout(getAppHwnd())
  If IsClipboardFormatAvailable(iFormat) Then
    Select Case iFormat
      Case CF_BITMAP, CF_ENHMETAFILE, CF_METAFILEPICT
        IsFormatAvailable = True
      Case Else
        'Ensure data exists in clipboard for format
        IsFormatAvailable = GetFormatSize(iFormat) > 0
    End Select
  End If
  Call CloseClipboard
End Function

'Get the size in bytes of a particular format
'@param format - The format to retrieve
'@returns - The number of bytes being the size of the format.
'@remark - CF_BITMAP and CF_ENHMETAFILE may be different due to the nature of obtaining their value via stdPicture
Public Function FormatSize(ByVal format As CLIPFORMAT) As Long
  Call OpenClipboardTimeout(getAppHwnd())
  If IsClipboardFormatAvailable(format) Then
    FormatSize = GetFormatSize(format)
  Else
    FormatSize = -1
  End If
  Call CloseClipboard
End Function

'Register a clipboard format and/or return it's identifier if already present.
'@param sFormatName - Name of format to obtain
'@returns - Clipboard format identifier
'@example ```vb
'CF_RTF = stdClipboard.registerFormat("Rich Text Format")
'stdClipboard.value(CF_RTF) = "{\rtf1\ansi{\fonttbl\f0\fswiss Helvetica;}\f0\pard" & _
'  "This is some {\b bold} text.\par" & _
'"}"
'```
'@example ```vb
'CF_HTML = stdClipboard.registerFormat("HTML Format")
'stdClipboard.value(CF_HTML) = _
'  "Version:0.9" & vbCrLf & _
'  "StartHTML:00000000" & vbCrLf & _
'  "EndHTML:00000000" & vbCrLf & _
'  "StartFragment:00000000" & vbCrLf & _
'  "EndFragment:00000000" & vbCrLf & _
'  "<html><body>" & vbCrLf & _
'  "<!--StartFragment -->" & vbCrLf & _
'  "  hello <b>world</b>" & vbCrLf & _
'  "<!--EndFragment-->\r\n" & vbCrLf & _
'  "</body>\r\n" & vbCrLf & _
'  "</html>"
'```
Public Function RegisterFormat(ByVal sFormatName As String) As Long
  RegisterFormat = RegisterClipboardFormatA(sFormatName)
End Function

'Obtain the available formats as a Collection of strings
'@returns Collection<String> - The format names you can use for the current clipboard data.
'@remark see `stdClipboard::formatIDs()` to get ids used in Value() from these data.
Public Property Get formats() As Collection
  Dim iFormat As CLIPFORMAT: iFormat = CF_NOFORMAT
  Dim ret As Collection: Set ret = New Collection
  'Loop over all formats and add to collection
  Call OpenClipboardTimeout(getAppHwnd())
  Do
    iFormat = EnumClipboardFormats(iFormat)
    If CBool(iFormat) Then
      ret.Add GetClipboardFormatName(iFormat), CStr(iFormat)
    Else
      Exit Do
    End If
  Loop
  Call CloseClipboard
  
  Set formats = ret
End Property

'Obtain the available format IDs which can be used with value() function
'@returns Collection<Long> - Formats allowed. We use Collection for easy enumeration with stdEnumerator.
'@remark See `stdClipboard::formats()` to get the corresponding names from these data.
Public Property Get formatIDs() As Collection
  Dim iFormat As CLIPFORMAT: iFormat = CF_NOFORMAT
  Dim ret As Collection: Set ret = New Collection
  'Loop over all formats and add to collection
  Call OpenClipboardTimeout(getAppHwnd())
  Do
    iFormat = EnumClipboardFormats(iFormat)
    If CBool(iFormat) Then
      ret.Add iFormat, GetClipboardFormatName(iFormat)
    Else
      Exit Do
    End If
  Loop
  Call CloseClipboard
  
  Set formatIDs = ret
End Property

'Get/Set the clipboard as/to an IPicture object
'@returns - The picture stored on the clipboard
Public Property Get Picture() As stdole.IPicture
  If IsClipboardFormatAvailable(CF_BITMAP) Then
    Set Picture = GetPictureFromClipboard(CF_BITMAP)
  ElseIf IsClipboardFormatAvailable(CF_ENHMETAFILE) Then
    Set Picture = GetPictureFromClipboard(CF_ENHMETAFILE)
  ElseIf IsClipboardFormatAvailable(CF_METAFILEPICT) Then
    Set Picture = GetPictureFromClipboard(CF_METAFILEPICT) 'Unsure whether this would work if CF_ENHMETAFILE doesn't work...
  'elseif IsClipboardFormatAvailable(CF_HDROP) then
  '  ???
  Else
    Call CriticalRaise("Picture", "Invalid clipboard format")
  End If
End Property
Public Property Set Picture(ByVal olePic As stdole.IPicture)
  'TODO: refactor away from using Excel specific functionality
  If Application.Name = "Microsoft Excel" Then
    Dim oApp As Object: Set oApp = Application
    Const sTmpPath = "C:\Temp\stdVBA"
    Dim size As Long
    Call olePic.SaveAsFile(sTmpPath, True, size)
    With oApp.ThisWorkbook.Sheets(1).Pictures.Insert(sTmpPath)
      Call .CopyPicture(xlScreen, XlCopyPictureFormat.xlPicture)
      Call .Delete
    End With
    Kill sTmpPath
  End If
End Property



'Set the clipboard to an Excel picture object, and then proceeds to wait until the format is available.
'@param iAppearance - xlPrint or xlScreen
'@param iFormat - xlBitmap or xlPicture
'@param xlPic - The picture to copy
Public Property Set XLPicture(Optional ByVal iAppearance As XlPictureAppearance = xlScreen, Optional ByVal iFormat As XlCopyPictureFormat = xlBitmap, ByVal xlPic As Excel.IPicture)
  Call Clear
  Call xlPic.CopyPicture(iAppearance, 2)
  While Not (IsFormatAvailable(CF_BITMAP) Or IsFormatAvailable(CF_METAFILEPICT))
    DoEvents
  Wend
End Property

'Copies a Excel shape as a picture, and then proceeds to wait until the format is available.
'@param iAppearance - xlPrint or xlScreen
'@param iFormat - xlBitmap or xlPicture
'@param xlShp - The shape to turn into an image
Public Property Set XLShapeAsPicture(Optional iAppearance As XlPictureAppearance = xlScreen, Optional iFormat As XlCopyPictureFormat = XlCopyPictureFormat.xlBitmap, ByVal xlShp As Excel.Shape)
  Call Clear
  Call xlShp.CopyPicture(iAppearance, iFormat)
  While Not (IsFormatAvailable(CF_BITMAP) Or IsFormatAvailable(CF_METAFILEPICT))
    DoEvents
  Wend
End Property


'Get or Set the clipboard to a selection of files paths. Paths provided as a collection for easy stdEnumerator wrapping.
'@returns Collection<String> - The collection of file paths stored in the clipboard
Public Property Get files() As Collection
  Set files = New Collection
  Dim sFiles() As String: sFiles = protGetFilesArr()
  Dim i As Long: For i = LBound(sFiles) To UBound(sFiles)
    files.Add sFiles(i)
  Next
End Property
Public Property Set files(ByVal vFiles As Collection)
  Dim sFiles As String, vFile As Variant
  For Each vFile In vFiles
    sFiles = sFiles & vFile & vbNullChar
  Next
  sFiles = sFiles & vbNullChar
  Call protSetFilesText(StrConv(sFiles, vbFromUnicode))
End Property

'Set the clipboard to a selection of Files passed in as an array
'@param vFiles as Variant<Array<Variant>> - A set of files as an array
Public Sub SetFilesArr(ByVal vFiles As Variant)
  Dim sFiles As String, i As Long
  For i = LBound(vFiles) To UBound(vFiles)
    sFiles = sFiles & vFiles(i) & vbNullChar
  Next
  sFiles = sFiles & vbNullChar
  Call protSetFilesText(sFiles)
End Sub

'Internal method for setting the clipboard to a selection of Files passed in as a raw file string
'Realistically it might be better for this method to use Value LET directly.
'@protected
'@param sFiles - Files string delimited by vbNullChar and ending in 2x vbNullChar
#If Win64 Then
Public Sub protSetFilesText(ByVal sFiles As String)
    MsgBox "Not implemented. Current version produces crash for unknown reasons.", vbCritical
End Sub
#Else
Public Sub protSetFilesText(ByVal sFiles As String)
  #If VBA7 Then
    Dim hClipMemory     As LongPtr
    Dim lpClipMemory    As LongPtr
  #Else
    Dim hClipMemory     As Long
    Dim lpClipMemory    As Long
  #End If
  Dim tDropFiles As DROPFILES

  'Allocate blob to global moveable memory
  hClipMemory = GlobalAlloc(GAlloc.GHND, Len(tDropFiles) + LenB(sFiles))

  If Not CBool(hClipMemory) Then Call CriticalRaise("protSetFilesText", "Unable to allocate global memory (DllError: " & Err.LastDllError & ")")
  
  'Lock global memory
  lpClipMemory = GlobalLock(hClipMemory)
  If Not CBool(lpClipMemory) Then Call CriticalRaise("protSetFilesText", "Unable to lock global memory (DllError: " & Err.LastDllError & ")")
  
  'Specify that the offset to where the files are, from the beginning of this structure, are at the end of this structure.
  'I.E. The Binary memory data is literally going to look like:
  '   {STRUCTDATA}{FILE1}\0{FILE2}\0{FILE3}\0\0
  tDropFiles.pFiles = Len(tDropFiles)
  Call CopyMemory(lpClipMemory, VarPtr(tDropFiles.pFiles), Len(tDropFiles))
  Call CopyMemory(lpClipMemory + Len(tDropFiles), StrPtr(sFiles), LenB(sFiles))
  
  If CBool(GlobalUnlock(hClipMemory)) Then If Err.LastDllError <> 0 Then Call CriticalRaise("protSetFilesText", "Unable to unlock global mem (DllError: " & Err.LastDllError & ")")
  
  'Try to set clipboard data:
  If Not CBool(OpenClipboardTimeout(getAppHwnd())) Then Call CriticalRaise("protSetFilesText", "Unable to open clipboard (DllError: " & Err.LastDllError & ")")
  If Not CBool(EmptyClipboard()) Then Call CriticalRaise("protSetFilesText", "Cannot empty clipboard (DllError: " & Err.LastDllError & ")")
  If Not CBool(SetClipboardData(CF_HDROP, hClipMemory)) Then Call CriticalRaise("protSetFilesText", "Unable to set clipboard data (DllError: " & Err.LastDllError & ")")
  If Not CBool(CloseClipboard()) Then Call CriticalRaise("protSetFilesText", "Unable to close clipboard (DllError: " & Err.LastDllError & ")")
  
  'No need to call `GlobalFree` as system does this for us re: https://stackoverflow.com/questions/1264137/how-to-copy-string-to-clipboard-in-c#comment48533963_1264179
  'If CBool(GlobalFree(hClipMemory)) Then If Err.LastDllError <> 0 Then Call CriticalRaise("Value [GET]", "Cannot free Global mem (DllError: " & Err.LastDllError & ")")
End Sub
#End If

'Get the file paths of files stored on the clipboard
'@protected
'@returns Array<String> - The array of file paths stored in the clipboard
Public Function protGetFilesArr() As String()
  #If VBA7 Then
    Dim hClipMemory     As LongPtr
    Dim lpClipMemory    As LongPtr
  #Else
    Dim hClipMemory     As Long
    Dim lpClipMemory    As Long
  #End If
  
  'Get pointer to clipboard data
  If Not CBool(OpenClipboardTimeout(getAppHwnd())) Then Call CriticalRaise("protGetFilesArr", "Unable to open clipboard (DllError: " & Err.LastDllError & ")")
  If Not CBool(IsClipboardFormatAvailable(CF_HDROP)) Then Call CriticalRaise("protGetFilesArr", "Specified clipboard format is not available")
  hClipMemory = GetClipboardData(CF_HDROP)
  Dim nFiles As Long: nFiles = DragQueryFile(hClipMemory, -1&, "", 0)
  If Not CBool(hClipMemory) Then Call CriticalRaise("protGetFilesArr", "Unable to obtain pointer to clipboard data (DllError: " & Err.LastDllError & ")")

  
  Dim sRet() As String: ReDim sRet(1 To nFiles)
  Dim sFileName As String
  
  Dim i As Long
  For i = 1 To nFiles
    sFileName = Space(260)
    Call DragQueryFile(hClipMemory, i - 1, sFileName, 260)
    sRet(i) = VBA.Trim(sFileName)
  Next
  
  protGetFilesArr = sRet
  If Not CBool(CloseClipboard()) Then Call CriticalRaise("protGetFilesArr", "Unable to close clipboard (DllError: " & Err.LastDllError & ")")
End Function


'Raise a paste event to a window with handle hWnd
'@param hWnd - Window to send paste event to
#If VBA7 Then
Public Sub SendPaste(ByVal hwnd As LongPtr)
#Else
Public Sub SendPaste(ByVal hwnd As Long)
#End If
  Const WM_PASTE As Long = &H302
  Call SendMessage(hwnd, WM_PASTE, 0, 0)
End Sub

'Raise a copy event to a window with handle hWnd
'@param hWnd - Window to send copy event to
#If VBA7 Then
Public Sub SendCopy(ByVal hwnd As LongPtr)
#Else
Public Sub SendCopy(ByVal hwnd As Long)
#End If
  Const WM_COPY As Long = &H301
  Call SendMessage(hwnd, WM_COPY, 0, 0)
End Sub



'Converts xlCopyPictureFormat to CLIPFORMAT
'Public Enum XlCopyPictureFormat
'  xlBitmap = 2
'  xlPicture = -4147
'End Enum
'@param CopyPictureFormat - The xlCopyPictureFormat to convert
'@returns CLIPFORMAT - The CLIPFORMAT to use
Public Function GetClipFormatFromCopyPictureFormat(ByVal CopyPictureFormat As Long) As CLIPFORMAT
  Const xxlBitmap = 2
  Const xxlPicture = -4147
  Select Case CopyPictureFormat
    Case xxlBitmap
      GetClipFormatFromCopyPictureFormat = CLIPFORMAT.CF_BITMAP
    Case xxlPicture
      GetClipFormatFromCopyPictureFormat = CLIPFORMAT.CF_ENHMETAFILE
  End Select
End Function



'******************************************************************
'* HELPERS
'******************************************************************

'Get the size of the data stored at the pointer location (not necessarily the size of the file!)
'@private
'@param iFormat - Format to get size of
'@returns - Number of bytes
'@remark - CF_BITMAP, CF_ENHMETAFILE and CF_METAFILEPICT will sometimes crash on `GlobalSize()`. The size of these is `0` though.
'The data is not stored in memory under the handle. Instead the handle returned by `GetClipboardData()` returns a handle 
'(HBITMAP or HENHMETAFILE) which internally have very different structures and handling mechanisms.
Private Function GetFormatSize(ByRef iFormat As CLIPFORMAT) As Long
  Select Case iFormat
    Case CF_BITMAP, CF_ENHMETAFILE, CF_METAFILEPICT
      GetFormatSize = 0
    Case Else
      #If VBA7 Then
        Dim hClipMemory     As LongPtr
      #Else
        Dim hClipMemory     As Long
      #End If
      hClipMemory = GetClipboardData(iFormat)
      If CBool(hClipMemory) Then
        'Ensure data exists in clipboard for format
        GetFormatSize = GlobalSize(hClipMemory)
      End If
  End Select
End Function

'Raises an error in stdError if available, else it will be raised in Err
'@private
'@param sFuncName - The name of the function raising the error
'@param sMessage - The error message to raise
Private Sub CriticalRaise(ByVal sFuncName As String, ByVal sMessage As String)
  Call CloseClipboard
  'If stdError exists
    If VarType(stdError) Then
      Call stdError.Raise("stdClipboard::" & sFuncName & " - " & sMessage)
    Else
      Call Err.Raise(IIf(Err.LastDllError < 0, Err.LastDllError, &H80070000 Or Err.LastDllError), "stdClipboard::" & sFuncName, sMessage)
    End If
    End
End Sub


'Returns a byte array stored at a pointer using RtlMoveMemory
'@private
'@param pointer - The pointer to the data
'@param size - The size of the data
'@returns Array<Byte> - The data stored at the pointer
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
    Call CriticalRaise("BytesFromPointer [PRIVATE]", "No size supplied.")
  End If
  BytesFromPointer = buff
End Function

'Obtain the Unicode string held at a particular pointer
'@private
'@param hLPTSTR - The pointer to the unicode string
'@returns - The unicode string stored at the pointer
#If VBA7 Then
Private Function StringFromPointerW(ByVal hLPTSTR As LongPtr) As String
#Else
Private Function StringFromPointerW(ByVal hLPTSTR As Long) As String
#End If
  If lstrlenW(hLPTSTR) > 0 Then
    StringFromPointerW = BytesFromPointer(hLPTSTR, lstrlenW(hLPTSTR) * 2)
  Else
    StringFromPointerW = ""
  End If
End Function

'Obtain the string held at a particular pointer
'@private
'@param hLPTSTR - The pointer to the ascii string
'@returns - The ascii string stored at the pointer
#If VBA7 Then
Private Function StringFromPointerA(ByVal hLPTSTR As LongPtr) As String
#Else
Private Function StringFromPointerA(ByVal hLPTSTR As Long) As String
#End If
  If lstrlenA(hLPTSTR) > 0 Then
    StringFromPointerA = BytesFromPointer(hLPTSTR, lstrlenA(hLPTSTR))
  Else
    StringFromPointerA = ""
  End If
End Function

'Get the name of a clipboard format
'@private
'@param iFormat - The clipboard format to get the name of
'@returns - The name of the clipboard format
Private Function GetClipboardFormatName(ByVal iFormat As Long) As String
  Select Case iFormat
    Case CLIPFORMAT.CF_UNICODETEXT: GetClipboardFormatName = "CF_UNICODETEXT"
    Case CLIPFORMAT.CF_TEXT: GetClipboardFormatName = "CF_TEXT"
    Case CLIPFORMAT.CF_NOFORMAT: GetClipboardFormatName = "CF_NOFORMAT"
    Case CLIPFORMAT.CF_BITMAP: GetClipboardFormatName = "CF_BITMAP"
    Case CLIPFORMAT.CF_DIB: GetClipboardFormatName = "CF_DIB"
    Case CLIPFORMAT.CF_DIBV5: GetClipboardFormatName = "CF_DIBV5"
    Case CLIPFORMAT.CF_DIF: GetClipboardFormatName = "CF_DIF"
    Case CLIPFORMAT.CF_DSPBITMAP: GetClipboardFormatName = "CF_DSPBITMAP"
    Case CLIPFORMAT.CF_DSPENHMETAFILE: GetClipboardFormatName = "CF_DSPENHMETAFILE"
    Case CLIPFORMAT.CF_DSPMETAFILEPICT: GetClipboardFormatName = "CF_DSPMETAFILEPICT"
    Case CLIPFORMAT.CF_DSPTEXT: GetClipboardFormatName = "CF_DSPTEXT"
    Case CLIPFORMAT.CF_ENHMETAFILE: GetClipboardFormatName = "CF_ENHMETAFILE"
    Case CLIPFORMAT.CF_GDIOBJFIRST: GetClipboardFormatName = "CF_GDIOBJFIRST"
    Case CLIPFORMAT.CF_GDIOBJLAST: GetClipboardFormatName = "CF_GDIOBJLAST"
    Case CLIPFORMAT.CF_HDROP: GetClipboardFormatName = "CF_HDROP"
    Case CLIPFORMAT.CF_LOCALE: GetClipboardFormatName = "CF_LOCALE"
    Case CLIPFORMAT.CF_METAFILEPICT: GetClipboardFormatName = "CF_METAFILEPICT"
    Case CLIPFORMAT.CF_OEMTEXT: GetClipboardFormatName = "CF_OEMTEXT"
    Case CLIPFORMAT.CF_OWNERDISPLAY: GetClipboardFormatName = "CF_OWNERDISPLAY"
    Case CLIPFORMAT.CF_PALETTE: GetClipboardFormatName = "CF_PALETTE"
    Case CLIPFORMAT.CF_PENDATA: GetClipboardFormatName = "CF_PENDATA"
    Case CLIPFORMAT.CF_PRIVATEFIRST: GetClipboardFormatName = "CF_PRIVATEFIRST"
    Case CLIPFORMAT.CF_PRIVATELAST: GetClipboardFormatName = "CF_PRIVATELAST"
    Case CLIPFORMAT.CF_RIFF: GetClipboardFormatName = "CF_RIFF"
    Case CLIPFORMAT.CF_SYLK: GetClipboardFormatName = "CF_SYLK"
    Case CLIPFORMAT.CF_TIFF: GetClipboardFormatName = "CF_TIFF"
    Case CLIPFORMAT.CF_WAVE: GetClipboardFormatName = "CF_WAVE"
    Case Else
      Dim sName As String: sName = ""
      Dim sBuff As String: sBuff = Space(200)
      If GetClipboardFormatNameW(iFormat, StrPtr(sBuff), 200) Then
        GetClipboardFormatName = Mid(sBuff, 1, InStr(1, sBuff, vbNullChar) - 1)
      Else
        GetClipboardFormatName = "Unknown Name"
      End If
  End Select
End Function


'******************************************************************
'* PICTURE HELPERS
'******************************************************************

'Obtain a picture object from the clipboard
'@private
'@param lPicType - The type of picture to obtain
'@returns - The picture object stored on the clipboard
Private Function GetPictureFromClipboard(Optional ByVal lPicType As CLIPFORMAT) As stdole.IPicture
    'Define pointers
    #If VBA7 Then
      Dim hPtr As LongPtr, hCopy As LongPtr
    #Else
      Dim hPtr As Long, hCopy As Long
    #End If
    
    'Constants
    Const IMAGE_BITMAP = 0
    Const LR_COPYRETURNORG = &H4
    
    'Check if the clipboard contains the required format
    If IsClipboardFormatAvailable(lPicType) Then
      'Get access to the clipboard
      If OpenClipboardTimeout(getAppHwnd()) Then
        'Get a handle to the image data
        hPtr = GetClipboardData(lPicType)
        If Not CBool(hPtr) Then Call CriticalRaise("GetPictureFromClipboard [PRIVATE]", "Unable to get clipboard data handle (DllError: " & Err.LastDllError & ")")

        'Create our own copy of the image on the clipboard, in the appropriate format.
        If lPicType = CLIPFORMAT.CF_BITMAP Then
            hCopy = CopyImage(hPtr, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG)
        Else
            hCopy = CopyEnhMetaFile(hPtr, vbNullString)
        End If

        'Release the clipboard to other programs
        If Not CBool(CloseClipboard()) Then Call CriticalRaise("GetPictureFromClipboard [PRIVATE]", "Unable to close clipboard (DllError: " & Err.LastDllError & ")")

        'If we got a handle to the image, convert it into a Picture object and return it
        If hCopy Then
          Set GetPictureFromClipboard = CreatePicture(hCopy, 0, lPicType)
        Else
          Call CriticalRaise("GetPictureFromClipboard [PRIVATE]", "Unable to obtain image handle (DllError: " & Err.LastDllError & ")")
        End If
      Else
        Call CriticalRaise("GetPictureFromClipboard [PRIVATE]", "Unable to open clipboard (DllError: " & Err.LastDllError & ")")
      End If
    Else
      Call CriticalRaise("GetPictureFromClipboard [PRIVATE]", "Clipboard format is not available")
    End If
End Function

'Obtains an stdole.IPicture object from the handle supplied. This is used in conjunction with GetPictureFromClipboard()
'@private
'@param hPic - The handle to the picture
'@param hPal - The handle to the palette
'@param lPicType - The type of picture to obtain
'@returns - The picture object stored on the clipboard
#If VBA7 Then
Private Function CreatePicture(ByVal hPic As LongPtr, ByVal hPal As LongPtr, ByVal lPicType As CLIPFORMAT) As stdole.IPicture
#Else
Private Function CreatePicture(ByVal hPic As Long, ByVal hPal As Long, ByVal lPicType As CLIPFORMAT) As stdole.IPicture
#End If
  'IID for IDispatch (used while creating stdole.IPicture)
  Const IID_IDISPATCH As String = "{00020400-0000-0000-C000-000000000046}"
  
  ' stdole.IPicture requires a reference to "OLE Automation"
  Dim r As Long, uPicInfo As PICTDESC, iid As GUID, IPic As stdole.IPicture
  
  ' OLE Picture types
  Const PICTYPE_BITMAP = 1
  Const PICTYPE_METAFILE = 2
  Const PICTYPE_ICON = 3
  Const PICTYPE_ENHMETAFILE = 4

  ' Create the Interface GUID (for the stdole.IPicture interface)
  If CBool(IIDFromString(StrPtr(IID_IDISPATCH), iid)) Then Call CriticalRaise("CreatePicture [PRIVATE]", "Cannot get IDispatch IID (DllError: " & Err.LastDllError & ")")

  ' Fill uPicInfo with necessary parts.
  With uPicInfo
      .size = Len(uPicInfo)                                                              ' Length of structure.
      .Type = IIf(lPicType = CLIPFORMAT.CF_BITMAP, PICTYPE_BITMAP, PICTYPE_ENHMETAFILE)  ' Type of Picture
      .hPic = hPic                                                                       ' Handle to image.
      .hPal = IIf(lPicType = CLIPFORMAT.CF_BITMAP, hPal, 0)                              ' Handle to palette (if bitmap).
  End With

  ' Create the Picture object.
  If Not CBool(OleCreatePictureIndirect(uPicInfo, iid, True, IPic)) Then
    Set CreatePicture = IPic
  Else
    Call CriticalRaise("CreatePicture [PRIVATE]", "An error occurred in OleCreatePictureIndirect. Error code (" & r & ": " & fnOLEError(r) & ")")
  End If
End Function

'Given an OLE Error ID, obtain a description of what this ID means:
'@private
'@param lErrNum - The error number to obtain a description for
'@returns - The description of the error
Private Function fnOLEError(ByVal lErrNum As Long) As String
  'OLECreatePictureIndirect return values
  Const E_ABORT = &H80004004
  Const E_ACCESSDENIED = &H80070005
  Const E_FAIL = &H80004005
  Const E_HANDLE = &H80070006
  Const E_INVALIDARG = &H80070057
  Const E_NOINTERFACE = &H80004002
  Const E_NOTIMPL = &H80004001
  Const E_OUTOFMEMORY = &H8007000E
  Const E_POINTER = &H80004003
  Const E_UNEXPECTED = &H8000FFFF
  Const S_OK = &H0
  
  
  Select Case lErrNum
    Case E_ABORT
      fnOLEError = " Aborted"
    Case E_ACCESSDENIED
      fnOLEError = " Access Denied"
    Case E_FAIL
      fnOLEError = " General Failure"
    Case E_HANDLE
      fnOLEError = " Bad/Missing Handle"
    Case E_INVALIDARG
      fnOLEError = " Invalid Argument"
    Case E_NOINTERFACE
      fnOLEError = " No Interface"
    Case E_NOTIMPL
      fnOLEError = " Not Implemented"
    Case E_OUTOFMEMORY
      fnOLEError = " Out of Memory"
    Case E_POINTER
      fnOLEError = " Invalid Pointer"
    Case E_UNEXPECTED
      fnOLEError = " Unknown Error"
    Case S_OK
      fnOLEError = " Success!"
  End Select
End Function

'Copies one variant to a destination
'@private
'@param dest - The destination to copy to
'@param value - The value to copy
Private Sub CopyVariant(ByRef dest As Variant, ByVal value As Variant)
  If IsObject(value) Then
    Set dest = value
  Else
    dest = value
  End If
End Sub


'Tries to open the clipbaord with a timeout
'@private
'@param hwnd - The window handle to open the clipboard for
'@param dwTimeoutMilliseconds - The timeout in milliseconds
'@returns - `True` if the clipboard was opened, `False` if the timeout expired
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
        hWnd = oApp.Hwnd
    Case "Microsoft Word"
        hWnd = oApp.ActiveWindow.Hwnd
    Case "Microsoft Access"
        hWnd = oApp.hWndAccessApp
    Case "Microsoft PowerPoint"
        Call WindowFromAccessibleObject(oApp.CommandBars("Status Bar"), hWnd)
    Case "Outlook"
        Call WindowFromAccessibleObject(oApp.ActiveWindow.CommandBars("Status Bar"), hWnd)
    Case "Microsoft Publisher"
        Call WindowFromAccessibleObject(oApp.CommandBars("Status Bar"), hWnd)
    Case Else
        Err.Raise 1, "stdAcc::CreateFromApplication()", "No implementation for getting application window of " & oApp.Name
    End Select
End If
getAppHwnd = hWnd
End Function