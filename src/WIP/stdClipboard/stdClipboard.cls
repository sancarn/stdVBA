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
  CF_BITMAP = 2
  CF_DIB = 8
  CF_DIBV5 = 17
  CF_DIF = 5
  CF_DSPBITMAP = &H82
  CF_DSPENHMETAFILE = &H8E
  CF_DSPMETAFILEPICT = &H83
  CF_DSPTEXT = &H81
  CF_ENHMETAFILE = &H14
  CF_GDIOBJFIRST = &H300
  CF_GDIOBJLAST = &H3FF
  CF_HDROP = 15
  CF_LOCALE = 16
  CF_METAFILEPICT = 3
  CF_OEMTEXT = 7
  CF_OWNERDISPLAY = &H80
  CF_PALETTE = 9
  CF_PENDATA = 10
  CF_PRIVATEFIRST = &H200
  CF_PRIVATELAST = &H2FF
  CF_RIFF = 11
  CF_SYLK = 4
  CF_TEXT = 1
  CF_TIFF = 6
  CF_UNICODETEXT = 13
  CF_WAVE = 12
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
  Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
  Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As Long
  Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
  Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As Long
  Private Declare PtrSafe Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As LongPtr, ByVal lpString2 As LongPtr) As Long
  Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As CLIPFORMAT) As Long
  Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As Long
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
  Private Declare PtrSafe Function SendMessage Lib "user32" (ByVal hWnd As LongPtr, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
  Private Declare PtrSafe Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As LongPtr, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
  Private Declare PtrSafe Function lstrlenW Lib "kernel32.dll" (ByVal lpString As LongPtr) As Long
  Private Declare PtrSafe Function lstrlenA Lib "kernel32.dll" (ByVal lpString As LongPtr) As Long
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
  Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Lont, ByVal lpString2 As Lont) As Long
  Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As CLIPFORMAT) As Long
  Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
  Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As CLIPFORMAT) As Long
  Private Declare Function CloseClipboard Lib "user32" () As Long
  Private Declare Function OleCreatePictureIndirect Lib "olepro32" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As stdole.IPicture) As Long
  Private Declare Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As Long, ByVal lpszFile As String) As Long
  Private Declare Function CopyImage Lib "user32" (ByVal Handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
  Private Declare Function EnumClipboardFormats Lib "User32" (ByVal Format As CLIPFORMAT) As CLIPFORMAT
  Private Declare Function GetClipboardFormatNameW Lib "User32" (ByVal Format As CLIPFORMAT, ByVal lpszFormatName As Long, ByVal cchMaxCount As Long) As Long
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
  Private Declare Function GetLastError Lib "kernel32" () As Long
  Private Declare Function IIDFromString Lib "ole32" ( ByVal lpszIID As Long, ByRef iid As Any) As Long
  Private Declare Function SendMessage Lib "user32" (ByVal hWnd as Long, ByVal iMsg as long, ByVal wParam as long, ByVal lParam as long)
  Private Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
  Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal lpString As Long) As Long
  Private Declare Function lstrlenA Lib "kernel32.dll" (ByVal lpString As Long) As Long
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

'The API format types we're interested in
Const IMAGE_BITMAP = 0
Const LR_COPYRETURNORG = &H4

'IID for IDispatch (used while creating stdole.IPicture)
Private Const IID_DISPATCH As String = "{00020400-0000-0000-C000-000000000046}"

Const WM_COPY As Long = &H301
Const WM_PASTE As Long = &H302



'************************************************************************************************
'* PUBLIC INTERFACE
'************************************************************************************************

'Get or set the value of the current
'@param
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
  If Not CBool(OpenClipboard(0)) Then Call CriticalRaise("Value [GET]", "Unable to open clipboard")
  If Not CBool(IsClipboardFormatAvailable(iFormat)) Then Call CriticalRaise("Value [GET]", "Specified clipboard format is not available")
  hClipMemory = GetClipboardData(iFormat)
  If Not CBool(hClipMemory) Then Call CriticalRaise("Value [GET]", "Unable to obtain pointer to clipboard data")
  
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
    Call CriticalRaise("Value [GET]", "Unable to lock global memory")
  End If
  If Not CBool(GlobalUnlock(hClipMemory)) Then Call CriticalRaise("Value [GET]", "Unable to unlock global mem.")
  If Not CBool(CloseClipboard()) Then Call CriticalRaise("Value [GET]", "Unable to close clipboard. Error Code: " & GetLastError())
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
      size = LenB(v)
    Case Else
      Dim data() As Byte
      data = v
      size = UBound(data) - LBound(data) + 1
  End Select
  
  'Allocate blob to global moveable memory
  hClipMemory = GlobalAlloc(GAlloc.GHND, size)
  If Not CBool(hClipMemory) Then Call CriticalRaise("Value [LET]", "Unable to allocate global memory")
  
  'Lock global memory
  lpClipMemory = GlobalLock(hClipMemory)
  If Not CBool(lpClipMemory) Then Call CriticalRaise("Value [LET]", "Unable to lock global memory")
  
  Select Case iFormat
    Case CLIPFORMAT.CF_TEXT
      lpClipMemory = lstrcpy(lpClipMemory, strPtr(v))
      If Not CBool(lpClipMemory) Then Call CriticalRaise("Value [LET]", "Unable to copy string to global memory")
    Case Else
      Call CopyMemory(lpClipMemory, VarPtr(data(LBound(data))), size)
  End Select
  If Not CBool(GlobalUnlock(hClipMemory)) Then Call CriticalRaise("Value [LET]", "Unable to unlock global mem.")

  'Try to set clipboard data:
  If Not CBool(OpenClipboard(0)) Then Call CriticalRaise("Value [LET]", "Unable to open clipboard")
  If Not CBool(EmptyClipboard()) Then Call CriticalRaise("Value [LET]", "Cannot empty clipboard")
  If Not CBool(SetClipboardData(iFormat, hGlobalMemory)) Then Call CriticalRaise("Value [LET]", "Unable to set clipboard data")
  If Not CBool(CloseClipboard()) Then Call CriticalRaise("Value [LET]", "Unable to close clipboard. Error Code: " & GetLastError())
End Property

'Check whether a given format is available based on the clipboards current data.
'@param {ByVal CLIPFORMAT} - The format to check support for.
'@returns {Boolean} True if the format is allowed, false otherwise
Public Function IsFormatAvailable(ByVal iFormat As CLIPFORMAT) As Boolean
  OpenClipboard 0
  IsFormatAvailable = IsClipboardFormatAvailable(iFormat)
  CloseClipboard
End Function

'Obtain the available formats as a collection of strings.
'@returns {Collection} Formats allowed. We use string for easy enumeration with stdEnumerator.
Public Function GetFormats() As Collection
  Dim iFormat As CLIPFORMAT: iFormat = CF_NOFORMAT
  Set GetFormats = New Collection
  'Loop over all formats and add to collection
  Do
    iFormat = EnumClipboardFormats(iFormat)
    If iFormat = 0 Then Exit Do
    Call GetFormats.add(CStr(iFormat), iFormat)
  Loop
End Function

'Obtain the available formats as a string, delimited by your chosen delimiter
'@param {Optional ByVal String} Delimiter to use between format types
'@returns {String} The formats you can use for the current clipboard data.
Public Function GetFormatsString(Optional delim As String = ",") As String
  Dim iFormat As CLIPFORMAT: iFormat = CF_NOFORMAT
  Dim sBuff As String, sRet As String: sRet = ""
  'Loop over all formats and add to collection
  Do
    iFormat = EnumClipboardFormats(iFormat)
    sBuff = Space(200)
    If CBool(iFormat) Then
      If GetClipboardFormatNameW(iFormat, strPtr(buffer), 200) Then
        sRet = sRet & delim & Trim(sBuff) & " (" & iFormat & ")"
      Else
        sRet = sRet & delim & "Unknown Name (" & iFormat & ")"
      End If
    Else
      Exit Do
    End If
  Loop
  GetFormatsText = mid(sRet, Len(delim) + 1)
End Function

'Get or set the clipboard as/to an IPicture object
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
    Const sTmpPath = "C:\Temp\stdVBA"
    Dim size As Long
    Call olePic.SaveAsFile(sTmpPath, True, size)
    With Application.ThisWorkbook.Sheets(1).Pictures.Insert(fNameAndPath)
      Call .CopyPicture(xlScreen, XlCopyPictureFormat.XLPicture)
      Call .Delete
    End With
    Kill sTmpPath
  End If
End Property

'Set the clipboard to an Excel picture object. This is a helper method and does nothing more than call xlPic.CopyPicture(1,2).
'@param {ByVal Excel.IPicture} The picture to set the clipboard too.
Public Property Set XLPicture(ByVal xlPic As Excel.IPicture)
  'xlScreen = 1, xlBitmap = 2
  Call xlPic.CopyPicture(1, 2)
End Property

'Get or Set the clipboard to a selection of files paths. Paths provided as a collection for easy stdEnumerator wrapping.
Public Property Get Files() As Collection
  Set Files = New Collection
  Dim sFiles() As String: sFiles = protGetFilesArr()
  Dim i As Long: For i = LBound(sFiles) To UBound(sFiles)
    Files.add sFiles(i)
  Next
End Property
Public Property Set Files(ByVal vFiles As Collection)
  Dim sFiles As String, vFile As Variant
  For Each vFile In vFiles
    sFiles = sFiles & vFile & vbNullChar
  Next
  sFiles = sFiles & vbNullChar
  Call protSetFilesText(sFiles)
End Property

'Set the clipboard to a selection of Files passed in as an array
'@param {ByVal Variant()} A set of files as an array
Public Sub SetFilesArr(ByVal vFiles As Variant)
  Dim sFiles As String
  For i = LBound(vFiles) To UBound(vFiles)
    sFiles = sFiles & vFiles(i) & vbNullChar
  Next
  sFiles = sFiles & vbNullChar
  Call protSetFilesText(sFiles)
End Sub

'Internal method for setting the clipboard to a selection of Files passed in as a raw file string
'Realistically it might be better for this method to use Value LET directly.
'@protected
'@param {ByVal String} Files string delimited by vbNullChar and ending in 2x vbNullChar
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
  If Not CBool(hClipMemory) Then Call CriticalRaise("Value [LET]", "Unable to allocate global memory")
  
  'Lock global memory
  lpClipMemory = GlobalLock(hClipMemory)
  If Not CBool(lpClipMemory) Then Call CriticalRaise("Value [LET]", "Unable to lock global memory")
  
  'Specify that the offset to where the files are, from the beginning of this structure, are at the end of this structure.
  'I.E. The Binary memory data is literally going to look like:
  '   {STRUCTDATA}{FILE1}\0{FILE2}\0{FILE3}\0\0
  tDropFiles.pFiles = Len(tDropFiles)
  Call CopyMemory(lpClipMemory, VarPtr(tDropFiles.pFiles), Len(tDropFiles))
  Call CopyMemory(lpClipMemory + Len(tDropFiles), strPtr(sFiles), LenB(sFiles))
  
  If Not CBool(GlobalUnlock(hClipMemory)) Then Call CriticalRaise("Value [LET]", "Unable to unlock global mem.")

  'Try to set clipboard data:
  If Not CBool(OpenClipboard(0)) Then Call CriticalRaise("Value [LET]", "Unable to open clipboard")
  If Not CBool(EmptyClipboard()) Then Call CriticalRaise("Value [LET]", "Cannot empty clipboard")
  If Not CBool(SetClipboardData(CF_HDROP, hClipMemory)) Then Call CriticalRaise("Value [LET]", "Unable to set clipboard data")
  If Not CBool(CloseClipboard()) Then Call CriticalRaise("Value [LET]", "Unable to close clipboard. Error Code: " & GetLastError())
End Sub

'Get the file paths of files stored on the clipboard
'@protected
'@param {ByVal String} Files string delimited by vbNullChar and ending in 2 vbNullChar
'@returns {String()} The array of file paths stored in the clipboard
Public Function protGetFilesArr() As String()
  #If VBA7 Then
    Dim hClipMemory     As LongPtr
    Dim lpClipMemory    As LongPtr
  #Else
    Dim hClipMemory     As Long
    Dim lpClipMemory    As Long
  #End If
  
  'Get pointer to clipboard data
  If Not CBool(OpenClipboard(0)) Then Call CriticalRaise("Value [GET]", "Unable to open clipboard")
  If Not CBool(IsClipboardFormatAvailable(CF_HDROP)) Then Call CriticalRaise("Value [GET]", "Specified clipboard format is not available")
  hClipMemory = GetClipboardData(CF_HDROP)
  Dim nFiles As Long: nFiles = DragQueryFile(hClipMemory, -1&, "", 0)
  If Not CBool(hClipMemory) Then Call CriticalRaise("Value [GET]", "Unable to obtain pointer to clipboard data")
  
  Dim sRet() As String: ReDim sRet(1 To nFiles)
  Dim sFileName As String
  
  Dim i As Long
  For i = 1 To nFiles
    sFileName = Space(260)
    Call DragQueryFile(hClipMemory, i - 1, sFileName, 260)
    sRet(i) = Trim(sFileName)
  Next
  
  protGetFilesArr = sRet
  If Not CBool(CloseClipboard()) Then Call CriticalRaise("Value [GET]", "Unable to close clipboard. Error Code: " & GetLastError())
End Function


'Raise a paste event to a window with handle hWnd
'@param {ByVal LongPtr/Long} hWnd to send paste event to
#If VBA7 Then
Public Sub SendPaste(ByVal hWnd As LongPtr)
#Else
Public Sub SendPaste(ByVal hWnd As Long)
#End If
  Call SendMessage(hWnd, WM_PASTE, 0, 0)
End Sub

'Raise a copy event to a window with handle hWnd
'@param {ByVal LongPtr/Long} hWnd to send copy event to
#If VBA7 Then
Public Sub SendCopy(ByVal hWnd As LongPtr)
#Else
Public Sub SendCopy(ByVal hWnd As Long)
#End If
  Call SendMessage(hWnd, WM_COPY, 0, 0)
End Sub



'Converts xlCopyPictureFormat to CLIPFORMAT
'Public Enum XlCopyPictureFormat
'  xlBitmap = 2
'  xlPicture = -4147
'End Enum
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

'Raises an error in stdError if available, else it will be raised in Err
'@param {ByVal String} The Function name, to be concatenated with "stdClipboard::" to determine the source.
'@param {ByVal String} The message to relay to the user.
Private Sub CriticalRaise(ByVal sFuncName As String, ByVal sMessage As String)
  'If stdError exists
    If VarType(stdError) Then
      Call stdError.Raise("stdClipboard::" & sFuncName & " - " & sMessage)
    Else
      Call Err.Raise(1, "stdCallback::" & sFuncName, sMessage)
    End If
    End
End Sub


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
    Call CriticalRaise("BytesFromPointer [PRIVATE]", "No size supplied.")
  End If
  BytesFromPointer = buff
End Function

'Obtain the Unicode string held at a particular pointer
'@param {ByVal LongPtr/Long} The pointer to copy bytes at
'@returns {String} The string copied
#If VBA7 Then
Private Function StringFromPointerW(ByVal hLPTSTR As LongPtr) As String
#Else
Private Function StringFromPointerW(ByVal hLPTSTR As Long) As String
#End If
  StringFromPointerW = BytesFromPointer(hLPTSTR, lstrlenW(hLPTSTR) * 2)
End Function

'Obtain the string held at a particular pointer
'@param {ByVal LongPtr/Long} The pointer to copy bytes at
'@returns {String} The string copied
#If VBA7 Then
Private Function StringFromPointerA(ByVal hLPTSTR As LongPtr) As String
#Else
Private Function StringFromPointerA(ByVal hLPTSTR As Long) As String
#End If
  StringFromPointerA = BytesFromPointer(hLPTSTR, lstrlenA(hLPTSTR))
End Function

'******************************************************************
'* PICTURE HELPERS
'******************************************************************

'Obtain a picture object from the clipboard
'@param {ByVal CLIPFORMAT} - The format of the picture stored in the clipboard
'@returns {stdole.IPicture}   - Picture object representing clipboard data
Private Function GetPictureFromClipboard(Optional ByVal lPicType As CLIPFORMAT) As stdole.IPicture
    'Define pointers
    #If VBA7 Then
      Dim hPtr As LongPtr, hCopy As LongPtr
    #Else
      Dim hPtr As Long, hCopy As Long
    #End If
    
    'Check if the clipboard contains the required format
    If IsClipboardFormatAvailable(lPicType) Then
      'Get access to the clipboard
      If OpenClipboard(0&) > 0 Then
        'Get a handle to the image data
        hPtr = GetClipboardData(lPicType)
        If Not CBool(hPtr) Then Call CriticalRaise("GetPictureFromClipboard [PRIVATE]", "Unable to get clipboard data handle")

        'Create our own copy of the image on the clipboard, in the appropriate format.
        If lPicType = CLIPFORMAT.CF_BITMAP Then
            hCopy = CopyImage(hPtr, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG)
        Else
            hCopy = CopyEnhMetaFile(hPtr, vbNullString)
        End If

        'Release the clipboard to other programs
        If Not CBool(CloseClipboard()) Then Call CriticalRaise("GetPictureFromClipboard [PRIVATE]", "Unable to close clipboard")

        'If we got a handle to the image, convert it into a Picture object and return it
        If hCopy Then
          Set GetPictureFromClipboard = CreatePicture(hCopy, 0, lPicType)
        Else
          Call CriticalRaise("GetPictureFromClipboard [PRIVATE]", "Unable to obtain image handle")
        End If
      Else
        Call CriticalRaise("GetPictureFromClipboard [PRIVATE]", "Unable to open clipboard")
      End If
    Else
      Call CriticalRaise("GetPictureFromClipboard [PRIVATE]", "Clipboard format is not available")
    End If
End Function

'Obtains an stdole.IPicture object from the handle supplied. This is used in conjunction with GetPictureFromClipboard()
'@param {ByVal LongPtr/Long} - The HBITMAP handle identifying the bitmap assigned to the picture object.
'@param {ByVal LongPtr/Long} - The HPALETTE handle identifying the color palette for the bitmap. Can be created with CreatePallette(LOGPALETTE) function from Gdi32.dll
'@param {Byval CLIPFORMAT}   - The format of the picture stored in the clipboard
'@returns {stdole.IPicture}     - Picture object representing clipboard data
#If VBA7 Then
Private Function CreatePicture(ByVal hPic As LongPtr, ByVal hPal As LongPtr, ByVal lPicType As CLIPFORMAT) As stdole.IPicture
#Else
Private Function CreatePicture(ByVal hPic As Long, ByVal hPal As Long, ByVal lPicType As CLIPFORMAT) As stdole.IPicture
#End If
  ' stdole.IPicture requires a reference to "OLE Automation"
  Dim r As Long, uPicInfo As PICTDESC, iid As GUID, IPic As stdole.IPicture
  
  ' OLE Picture types
  Const PICTYPE_BITMAP = 1
  Const PICTYPE_METAFILE = 2
  Const PICTYPE_ICON = 3
  Const PICTYPE_ENHMETAFILE = 4

  ' Create the Interface GUID (for the stdole.IPicture interface)
  If Not CBool(IIDFromString(strPtr(IID_DISPATCH), iid)) Then Call CriticalRaise("CreatePicture", "Cannot get IDispatch IID")

  ' Fill uPicInfo with necessary parts.
  With uPicInfo
      .size = Len(uPicInfo)                                                              ' Length of structure.
      .Type = IIf(lPicType = CLIPFORMAT.CF_BITMAP, PICTYPE_BITMAP, PICTYPE_ENHMETAFILE)  ' Type of Picture
      .hPic = hPic                                                                       ' Handle to image.
      .hPal = IIf(lPicType = CLIPFORMAT.CF_BITMAP, hPal, 0)                              ' Handle to palette (if bitmap).
  End With

  ' Create the Picture object.
  If CBool(OleCreatePictureIndirect(uPicInfo, iid, True, IPic)) Then
    Set CreatePicture = IPic
  Else
    Call CriticalRaise("CreatePicture [PRIVATE]", "An error occurred in OleCreatePictureIndirect. Error code (" & r & ": " & fnOLEError(r) & ")")
  End If
End Function

'Given an OLE Error ID, obtain a description of what this ID means:
'@param {ByVal Long} - The error ID obtained from OLE
'@returns {String} - Description of what the id means
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
'@param {ByRef Variant} dest Destination to copy variant to
'@param {Variant} value Source to copy variant from.
'@perf This appears to be a faster variant of "oleaut32.dll\VariantCopy" + it's multi-platform
Private Sub CopyVariant(ByRef dest As Variant, ByVal value As Variant)
  If IsObject(value) Then
    Set dest = value
  Else
    dest = value
  End If
End Sub
