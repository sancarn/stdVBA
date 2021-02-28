'SRC: https://social.msdn.microsoft.com/Forums/en-US/e624729a-e8bd-4d16-867f-6bd48000bbaa/copy-file-into-clipboard?forum=isvvba

Option Explicit

' Required data structures
Private Type POINTAPI
x As Long
y As Long
End Type

' Clipboard Manager Functions
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long

' Other required Win32 APIs
Private Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Private Declare Function DragQueryPoint Lib "shell32.dll" (ByVal hDrop As Long, lpPoint As POINTAPI) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' Predefined Clipboard Formats
Private Const CF_TEXT = 1
Private Const CF_BITMAP = 2
Private Const CF_METAFILEPICT = 3
Private Const CF_SYLK = 4
Private Const CF_DIF = 5
Private Const CF_TIFF = 6
Private Const CF_OEMTEXT = 7
Private Const CF_DIB = 8
Private Const CF_PALETTE = 9
Private Const CF_PENDATA = 10
Private Const CF_RIFF = 11
Private Const CF_WAVE = 12
Private Const CF_UNICODETEXT = 13
Private Const CF_ENHMETAFILE = 14
Private Const CF_HDROP = 15
Private Const CF_LOCALE = 16
Private Const CF_MAX = 17

' New shell-oriented clipboard formats
Private Const CFSTR_SHELLIDLIST As String = "Shell IDList Array"
Private Const CFSTR_SHELLIDLISTOFFSET As String = "Shell Object Offsets"
Private Const CFSTR_NETRESOURCES As String = "Net Resource"
Private Const CFSTR_FILEDESCRIPTOR As String = "FileGroupDescriptor"
Private Const CFSTR_FILECONTENTS As String = "FileContents"
Private Const CFSTR_FILENAME As String = "FileName"
Private Const CFSTR_PRINTERGROUP As String = "PrinterFriendlyName"
Private Const CFSTR_FILENAMEMAP As String = "FileNameMap"

' Global Memory Flags
Private Const GMEM_FIXED = &H0
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_NOCOMPACT = &H10
Private Const GMEM_NODISCARD = &H20
Private Const GMEM_ZEROINIT = &H40
Private Const GMEM_MODIFY = &H80
Private Const GMEM_DISCARDABLE = &H100
Private Const GMEM_NOT_BANKED = &H1000
Private Const GMEM_SHARE = &H2000
Private Const GMEM_DDESHARE = &H2000
Private Const GMEM_NOTIFY = &H4000
Private Const GMEM_LOWER = GMEM_NOT_BANKED
Private Const GMEM_VALID_FLAGS = &H7F72
Private Const GMEM_INVALID_HANDLE = &H8000
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Private Type DROPFILES
    pFiles As Long
    pt As POINTAPI
    fNC As Long
    fWide As Long
End Type

Public Function ClipboardCopyFiles(Files() As String) As Boolean
    Dim data As String
    Dim df As DROPFILES
    Dim hClipMemory As Long
    Dim lpClipMemory As Long
    Dim i As Long

    ' Open and clear existing crud off clipboard.
    If OpenClipboard(0&) Then
        Call EmptyClipboard

        ' Build double-null terminated list of files.
        For i = LBound(Files) To UBound(Files)
            data = data & Files(i) & vbNullChar
        Next
        data = data & vbNullChar

        ' Allocate and get pointer to global memory,
        ' then copy file list to it.
        hClipMemory = GlobalAlloc(GHND, Len(df) + Len(data))
        If hClipMemory Then
            lpClipMemory = GlobalLock(hClipMemory)

            ' Build DROPFILES structure in global memory.
            df.pFiles = Len(df)
            Call CopyMem(ByVal lpClipMemory, df, Len(df))
            Call CopyMem(ByVal (lpClipMemory + Len(df)), ByVal data, Len(data))
            Call GlobalUnlock(hClipMemory)

            ' Copy data to clipboard, and return success.
            If SetClipboardData(CF_HDROP, hClipMemory) Then
                ClipboardCopyFiles = True
            End If
        End If

        ' Clean up
        Call CloseClipboard
    End If
End Function

Public Function ClipboardPasteFiles(Files() As String) As Long
    Dim hDrop As Long
    Dim nFiles As Long
    Dim i As Long
    Dim desc As String
    Dim filename As String
    Dim pt As POINTAPI
    Const MAX_PATH As Long = 260

    ' Insure desired format is there, and open clipboard.
    If IsClipboardFormatAvailable(CF_HDROP) Then
        If OpenClipboard(0&) Then
            ' Get handle to Dropped Filelist data, and number of files.
            hDrop = GetClipboardData(CF_HDROP)
            nFiles = DragQueryFile(hDrop, -1&, "", 0)

            ' Allocate space for return and working variables.
            ReDim Files(0 To nFiles - 1) As String
            filename = Space(MAX_PATH)

            ' Retrieve each filename in Dropped Filelist.
            For i = 0 To nFiles - 1
                Call DragQueryFile(hDrop, i, filename, Len(filename))
                Files(i) = TrimNull(filename)
            Next

            ' Clean up
            Call CloseClipboard
        End If

    ' Assign return value equal to number of files dropped.
    ClipboardPasteFiles = nFiles
    End If

End Function

Private Function TrimNull(ByVal sTmp As String) As String
    Dim nNul As Long

    '
    ' Truncate input sTmpg at first Null.
    ' If no Nulls, perform ordinary Trim.
    '
    nNul = InStr(sTmp, vbNullChar)
    Select Case nNul
        Case Is > 1
            TrimNull = Left(sTmp, nNul - 1)
        Case 1
            TrimNull = ""
        Case 0
            TrimNull = Trim(sTmp)
    End Select
End Function

Sub maaa()

    'i = "c:\" & ActiveDocument.Name
    'ActiveDocument.SaveAs i
    Dim afile(0) As String

    afile(0) = "c:\070206.excel" 'The file actually exists
    MsgBox ClipboardCopyFiles(afile)

End Sub

