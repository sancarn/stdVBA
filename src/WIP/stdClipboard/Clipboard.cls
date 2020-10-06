VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Clipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Src: https://codekabinett.com/rdumps.php?Lang=2&targetDoc=vba-clipboard-file-content
Option Compare Database
Option Explicit

Private Const CLASS_NAME As String = "Clipboard"

Private Declare Function OpenClipboard Lib "User32" (ByVal hwnd As LongPtr) As Long
Private Declare Function CloseClipboard Lib "User32" () As Long
Private Declare Function GetClipBoardData Lib "User32" Alias "GetClipboardData" (ByVal wFormat As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags&, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As Long

Private Const CF_TEXT As Long = 1
Private Const CF_UNICODETEXT As Long = 13

Private Const CFSTR_FILECONTENTS  As String = "FileContents"

Private Declare Function EnumClipboardFormats Lib "User32" (ByVal Format As Long) As Long
Private Declare Function GetClipboardFormatNameW Lib "User32" (ByVal Format As Long, ByVal lpszFormatName As LongPtr, ByVal cchMaxCount As Long) As Long
Private Declare Function IsClipboardFormatAvailable Lib "User32" (ByVal Format As Long) As Long
Private Declare Function RegisterClipboardFormatW Lib "User32" (ByVal lpszFormat As LongPtr) As Long

Private Declare PtrSafe Function lstrlenA Lib "kernel32.dll" (ByVal lpString As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As Long)
Private Declare PtrSafe Function lstrlenW Lib "kernel32.dll" (ByVal lpString As LongPtr) As Long


Private Const ERR_OPEN_FAILED As Long = vbObjectError + 1
Private Const ERR_GETHANDLE As Long = vbObjectError + 2
Private Const ERR_LOCKMEMORY As Long = vbObjectError + 3
Private Const ERR_CONTENTTYPE As Long = vbObjectError + 4
Private Const ERR_INVALID_FORMAT As Long = vbObjectError + 5


Public Property Get FormatText() As Long
    FormatText = CF_TEXT
End Property
Public Property Get FormatUnicodeText() As Long
    FormatUnicodeText = CF_UNICODETEXT
End Property
Public Property Get FormatFileContent() As String
    FormatFileContent = CFSTR_FILECONTENTS
End Property

Private Function FormatFromName(ByVal formatName As String) As Long
    FormatFromName = RegisterClipboardFormatW(StrPtr(formatName))
End Function

Private Function StringFromPointerW(ByVal pointerToString As LongPtr) As String
' Source and Info: https://codekabinett.com/rdumps.php?Lang=2&targetDoc=api-pointer-convert-vba-string-ansi-unicode

    Const BYTES_PER_CHAR As Integer = 2

    Dim tmpBuffer()    As Byte
    Dim ByteCount      As Long
 
    ' determine size of source string in bytes
    ByteCount = lstrlenW(pointerToString) * BYTES_PER_CHAR
    
    If ByteCount > 0 Then
        ' Resize the buffer as required
        ReDim tmpBuffer(0 To ByteCount - 1) As Byte
        
        ' Copy the bytes from pointerToString to tmpBuffer
        Call CopyMemory(VarPtr(tmpBuffer(0)), pointerToString, ByteCount)
    End If
 
    ' Straigth assigment Byte() to String possible - Both are Unicode!
    StringFromPointerW = tmpBuffer

End Function

Private Function StringFromPointerA(ByVal pointerToString As LongPtr) As String
' Source and Info: https://codekabinett.com/rdumps.php?Lang=2&targetDoc=api-pointer-convert-vba-string-ansi-unicode

    Dim tmpBuffer()    As Byte
    Dim ByteCount      As Long
    Dim retVal         As String
 
    ' determine size of source string in bytes
    ByteCount = lstrlenA(pointerToString)
    
    If ByteCount > 0 Then
        ' Resize the buffer as required
        ReDim tmpBuffer(0 To ByteCount - 1) As Byte
        
        ' Copy the bytes from pointerToString to tmpBuffer
        Call CopyMemory(VarPtr(tmpBuffer(0)), pointerToString, ByteCount)
    End If
 
    ' Convert (ANSI) buffer to VBA string
    retVal = StrConv(tmpBuffer, vbUnicode)
    
    StringFromPointerA = retVal

End Function

Private Function ByteArrayFromPointer(ByVal pointer As LongPtr, ByVal ByteCount As Long) As Byte()

    Dim tmpBuffer()    As Byte
    Dim retVal         As String
 
    If ByteCount > 0 Then
        ' Resize the buffer as required
        ReDim tmpBuffer(0 To ByteCount - 1) As Byte
        
        ' Copy the bytes from pointerToString to tmpBuffer
        Call CopyMemory(VarPtr(tmpBuffer(0)), pointer, ByteCount)
    End If
 
    ByteArrayFromPointer = tmpBuffer

End Function

Public Function GetAvailableClipBoardFormats() As String

    Dim retVal As String
    Dim currentFormat As Long
    Dim apiRetVal As Long
    Dim buffer As String
    Const bufferSize As Long = 200
  
On Error GoTo GetClipBoardFormats_Exit
    OpenClipboard 0
    
    Do
        currentFormat = EnumClipboardFormats(currentFormat)
        
        buffer = Space(bufferSize)
        apiRetVal = GetClipboardFormatNameW(currentFormat, StrPtr(buffer), bufferSize)
        If currentFormat <> 0 Then
            retVal = retVal & vbCrLf & currentFormat & ": " & Trim(buffer)
        End If
    Loop Until currentFormat = 0
    
    GetAvailableClipBoardFormats = retVal
    
GetClipBoardFormats_Exit:
    CloseClipboard
    
End Function

Private Function getFormat(ByVal varFormat As Variant) As Long
    Dim dataFormat      As Long
     
     Select Case True
        Case VarType(varFormat) = vbLong
            dataFormat = CLng(varFormat)
        Case VarType(varFormat) = vbString
            dataFormat = FormatFromName(CStr(varFormat))
        Case Else
            Err.Raise ERR_INVALID_FORMAT, CLASS_NAME, "Invalid Format"
    End Select
    
    getFormat = dataFormat
    
End Function

Public Function IsFormatAvailable(ByVal varFormat As Variant) As Boolean
    OpenClipboard 0
    IsFormatAvailable = CBool(IsClipboardFormatAvailable(getFormat(varFormat)))
    CloseClipboard
End Function

Public Function GetData(ByVal varFormat As Variant) As Variant

    Dim hClipMemory     As LongPtr
    Dim lpClipMemory    As LongPtr
    Dim memSize         As Long
    Dim retVal          As Variant
    Dim apiRetVal       As Long
    Dim dataFormat      As Long
    
    dataFormat = getFormat(varFormat)

    If Not CBool(OpenClipboard(0)) Then
        FailWithError ERR_OPEN_FAILED
    End If
    
    If Not CBool(IsClipboardFormatAvailable(dataFormat)) Then
        FailWithError ERR_CONTENTTYPE
    End If
          
    hClipMemory = GetClipBoardData(dataFormat)
    If Not CBool(hClipMemory) Then
        FailWithError ERR_GETHANDLE
    End If
   
    memSize = GlobalSize(hClipMemory)
    lpClipMemory = GlobalLock(hClipMemory)
 
    If CBool(lpClipMemory) Then
        Select Case True
            Case dataFormat = CF_UNICODETEXT
                retVal = StringFromPointerW(lpClipMemory)
            Case dataFormat = CF_TEXT
                retVal = StringFromPointerA(lpClipMemory)
            Case Else
                retVal = ByteArrayFromPointer(lpClipMemory, memSize)
        End Select
        apiRetVal = GlobalUnlock(hClipMemory)
    Else
        FailWithError ERR_LOCKMEMORY
    End If
 
    GetData = retVal
    CloseClipboard
 
End Function


Private Sub FailWithError(ByVal ErrorNumber As Long)

    CloseClipboard
    
    Select Case ErrorNumber
        Case ERR_OPEN_FAILED
            Err.Raise ERR_OPEN_FAILED, CLASS_NAME, "Failed to open clipboard"
        Case ERR_CONTENTTYPE
            Err.Raise ERR_CONTENTTYPE, CLASS_NAME, "Content not in clipboard"
        Case ERR_GETHANDLE
            Err.Raise ERR_GETHANDLE, CLASS_NAME, "Failed to get clipboard memory handle"
        Case ERR_LOCKMEMORY
            Err.Raise ERR_LOCKMEMORY, CLASS_NAME, "Failed to lock clipboard memory "
    End Select
    
End Sub
 
