VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JSON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute Value.VB_UserMemId = 0


'CREDITS:
'----------
'Initial pictures implementation by Stephen Bullen (Office Automation Ltd) on 30/10/1998
'----------



'API DEFINITIONS:
'TODO: Make declarations VBA7 compatible
Private Declare Function EmptyClipboard Lib "user32.dll" () As Long
Private Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long

#If VBA7 Then
  Private Type uPicDesc
      Size As Long
      Type As Long
      hPic As LongPtr
      hPal As LongPtr
  End Type

  Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Integer) As Long
  Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As Long
  Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Integer) As LongPtr
  Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
  Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (PicDesc As uPicDesc, RefIID As GUID, ByVal fPictureOwnsHandle As LongPtr, IPic As IPicture) As Long
  Private Declare PtrSafe Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As LongPtr, ByVal lpszFile As String) As LongPtr
  Private Declare PtrSafe Function CopyImage Lib "user32" (ByVal Handle As LongPtr, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As LongPtr
#Else
  Private Type uPicDesc
      Size As Long
      Type As Long
      hPic As Long
      hPal As Long
  End Type

  Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Integer) As Long
  Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
  Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Integer) As Long
  Private Declare Function CloseClipboard Lib "user32" () As Long
  Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As uPicDesc, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
  Declare Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As Long, ByVal lpszFile As String) As Long
  Declare Function CopyImage Lib "user32" (ByVal Handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
#End If

'TODO: Would be better to use stdCOM here (cast to bytes[])
'Declare a UDT to store a GUID for the IPicture OLE Interface
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

'The API format types we're interested in
Const IMAGE_BITMAP = 0
Const LR_COPYRETURNORG = &H4

Private Enum CLIPFORMAT
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






'TODO: REQUIRES MAIN INTERFACE IMPLEMENTATION

Private pValue as Variant
Private pType as String

Public Property Get Value() as Variant
  
End Property
Public Property Let Value(v as variant)

End Property

Public Property Get Photo() as IPhoto
  if pType = "CF_BITMAP" then
    
  elseif pType = "CF_METAFILEPIC" then

  else
    Err.Raise(1,"stdClipboard::Photo","Invalid clipboard format")
  end if
End Property

Public Property Set Photo(v as IPhoto)
  'Set format based on passed photo v
End Property

Public Function GetBytes() as Byte()
  
End Function








'Picture helpers
'TODO: REQUIRES REFACTORING!!



'Save a picture from a shape
'TODO: This needs a new home
Private Sub VbSavePicture(ByVal shp As Shape, ByVal file As String)
  Call shp.CopyPicture(Appearance:=xlScreen, format:=xlBitmap)
  StdOle.SavePicture PastePicture(xlBitmap), file
End Sub

Private Function PastePicture(Optional lXlPicType As Long = xlPicture) As IPicture
  #If VBA7 Then
    Set PastePicture = PastePictureVBA7(lXlPicType)
  #Else
    Set PastePicture = PastePictureWin32(lXlPicType)
  #End If
End Function

Private Function PastePictureVBA7(Optional lXlPicType As Long = xlPicture) As IPicture
    'Some pointers
    Dim H As Long, hPicAvail As Long, hPtr As LongPtr, hPal As LongPtr, lPicType As Long, hCopy As LongPtr
    
    'Convert the type of picture requested from the xl constant to the API constant
    lPicType = IIf(lXlPicType = xlBitmap, CLIPFORMAT.CF_BITMAP, CLIPFORMAT.CF_ENHMETAFILE)
    
    'Check if the clipboard contains the required format
    hPicAvail = IsClipboardFormatAvailable(lPicType)
    
    If hPicAvail <> 0 Then
        'Get access to the clipboard
        H = OpenClipboard(0&)
    
        If H > 0 Then
            'Get a handle to the image data
            hPtr = GetClipboardData(lPicType)
    
            'Create our own copy of the image on the clipboard, in the appropriate format.
            If lPicType = CLIPFORMAT.CF_BITMAP Then
                hCopy = CopyImage(hPtr, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG)
            Else
                hCopy = CopyEnhMetaFile(hPtr, vbNullString)
            End If
    
            'Release the clipboard to other programs
            H = CloseClipboard
    
            'If we got a handle to the image, convert it into a Picture object and return it
            If hPtr <> 0 Then Set PastePictureVBA7 = CreatePictureVBA7(hCopy, 0, lPicType)
        End If
    End If
End Function


Private Function PastePictureWin32(Optional lXlPicType As Long = xlPicture) As IPicture
    'Some pointers
    Dim H As Long, hPicAvail As Long, hPtr As Long, hPal As Long, lPicType As Long, hCopy As Long
    
    'Convert the type of picture requested from the xl constant to the API constant
    lPicType = IIf(lXlPicType = xlBitmap, CLIPFORMAT.CF_BITMAP, CLIPFORMAT.CF_ENHMETAFILE)
    
    'Check if the clipboard contains the required format
    hPicAvail = IsClipboardFormatAvailable(lPicType)
    
    If hPicAvail <> 0 Then
        'Get access to the clipboard
        H = OpenClipboard(0&)
    
        If H > 0 Then
            'Get a handle to the image data
            hPtr = GetClipboardData(lPicType)
    
            'Create our own copy of the image on the clipboard, in the appropriate format.
            If lPicType = CLIPFORMAT.CF_BITMAP Then
                hCopy = CopyImage(hPtr, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG)
            Else
                hCopy = CopyEnhMetaFile(hPtr, vbNullString)
            End If
    
            'Release the clipboard to other programs
            H = CloseClipboard
    
            'If we got a handle to the image, convert it into a Picture object and return it
            If hPtr <> 0 Then Set PastePictureWin32 = CreatePicture(hCopy, 0, lPicType)
        End If
    End If
End Function



Private Function CreatePictureVBA7(ByVal hPic As LongPtr, ByVal hPal As LongPtr, ByVal lPicType) As IPicture
  ' IPicture requires a reference to "OLE Automation"
  Dim r As Long, uPicInfo As uPicDesc, IID_IDispatch As GUID, IPic As IPicture
  
  'OLE Picture types
  Const PICTYPE_BITMAP = 1
  Const PICTYPE_ENHMETAFILE = 4
  
  ' Create the Interface GUID (for the IPicture interface)
  With IID_IDispatch
      .Data1 = &H7BF80980
      .Data2 = &HBF32
      .Data3 = &H101A
      .Data4(0) = &H8B
      .Data4(1) = &HBB
      .Data4(2) = &H0
      .Data4(3) = &HAA
      .Data4(4) = &H0
      .Data4(5) = &H30
      .Data4(6) = &HC
      .Data4(7) = &HAB
  End With
  
  ' Fill uPicInfo with necessary parts.
  With uPicInfo
      .Size = Len(uPicInfo)                                                   ' Length of structure.
      .Type = IIf(lPicType = CLIPFORMAT.CF_BITMAP, PICTYPE_BITMAP, PICTYPE_ENHMETAFILE)  ' Type of Picture
      .hPic = hPic                                                            ' Handle to image.
      .hPal = IIf(lPicType = CLIPFORMAT.CF_BITMAP, hPal, 0)                              ' Handle to palette (if bitmap).
  End With
       
      ' Create the Picture object.
  r = OleCreatePictureIndirect(uPicInfo, IID_IDispatch, 1, IPic)
  
  
  ' If an error occured, show the description
  If r <> 0 Then Debug.Print "Create Picture: " & fnOLEError(r)
  
  ' Return the new Picture object.
  Set CreatePictureVBA7 = IPic
End Function


Private Function CreatePicture(ByVal hPic As Long, ByVal hPal As Long, ByVal lPicType) As IPicture
  ' IPicture requires a reference to "OLE Automation"
  Dim r As Long, uPicInfo As uPicDesc, IID_IDispatch As GUID, IPic As IPicture

  'OLE Picture types
  Const PICTYPE_BITMAP = 1
  Const PICTYPE_ENHMETAFILE = 4

  ' Create the Interface GUID (for the IPicture interface)
  With IID_IDispatch
      .Data1 = &H7BF80980
      .Data2 = &HBF32
      .Data3 = &H101A
      .Data4(0) = &H8B
      .Data4(1) = &HBB
      .Data4(2) = &H0
      .Data4(3) = &HAA
      .Data4(4) = &H0
      .Data4(5) = &H30
      .Data4(6) = &HC
      .Data4(7) = &HAB
  End With

  ' Fill uPicInfo with necessary parts.
  With uPicInfo
      .Size = Len(uPicInfo)                                                   ' Length of structure.
      .Type = IIf(lPicType = CLIPFORMAT.CF_BITMAP, PICTYPE_BITMAP, PICTYPE_ENHMETAFILE)  ' Type of Picture
      .hPic = hPic                                                            ' Handle to image.
      .hPal = IIf(lPicType = CLIPFORMAT.CF_BITMAP, hPal, 0)                              ' Handle to palette (if bitmap).
  End With

  ' Create the Picture object.
  r = OleCreatePictureIndirect(uPicInfo, IID_IDispatch, True, IPic)

  ' If an error occured, show the description
  If r <> 0 Then Debug.Print "Create Picture: " & fnOLEError(r)

  ' Return the new Picture object.
  Set CreatePicture = IPic
End Function

Private Function fnOLEError(lErrNum As Long) As String
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