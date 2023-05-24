



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
'<#stdPicture>.CreateFromSVG()?
Public Enum stdPicFormat
    stdPicFormatDefault
    stdPicFormatBMP
    stdPicFormatPNG
    stdPicFormatGIF
    stdPicFormatJPEG
    stdPicFormatTIFF
End Enum

'Clipboard Management
#if VBA7 then
    Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Integer) As Long
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (Optional ByVal hwnd As LongPtr = 0^) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Integer) As LongPtr

    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr

    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As Long)
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long

    Private Declare PtrSafe Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As LongPtr, ByVal offsetinVft As LongPtr, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As VbVarType, ByRef paValues As LongPtr, ByRef retVAR As Variant) As Long
#else
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

    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As Long)
    Private Declare Function GetTickCount Lib "kernel32" () As Long

    'FIX: The use of `VbVarType` for the type of `paTypes` on 32 bit causes Invalid Callee error. Workaround is to use `Integer` instead as below.
    Private Declare Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
#end if
Private Const CF_BITMAP = 2

'****************
'* Private Vars *
'****************
Private pImageFile as Object
Private pCurrentFormat as stdPicFormat

'****************
'* Constructors *
'****************
Public Function CreateFromBinary(ByRef b() as Byte) as stdPicture
    Set CreateFromBinary = new stdPicture
    With CreateObject("WIA.Vector")
        .binaryData = b
        Call CreateFromBinary.protInit(.ImageFile)
    end with
End Function

Public Function CreateFromFile(ByVal sPath as string) as stdPicture
    set CreateFromFile = new stdPicture
    set wia = CreateObject("WIA.ImageFile")
    call wia.LoadFile(sPath)
    call CreateFromFile.protInit(wia)
End Function

Public Function CreateFromUrl(ByVal sURL as string) as stdPicture
    set CreateFromUrl = new stdPicture
    Dim HTTP As Object: Set HTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    HTTP.Open "GET", TargetURL, False
    'TODO: AutoProxy? 
    HTTP.send
    If HTTP.Status = 200 Then
        set CreateFromURL = CreateFromBinary(HTTP.responseBody)
    else
        Err.Raise 1, "", "Could not get file (HTTP Status: " & http.status & ")"
    End If
End Function

Public Function CreateFromClipboard() as stdPicture
    set CreateFromClipboard = CreateFromBinary(getClipboardBytes())
End Function

'@param {Object<IPictureDisp>}
Public Function CreateFromIPictureDisp(ByVal picture as IUnkown) as stdPicture
    set CreateFromIPictureDisp = new stdPicture

    'Create ADODB Stream
    Const adTypeBinary as long = 1
    Dim oStream as Object: set oStream = CreateObject("ADODB.Stream")
    oStream.Type = adTypeBinary
    oStream.Open

    'Obtain IStream object
    Dim oIStream as IUnknown: set oIStream = unkQueryInterface(oStream, "0000000c-0000-0000-C000-000000000046")

    'Save picture to stream
    Call picture.saveAsFile(oIStream, True, 0)

    'Create picture from stream bytes
    If oStream.size = 0 then Exit Function
    oStream.Position = 0
    set CreateFromIPictureDisp = CreateFromBinary(oStream.Read)
End Function

Public Function CreateFromShape(Byval shp as object) as stdPicture
    select case Application.Name
        case "Microsoft Excel"
            '...
    end select
End Function

Public Function CreateFromChart(Byval chart as object) as stdPicture
    select case Application.Name
        case "Microsoft Excel"
            '...
    end select
End Function

'@param {Object<WIA.ImageFile>}
Public Sub protInit(ByVal imageFile as object)
    set pImageFile = imageFile
End Sub

Public Property Get binary(Optional ByVal eFormat as stdPicFormat = stdPicFormatDefault) as Byte()
    'If default, leave in current format
    if eFormat = stdPicFormatDefault then
        Data = pImageFile.FileData.BinaryData
    else
        'Obtain required format
        Dim sFormatUID as string
        select case eFormat
            case stdPicFormatBMP:  sFormatUID = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
            case stdPicFormatPNG:  sFormatUID = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
            case stdPicFormatGIF:  sFormatUID = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}"
            case stdPicFormatJPEG: sFormatUID = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
            case stdPicFormatTIFF: sFormatUID = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}"
        end select

        With CreateObject("WIA.ImageProcess")
            .Filters.Add .FilterInfos("Convert").FilterID
            .Filters(.Filters.Count).Properties("FormatID").Value = wiaFormatPNG
            Data = .Apply(pImageFile).FileData.BinaryData
        End With
    end if
End Property

Public Property Let binary(ByVal eFormat as stdPicFormat = stdPicFormatDefault, Byval b Byte())

End Property

Public Sub saveAs(ByVal sSaveAsFile as string, optional ByVal eFormat as stdPicFormat = stdPicFormatDefault)
    Dim ff as Long: ff = FreeFile()
    Open File sSaveAsFile for Output as #ff
        Print #ff, Binary(eFormat)
    Close #ff
End Sub

Public Sub copyToClipboard()

End Sub


















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
Private Function unkQueryInterface(ByVal obj as IUnknown, ByVal sIID as string) as IUnknown
    Dim GUID() as Byte: GUID = IIDFromString(sIID)
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
Public Function CallVT(ByVal pObjPtr as LongPtr, ByVal VTableOffset As Long, ByVal ReturnType As VbVarType, ParamArray FunctionParameters() As Variant) As Variant
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
        vType(lIdx) = varType(vParam(lIdx))
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


'Obtain the clipboard as a byte array
'@returns {Byte()} Byte array of clipboard data
Private Function getClipboardBytes() as Byte()
  #If VBA7 Then
    Dim hClipMemory     As LongPtr
    Dim lpClipMemory    As LongPtr
  #Else
    Dim hClipMemory     As Long
    Dim lpClipMemory    As Long
  #End If
  
  'Get pointer to clipboard data
  If Not CBool(OpenClipboardTimeout(Application.hwnd)) Then Call CriticalRaise("Value [GET]", "Unable to open clipboard (DllError: " & Err.LastDllError & ")")
  If Not CBool(IsClipboardFormatAvailable(CF_BITMAP)) Then Call CriticalRaise("Value [GET]", "Specified clipboard format is not available")
  hClipMemory = GetClipboardData(CF_BITMAP)
  If Not CBool(hClipMemory) Then Call CriticalRaise("Value [GET]", "Unable to obtain pointer to clipboard data (DllError: " & Err.LastDllError & ")")
  
  'Get the data from the supplied pointer
  Dim size As Long
  size = GlobalSize(hClipMemory)
  lpClipMemory = GlobalLock(hClipMemory)
  If CBool(lpClipMemory) Then
    getClipboardBytes = BytesFromPointer(lpClipMemory, size)
  Else
    Call CriticalRaise("Value [GET]", "Unable to lock global memory (DllError: " & Err.LastDllError & ")")
  End If
  If CBool(GlobalUnlock(hClipMemory)) Then If Err.LastDllError <> 0 Then Call CriticalRaise("Value [GET]", "Unable to unlock global mem (DllError: " & Err.LastDllError & ")")
  If Not CBool(CloseClipboard()) Then Call CriticalRaise("Value [GET]", "Unable to close clipboard (DllError: " & Err.LastDllError & ")")
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

'Tries to open the clipbaord with a timeout
'@param {long} hWnd to open clipboard in
'@param {long} number of milliseconds to wait
'@returns {boolean} true if clipboard is open, false if clipboard couldn't be opened
Private Function OpenClipboardTimeout(ByVal hwnd As Long, Optional ByVal dwTimeoutMilliseconds As Long = 5000) As Boolean
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