



'Mostly Wraps GDI++ API
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













Option Explicit

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type
Private Type EncoderParameter
    GUID As GUID
    NumberOfValues As Long
    type As Long
    Value As Long
End Type
Private Type EncoderParameters
    count As Long
    Parameter As EncoderParameter
End Type


Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As Long, ByVal hPal As Long, BITMAP As Long) As Long
Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "GDIPlus" (ByVal Image As Long, ByVal FileName As Long, clsidEncoder As GUID, encoderParams As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal Str As Long, id As GUID) As Long
Private Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb As Long) As Long

Private pPicture as StdOle.StdPicture

Public Event PictureBeforeChanged(OldPicture as StdOle.StdPicture, NewPicture as StdOle.StdPicture)
Public Event PictureAfterChanged(NewPicture as StdOle.StdPicture)

Public Property Get Picture() as StdOle.StdPicture
  set Picture = pPicture
End Sub
Public Property Set Picture(pic as StdOle.StdPicture) 
  RaiseEvent PictureBeforeChanged(pPicture,pic)
  set pPicture = pic
  RaiseEvent PictureAfterChanged(pPicture)
End Sub

Public Sub ToClipboard()
  if typename(Clipboard) = vbObject then
    Clipboard.SetData pPicture
  else
    'TODO: Think this method is very hacky.
    'TODO: A less hacky method can be found here:  http://access.mvps.org/access/api/api0042.htm
    With ActiveSheet.OLEObjects.Add(ClassType:="Forms.Image.1", Link:=False, DisplayAsIcon:=False, Left:=330, Top:=215)
      .Activate
      .Picture = pPicture
      .AutoSize = true
      .CopyPicture
      .Delete
    End with
  end if
End Sub

Public Sub ToFile(ByVal FileName As String, Optional PicType As String="INFER", Optional ByVal Quality As Byte = 80, Optional ByVal TIFF_ColorDepth As Long = 24, Optional ByVal TIFF_Compression As Long = 6)
    Screen.MousePointer = vbHourglass
    Dim tSI As GdiplusStartupInput
    Dim lRes As Long
    Dim lGDIP As Long
    Dim lBitmap As Long
    Dim aEncParams() As Byte
    On Error GoTo ErrHandle:
    tSI.GdiplusVersion = 1
    lRes = GdiplusStartup(lGDIP, tSI)
    If lRes = 0 Then
        lRes = GdipCreateBitmapFromHBITMAP(pPicture.Handle, 0, lBitmap)
        If lRes = 0 Then
            Dim tJpgEncoder As GUID
            Dim tParams As EncoderParameters
            Select Case PicType
            Case ".jpg"
                CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
                tParams.count = 1
                With tParams.Parameter
                    CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID
                    .NumberOfValues = 1
                    .type = 4
                    .Value = VarPtr(Quality)
                End With
                ReDim aEncParams(1 To Len(tParams))
                Call CopyMemory(aEncParams(1), tParams, Len(tParams))
            Case ".png"
                CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
                ReDim aEncParams(1 To Len(tParams))
            Case ".gif"
                CLSIDFromString StrPtr("{557CF402-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
                ReDim aEncParams(1 To Len(tParams))
            Case ".tiff"
                CLSIDFromString StrPtr("{557CF405-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
                tParams.count = 2
                ReDim aEncParams(1 To Len(tParams) + Len(tParams.Parameter))
                With tParams.Parameter
                    .NumberOfValues = 1
                    .type = 4
                    CLSIDFromString StrPtr("{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"), .GUID
                    .Value = VarPtr(TIFF_Compression)
                End With
                Call CopyMemory(aEncParams(1), tParams, Len(tParams))
                With tParams.Parameter
                    .NumberOfValues = 1
                    .type = 4
                    CLSIDFromString StrPtr("{66087055-AD66-4C7C-9A18-38A2310B8337}"), .GUID
                    .Value = VarPtr(TIFF_ColorDepth)
                End With
                Call CopyMemory(aEncParams(Len(tParams) + 1), tParams.Parameter, Len(tParams.Parameter))
            Case ".bmp"
                SavePicture pPicture, FileName
                Screen.MousePointer = vbDefault
                Exit Sub
            End Select
            lRes = GdipSaveImageToFile(lBitmap, StrPtr(FileName), tJpgEncoder, aEncParams(1))
            GdipDisposeImage lBitmap
        End If
        GdiplusShutdown lGDIP
    End If
    Screen.MousePointer = vbDefault
    Erase aEncParams
    Exit Sub
  ErrHandle:
    Screen.MousePointer = vbDefault
    MsgBox "Error" & vbCrLf & vbCrLf & "Error No. " & Err.Number & vbCrLf & " Error .Description:  " & Err.Description, vbInformation Or vbOKOnly
End Sub
