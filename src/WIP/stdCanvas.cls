VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Enum CanvasRectFields
  [_First] = 0
  cvfRectLeft = 0
  cvfRectTop
  cvfRectWidth
  cvfRectHeight
  [_Last] = cvfRectHeight
End Enum

Public Enum CanvasObjectStyles
  [_First] = 0
  cvsFillColor = 0
  cvsLineColor
  cvsLineWidth
  [_cvsFillVisible]
  [_cvsLineVisible]
  [_Last] = [_cvsLineVisible]
End Enum

#If VBA7 Then
  Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
  Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As LongPtr
  Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
  Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
  Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
  Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
  Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
  Private Declare PtrSafe Function BitBlt Lib "gdi32" (ByVal hDestDC As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
  Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As rect) As Long
  Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" (ByRef token As LongPtr, ByRef inputbuf As GdiplusStartupInput, ByRef outputbuf As GdiplusStartupOutput) As Long
  Private Declare PtrSafe Sub GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr)
  Private Declare PtrSafe Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As LongPtr, ByRef graphics As LongPtr) As Long
  Private Declare PtrSafe Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As LongPtr) As Long
  Private Declare PtrSafe Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, ByRef brush As LongPtr) As Long
  Private Declare PtrSafe Function GdipDeleteBrush Lib "gdiplus" (ByVal brush As LongPtr) As Long
  Private Declare PtrSafe Function GdipFillRectangle Lib "gdiplus" (ByVal graphics As LongPtr, ByVal brush As LongPtr, ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single) As Long
  Private Declare PtrSafe Function GdipDrawRectangle Lib "gdiplus" (ByVal graphics As LongPtr, ByVal pen As LongPtr, ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single) As Long
  Private Declare PtrSafe Function GdipDrawLine Lib "gdiplus" (ByVal graphics As LongPtr, ByVal pen As LongPtr, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Long
  Private Declare PtrSafe Function GdipCreatePen1 Lib "gdiplus" (ByVal color As Long, ByVal width As Single, ByVal unit As Long, ByRef pen As LongPtr) As Long
  Private Declare PtrSafe Function GdipDeletePen Lib "gdiplus" (ByVal pen As LongPtr) As Long
  Private Declare PtrSafe Function GdipGraphicsClear Lib "gdiplus" (ByVal graphics As LongPtr, ByVal color As Long) As Long
#Else
  Enum LongPtr
    [_]
  End Enum
  Private Declare Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
  Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As LongPtr
  Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
  Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
  Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
  Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
  Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
  Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
  Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As rect) As Long
  Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As LongPtr, ByRef inputbuf As GdiplusStartupInput, ByRef outputbuf As GdiplusStartupOutput) As Long
  Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr)
  Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As LongPtr, ByRef graphics As LongPtr) As Long
  Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As LongPtr) As Long
  Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, ByRef brush As LongPtr) As Long
  Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal brush As LongPtr) As Long
  Private Declare Function GdipFillRectangle Lib "gdiplus" (ByVal graphics As LongPtr, ByVal brush As LongPtr, ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single) As Long
  Private Declare Function GdipDrawRectangle Lib "gdiplus" (ByVal graphics As LongPtr, ByVal pen As LongPtr, ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single) As Long
  Private Declare Function GdipDrawLine Lib "gdiplus" (ByVal graphics As LongPtr, ByVal pen As LongPtr, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Long
  Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal color As Long, ByVal width As Single, ByVal unit As Long, ByRef pen As LongPtr) As Long
  Private Declare Function GdipDeletePen Lib "gdiplus" (ByVal pen As LongPtr) As Long
  Private Declare Function GdipGraphicsClear Lib "gdiplus" (ByVal graphics As LongPtr, ByVal color As Long) As Long
#End If

Private Const SRCCOPY As Long = &HCC0020

Private Type GdiplusStartupInput
  GdiplusVersion As Long
  DebugEventCallback As LongPtr
  SuppressBackgroundThread As Long
  SuppressExternalCodecs As Long
End Type

Private Type GdiplusStartupOutput
  NotificationHook As LongPtr
  NotificationUnhook As LongPtr
End Type

Public Enum CanvasEvent
  Click
  MouseDown
  MouseUp
  MouseMove
  MouseEnter
  MouseLeave
End Enum

' Type for RECT structure
Private Type rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Enum CanvasObjectType
  cvtRectangle
  cvtCircle
  cvtLine
End Enum

Private Type CanvasObject
  ObjectType As CanvasObjectType
  ObjectParams() As Variant
  ObjectStyle() As Long
  
  ''TODO: Currently unimplemented
  'zIndex As Double
  
  'TODO: Implemented but not utilised. Use for events and such
  bbox As rect
End Type

Private Type TThis
  obj As Object
  hWnd As LongPtr
  hDC As LongPtr
  gdipToken As LongPtr
  objectsIndex As Long
  objects() As CanvasObject
  
  ' Double buffering components
  bufferDC As LongPtr
  bufferBitmap As LongPtr
  bufferWidth As Long
  bufferHeight As Long
End Type
Private This As TThis

Public Function CreateFromControl(ByVal f As MSForms.Control) As stdCanvas
  Set CreateFromControl = CreateFromHwnd(f.[_GethWnd], f)
End Function
Public Function CreateFromUserform(ByVal uf As MSForms.UserForm) As stdCanvas
  'TBC GetFromIUnknown
End Function

Public Function CreateFromHwnd(ByVal hWnd As LongPtr, Optional ByVal obj As Object = Nothing) As stdCanvas
  Set CreateFromHwnd = New stdCanvas
  Call CreateFromHwnd.protInit(hWnd, obj)
End Function

Public Sub protInit(ByVal hWnd As LongPtr, ByVal obj As Object)
  Set This.obj = obj
  This.hWnd = hWnd
  This.hDC = GetDC(This.hWnd)
  If This.hDC = 0 Then
    Err.Raise 5, "stdCanvas#protInit", "Failed to get device context."
    Exit Sub
  End If
  ReDim This.objects(1 To 8)
  This.objectsIndex = 0
  
  'Initialise gdip
  Dim gdip As GdiplusStartupInput
  gdip.GdiplusVersion = 1
  Dim gdipo As GdiplusStartupOutput
  If GdiplusStartup(This.gdipToken, gdip, gdipo) <> 0 Then
    Err.Raise 5, "stdCanvas#protInit", "Failed to start GDI+"
  End If
  
  ' Initialize double buffer
  Dim clientRect As rect
  Call GetClientRect(This.hWnd, clientRect)
  This.bufferWidth = clientRect.Right - clientRect.Left
  This.bufferHeight = clientRect.Bottom - clientRect.Top
  
  This.bufferDC = CreateCompatibleDC(This.hDC)
  If This.bufferDC <> 0 Then
    This.bufferBitmap = CreateCompatibleBitmap(This.hDC, This.bufferWidth, This.bufferHeight)
    If This.bufferBitmap <> 0 Then
      Call SelectObject(This.bufferDC, This.bufferBitmap)
    End If
  End If
End Sub



Private Sub Class_Terminate()
  'TODO: When sprites/fonts added
  'Dim i As Long
  'For i = 1 To UBound(This.objects)
  '  Erase bitmaps/fonts etc.
  'Next
  
  ' Cleanup double buffer
  If This.bufferBitmap <> 0 Then Call DeleteObject(This.bufferBitmap)
  If This.bufferDC <> 0 Then Call DeleteDC(This.bufferDC)
  
  If This.gdipToken <> 0 Then Call GdiplusShutdown(This.gdipToken)
  If This.hDC <> 0 Then Call ReleaseDC(This.hWnd, This.hDC)
End Sub

'Draw a rectangle
'@param left - X Coord of left edge
'@param top - Y Coord of top edge
'@param width - Width of rectangle
'@param height - Height of rectangle
'@returns - Pointer to a rectangle handle
Public Function DrawRectangle(ByVal left As Long, ByVal top As Long, ByVal width As Long, ByVal height As Long, Optional ByVal fillColor As Long = 0, Optional ByVal borderColor As Long = 0, Optional ByVal borderWeight As Long = 1) As Long
  Dim index As Long: index = getNewObject()
  With This.objects(index)
    .ObjectType = cvtRectangle
    ReDim .ObjectParams(CanvasRectFields.[_First] To CanvasRectFields.[_Last])
    ReDim .ObjectStyle(CanvasObjectStyles.[_First] To CanvasObjectStyles.[_Last])
  End With
  Data(index, CanvasRectFields.cvfRectLeft) = left
  Data(index, CanvasRectFields.cvfRectTop) = top
  Data(index, CanvasRectFields.cvfRectWidth) = width
  Data(index, CanvasRectFields.cvfRectHeight) = height
  Style(index, cvsFillColor) = fillColor
  Style(index, cvsLineColor) = borderColor
  Style(index, cvsLineWidth) = borderWeight
  
  DrawRectangle = index
End Function

Public Property Get Data(ByVal handle As Long, ByVal param As Long) As Variant
  Data = This.objects(handle).ObjectParams(param)
End Property
Public Property Let Data(ByVal handle As Long, ByVal param As Long, ByVal RHS As Variant)
  With This.objects(handle)
    .ObjectParams(param) = RHS
    Select Case .ObjectType
      Case cvtRectangle
        Select Case param
          Case CanvasRectFields.cvfRectLeft
            .bbox.Left = RHS
            .bbox.Right = RHS + .ObjectParams(CanvasRectFields.cvfRectWidth)
          Case CanvasRectFields.cvfRectTop
            .bbox.Top = RHS
            .bbox.Bottom = RHS + .ObjectParams(CanvasRectFields.cvfRectHeight)
          Case CanvasRectFields.cvfRectWidth
            .bbox.Right = .ObjectParams(CanvasRectFields.cvfRectLeft) + RHS
          Case CanvasRectFields.cvfRectHeight
            .bbox.Bottom = .ObjectParams(CanvasRectFields.cvfRectTop) + RHS
        End Select
    End Select
  End With
End Property

Public Property Get Style(ByVal handle As Long, ByVal param As CanvasObjectStyles) As Long
  Style = This.objects(handle).ObjectStyle(param)
End Property
Public Property Let Style(ByVal handle As Long, ByVal param As CanvasObjectStyles, RHS As Long)
  With This.objects(handle)
    .ObjectStyle(param) = RHS
    Select Case param
      Case CanvasObjectStyles.cvsFillColor
        .ObjectStyle(CanvasObjectStyles.[_cvsFillVisible]) = getAlpha(RHS) > 0
      Case CanvasObjectStyles.cvsLineColor
        .ObjectStyle(CanvasObjectStyles.[_cvsLineVisible]) = getAlpha(RHS) > 0
      Case CanvasObjectStyles.cvsLineWidth
        .ObjectStyle(CanvasObjectStyles.[_cvsLineVisible]) = RHS > 0
    End Select
  End With
End Property

Public Sub Redraw()
  Const UnitPixel As Long = 2
  
  ' Check if buffer needs resizing
  Dim clientRect As rect
  Call GetClientRect(This.hWnd, clientRect)
  Dim newWidth As Long: newWidth = clientRect.Right - clientRect.Left
  Dim newHeight As Long: newHeight = clientRect.Bottom - clientRect.Top
  
  If This.bufferBitmap = 0 Or newWidth > This.bufferWidth Or newHeight > This.bufferHeight Then
    ' Recreate buffer if needed
    If This.bufferBitmap <> 0 Then Call DeleteObject(This.bufferBitmap)
    If This.bufferDC <> 0 Then Call DeleteDC(This.bufferDC)
    
    This.bufferWidth = newWidth + 50  ' Some padding to reduce frequent resizing
    This.bufferHeight = newHeight + 50
    This.bufferDC = CreateCompatibleDC(This.hDC)
    If This.bufferDC <> 0 Then
      This.bufferBitmap = CreateCompatibleBitmap(This.hDC, This.bufferWidth, This.bufferHeight)
      If This.bufferBitmap <> 0 Then Call SelectObject(This.bufferDC, This.bufferBitmap)
    End If
  End If
  
  ' Use double buffering if available, otherwise fall back to direct
  If This.bufferBitmap <> 0 Then
    ' Create GDI+ graphics object from buffer DC
    Dim G As LongPtr
    If GdipCreateFromHDC(This.bufferDC, G) = 0 Then
      ' Clear the buffer
      Call GdipGraphicsClear(G, argb(0, 255, 255, 255))
      
      ' Draw all objects to the buffer
      Dim i As Long
      For i = 1 To This.objectsIndex
        With This.objects(i)
          Select Case .ObjectType
            Case cvtRectangle
              If .ObjectStyle(CanvasObjectStyles.[_cvsFillVisible]) Then
                Dim fillBrush As LongPtr
                If GdipCreateSolidFill(.ObjectStyle(cvsFillColor), fillBrush) = 0 Then
                  Call GdipFillRectangle(G, fillBrush, .ObjectParams(cvfRectLeft), .ObjectParams(cvfRectTop), .ObjectParams(cvfRectWidth), .ObjectParams(cvfRectHeight))
                  Call GdipDeleteBrush(fillBrush)
                End If
              End If
              If .ObjectStyle(CanvasObjectStyles.[_cvsLineVisible]) Then
                Dim strokePen As LongPtr
                If GdipCreatePen1(.ObjectStyle(cvsLineColor), .ObjectStyle(cvsLineWidth), UnitPixel, strokePen) = 0 Then
                  Call GdipDrawRectangle(G, strokePen, .ObjectParams(cvfRectLeft), .ObjectParams(cvfRectTop), .ObjectParams(cvfRectWidth), .ObjectParams(cvfRectHeight))
                  Call GdipDeletePen(strokePen)
                End If
              End If
          End Select
        End With
      Next
      
      Call GdipDeleteGraphics(G)
      
      ' Copy buffer to screen
      Call BitBlt(This.hDC, 0, 0, newWidth, newHeight, This.bufferDC, 0, 0, SRCCOPY)
      Exit Sub
    End If
  End If
  
  ' Fallback to direct rendering
  Call RedrawDirect
End Sub

' Fallback method for direct rendering (without double buffering)
Private Sub RedrawDirect()
  Const UnitPixel As Long = 2
  
  Dim G As LongPtr
  If GdipCreateFromHDC(This.hDC, G) <> 0 Then Exit Sub
  
  If Not This.obj Is Nothing Then Call This.obj.Repaint
  
  Dim i As Long
  For i = 1 To This.objectsIndex
    With This.objects(i)
      Select Case .ObjectType
        Case cvtRectangle
          If .ObjectStyle(CanvasObjectStyles.[_cvsFillVisible]) Then
            Dim fillBrush As LongPtr
            If GdipCreateSolidFill(.ObjectStyle(cvsFillColor), fillBrush) = 0 Then
              Call GdipFillRectangle(G, fillBrush, .ObjectParams(cvfRectLeft), .ObjectParams(cvfRectTop), .ObjectParams(cvfRectWidth), .ObjectParams(cvfRectHeight))
              Call GdipDeleteBrush(fillBrush)
            End If
          End If
          If .ObjectStyle(CanvasObjectStyles.[_cvsLineVisible]) Then
            Dim strokePen As LongPtr
            If GdipCreatePen1(.ObjectStyle(cvsLineColor), .ObjectStyle(cvsLineWidth), UnitPixel, strokePen) = 0 Then
              Call GdipDrawRectangle(G, strokePen, .ObjectParams(cvfRectLeft), .ObjectParams(cvfRectTop), .ObjectParams(cvfRectWidth), .ObjectParams(cvfRectHeight))
              Call GdipDeletePen(strokePen)
            End If
          End If
      End Select
    End With
  Next
  
  Call GdipDeleteGraphics(G)
End Sub

' Property to check if double buffering is available
Public Property Get DoubleBuffered() As Boolean
  DoubleBuffered = (This.bufferBitmap <> 0)
End Property

Private Function getNewObject() As Long
  This.objectsIndex = This.objectsIndex + 1
  Dim ub As Long: ub = UBound(This.objects)
  If ub < This.objectsIndex Then
    ReDim Preserve This.objects(1 To ub * 2)
  End If
  getNewObject = This.objectsIndex
End Function

' Convert RGB to ARGB
' @param A - Alpha channel (0-255)
' @param R - Red channel (0-255)
' @param G - Green channel (0-255)
' @param B - Blue channel (0-255)
' @returns - ARGB color
Public Function argb(A As Byte, R As Byte, G As Byte, B As Byte) As Long
  Const AlphaShift As Double = &H1000000
  Const RedShift As Long = &H10000
  Const GreenShift As Long = &H100
  
  ' 2^31–1 = max signed 32-bit; 2^32 = wrap modulus
  Const INT32_MAX As Long = 2 ^ 31 - 1
  Const UINT32_MODULO As Double = 2 ^ 32
  
  'Shift the colors without causing an overflow, by shifting in the double domain
  Dim rawValue As Double: rawValue = A * AlphaShift + R * RedShift + G * GreenShift + B
  
  ' If rawValue exceeds INT32_MAX, wrap into the signed 32-bit range
  If rawValue > INT32_MAX Then rawValue = rawValue - UINT32_MODULO
  argb = rawValue
End Function

Private Function getAlpha(ByVal u As Double) As Byte
  If u < 0 Then u = u + 4294967296#
  getAlpha = CByte(Fix(u / 16777216#) Mod 256)
End Function