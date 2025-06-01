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

Public Enum CanvasEllipseFields
  [_First] = 0
  cvfEllipseCenterX = 0
  cvfEllipseCenterY
  cvfEllipseWidth
  cvfEllipseHeight
  [_Last] = cvfEllipseHeight
End Enum

Public Enum CanvasArcFields
  [_First] = 0
  cvfArcCenterX = 0
  cvfArcCenterY
  cvfArcWidth
  cvfArcHeight
  cvfArcStartAngle
  cvfArcSweepAngle
  [_Last] = cvfArcSweepAngle
End Enum

Public Enum CanvasPolygonFields
  [_First] = 0
  cvfPolygonPoints = 0
  cvfPolygonCount
  [_Last] = cvfPolygonCount
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
  Private Declare PtrSafe Function GdipFillEllipse Lib "gdiplus" (ByVal graphics As LongPtr, ByVal brush As LongPtr, ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single) As Long
  Private Declare PtrSafe Function GdipDrawEllipse Lib "gdiplus" (ByVal graphics As LongPtr, ByVal pen As LongPtr, ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single) As Long
  Private Declare PtrSafe Function GdipFillPolygon Lib "gdiplus" (ByVal graphics As LongPtr, ByVal brush As LongPtr, ByRef points As Any, ByVal count As Long, ByVal fillMode As Long) As Long
  Private Declare PtrSafe Function GdipDrawPolygon Lib "gdiplus" (ByVal graphics As LongPtr, ByVal pen As LongPtr, ByRef points As Any, ByVal count As Long) As Long
  Private Declare PtrSafe Function GdipFillPie Lib "gdiplus" (ByVal graphics As LongPtr, ByVal brush As LongPtr, ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
  Private Declare PtrSafe Function GdipDrawArc Lib "gdiplus" (ByVal graphics As LongPtr, ByVal pen As LongPtr, ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
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
  Private Declare Function GdipFillEllipse Lib "gdiplus" (ByVal graphics As LongPtr, ByVal brush As LongPtr, ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single) As Long
  Private Declare Function GdipDrawEllipse Lib "gdiplus" (ByVal graphics As LongPtr, ByVal pen As LongPtr, ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single) As Long
  Private Declare Function GdipFillPolygon Lib "gdiplus" (ByVal graphics As LongPtr, ByVal brush As LongPtr, ByRef points As Any, ByVal count As Long, ByVal fillMode As Long) As Long
  Private Declare Function GdipDrawPolygon Lib "gdiplus" (ByVal graphics As LongPtr, ByVal pen As LongPtr, ByRef points As Any, ByVal count As Long) As Long
  Private Declare Function GdipFillPie Lib "gdiplus" (ByVal graphics As LongPtr, ByVal brush As LongPtr, ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
  Private Declare Function GdipDrawArc Lib "gdiplus" (ByVal graphics As LongPtr, ByVal pen As LongPtr, ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
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

' Point structure for polygon points
Private Type PointF
  x As Single
  y As Single
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
  cvtEllipse
  cvtArc
  cvtPolygon
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


'Create a new stdCanvas object from a Control
'@constructor
'@param f - Control to create the canvas from e.g a frame
'@returns - New stdCanvas object
Public Function CreateFromControl(ByVal cntrl As MSForms.Control) As stdCanvas
  Set CreateFromControl = CreateFromHwnd(cntrl.[_GethWnd], cntrl)
End Function

'Create a new stdCanvas object from a Frame
'@constructor
'@param frm - Frame to create the canvas from e.g a frame
'@returns - New stdCanvas object
Public Function CreateFromFrame(ByVal frm as MSForms.Frame) As stdCanvas
  Set CreateFromFrame = CreateFromControl(frm)
End Function


'Create a new stdCanvas object from a UserForm
'@constructor
'@param uf - UserForm to create the canvas from
'@returns - New stdCanvas object
Public Function CreateFromUserform(ByVal uf As MSForms.UserForm) As stdCanvas
  'TBC GetFromIUnknown
End Function

'Create a new stdCanvas object from a window handle
'@constructor
'@param hWnd - Window handle to create the canvas from
'@param obj - Object to call Repaint() on when the canvas needs to be redrawn
'@returns - New stdCanvas object
Public Function CreateFromHwnd(ByVal hWnd As LongPtr, Optional ByVal obj As Object = Nothing) As stdCanvas
  Set CreateFromHwnd = New stdCanvas
  Call CreateFromHwnd.protInit(hWnd, obj)
End Function

'Initialize the canvas
'@protected
'@param hWnd - Window handle to create the canvas from
'@param obj - Object to call Repaint() on when the canvas needs to be redrawn
Public Sub protInit(ByVal hWnd As LongPtr, ByVal obj As Object, Optional ByVal doubleBuffer As Boolean = True)
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
  If doubleBuffer then
    Dim clientRect As rect: Call GetClientRect(This.hWnd, clientRect)
    Call ResizeBuffer(clientRect.Right - clientRect.Left, clientRect.Bottom - clientRect.Top)
  End if
End Sub

'Terminate the canvas and clean up resources
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

'Draw an arc
'@param centerX - X Coord of arc center
'@param centerY - Y Coord of arc center
'@param width - Width of arc
'@param height - Height of arc
'@param startAngle - Start angle of arc
'@param sweepAngle - Sweep angle of arc
'@param fillColor - Fill color of arc
'@param borderColor - Border color of arc
'@param borderWeight - Border weight of arc
'@returns - Pointer to a arc handle
Public Function DrawArc(ByVal centerX As Long, ByVal centerY As Long, ByVal width As Long, ByVal height As Long, ByVal startAngle As Single, ByVal sweepAngle As Single, Optional ByVal fillColor As Long = 0, Optional ByVal borderColor As Long = 0, Optional ByVal borderWeight As Long = 1) As Long
  Dim index As Long: index = getNewObject()
  With This.objects(index)
    .ObjectType = cvtArc
    ReDim .ObjectParams(CanvasArcFields.[_First] To CanvasArcFields.[_Last])
    ReDim .ObjectStyle(CanvasObjectStyles.[_First] To CanvasObjectStyles.[_Last])
    
    ' Set bounding box
    .bbox.Left = centerX - width / 2
    .bbox.Top = centerY - height / 2
    .bbox.Right = centerX + width / 2
    .bbox.Bottom = centerY + height / 2
  End With
  
  Data(index, cvfArcCenterX) = centerX
  Data(index, cvfArcCenterY) = centerY
  Data(index, cvfArcWidth) = width
  Data(index, cvfArcHeight) = height
  Data(index, cvfArcStartAngle) = startAngle
  Data(index, cvfArcSweepAngle) = sweepAngle
  Style(index, cvsFillColor) = fillColor
  Style(index, cvsLineColor) = borderColor
  Style(index, cvsLineWidth) = borderWeight
  
  DrawArc = index
End Function

'Draw an ellipse
'@param centerX - X Coord of ellipse center
'@param centerY - Y Coord of ellipse center
'@param width - Width of ellipse
'@param height - Height of ellipse
'@param fillColor - Fill color of ellipse
'@param borderColor - Border color of ellipse
'@param borderWeight - Border weight of ellipse
'@returns - Pointer to an ellipse handle
Public Function DrawEllipse(ByVal centerX As Long, ByVal centerY As Long, ByVal width As Long, ByVal height As Long, Optional ByVal fillColor As Long = 0, Optional ByVal borderColor As Long = 0, Optional ByVal borderWeight As Long = 1) As Long
  Dim index As Long: index = getNewObject()
  With This.objects(index)
    .ObjectType = cvtEllipse
    ReDim .ObjectParams(CanvasEllipseFields.[_First] To CanvasEllipseFields.[_Last])
    ReDim .ObjectStyle(CanvasObjectStyles.[_First] To CanvasObjectStyles.[_Last])
    
    ' Set bounding box
    .bbox.Left = centerX - width / 2
    .bbox.Top = centerY - height / 2
    .bbox.Right = centerX + width / 2
    .bbox.Bottom = centerY + height / 2
  End With
  
  Data(index, cvfEllipseCenterX) = centerX
  Data(index, cvfEllipseCenterY) = centerY
  Data(index, cvfEllipseWidth) = width
  Data(index, cvfEllipseHeight) = height
  Style(index, cvsFillColor) = fillColor
  Style(index, cvsLineColor) = borderColor
  Style(index, cvsLineWidth) = borderWeight
  
  DrawEllipse = index
End Function


'Draw a circle
'@param centerX - X Coord of circle center
'@param centerY - Y Coord of circle center
'@param radius - Radius of circle
'@param fillColor - Fill color of circle
'@param borderColor - Border color of circle
'@param borderWeight - Border weight of circle
'@returns - Pointer to a circle handle
Public Function DrawCircle(ByVal centerX As Long, ByVal centerY As Long, ByVal radius As Long, Optional ByVal fillColor As Long = 0, Optional ByVal borderColor As Long = 0, Optional ByVal borderWeight As Long = 1) As Long
  DrawCircle = DrawEllipse(centerX, centerY, radius * 2, radius * 2, fillColor, borderColor, borderWeight)
End Function


'Draw a polygon
'@param points - Array of points (x1, y1, x2, y2, ...)
'@param fillColor - Fill color of polygon
'@param borderColor - Border color of polygon
'@param borderWeight - Border weight of polygon
'@returns - Pointer to a polygon handle
Public Function DrawPolygon(ByRef points() As Single, Optional ByVal fillColor As Long = 0, Optional ByVal borderColor As Long = 0, Optional ByVal borderWeight As Long = 1) As Long
  If (UBound(points) - LBound(points) + 1) Mod 2 <> 0 Then
    Err.Raise 5, "stdCanvas#DrawPolygon", "Points array must contain an even number of elements (x,y pairs)"
    Exit Function
  End If
  
  Dim pointCount As Long: pointCount = (UBound(points) - LBound(points) + 1) / 2
  If pointCount < 3 Then
    Err.Raise 5, "stdCanvas#DrawPolygon", "Polygon must have at least 3 points"
    Exit Function
  End If
  
  Dim index As Long: index = getNewObject()
  With This.objects(index)
    .ObjectType = cvtPolygon
    ReDim .ObjectParams(CanvasPolygonFields.[_First] To CanvasPolygonFields.[_Last])
    ReDim .ObjectStyle(CanvasObjectStyles.[_First] To CanvasObjectStyles.[_Last])
    
    ' Calculate bounding box
    Dim minX As Single, maxX As Single, minY As Single, maxY As Single
    minX = points(LBound(points)): maxX = points(LBound(points))
    minY = points(LBound(points) + 1): maxY = points(LBound(points) + 1)
    
    Dim i As Long
    For i = LBound(points) To UBound(points) Step 2
      If points(i) < minX Then minX = points(i)
      If points(i) > maxX Then maxX = points(i)
      If points(i + 1) < minY Then minY = points(i + 1)
      If points(i + 1) > maxY Then maxY = points(i + 1)
    Next
    
    .bbox.Left = Int(minX): .bbox.Top = Int(minY)
    .bbox.Right = Int(maxX) + 1: .bbox.Bottom = Int(maxY) + 1
  End With
  
  This.objects(index).ObjectParams(cvfPolygonPoints) = points
  This.objects(index).ObjectParams(cvfPolygonCount) = pointCount
  Style(index, cvsFillColor) = fillColor
  Style(index, cvsLineColor) = borderColor
  Style(index, cvsLineWidth) = borderWeight
  
  DrawPolygon = index
End Function

'Get/Set the data of an object
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
      Case cvtCircle
        Select Case param
          Case CanvasCircleFields.cvfCircleCenterX
            .bbox.Left = RHS - .ObjectParams(CanvasCircleFields.cvfCircleRadius)
            .bbox.Right = RHS + .ObjectParams(CanvasCircleFields.cvfCircleRadius)
          Case CanvasCircleFields.cvfCircleCenterY
            .bbox.Top = RHS - .ObjectParams(CanvasCircleFields.cvfCircleRadius)
            .bbox.Bottom = RHS + .ObjectParams(CanvasCircleFields.cvfCircleRadius)
          Case CanvasCircleFields.cvfCircleRadius
            Dim centerX As Long: centerX = .ObjectParams(CanvasCircleFields.cvfCircleCenterX)
            Dim centerY As Long: centerY = .ObjectParams(CanvasCircleFields.cvfCircleCenterY)
            .bbox.Left = centerX - RHS
            .bbox.Top = centerY - RHS
            .bbox.Right = centerX + RHS
            .bbox.Bottom = centerY + RHS
        End Select
    End Select
  End With
End Property

'Get/Set the style of an object
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

'Redraw the canvas
Public Sub Redraw()
  ' Check if buffer needs resizing
  Dim clientRect As rect
  Call GetClientRect(This.hWnd, clientRect)
  Dim newWidth As Long: newWidth = clientRect.Right - clientRect.Left
  Dim newHeight As Long: newHeight = clientRect.Bottom - clientRect.Top
  
  If This.bufferBitmap = 0 Or newWidth > This.bufferWidth Or newHeight > This.bufferHeight Then
    Call ResizeBuffer(newWidth, newHeight)
  End If
  
  ' Create graphics context (buffered or direct)
  Dim useBuffer As Boolean: useBuffer = (This.bufferBitmap <> 0)
  
  'Get target DC
  Dim targetDC As LongPtr
  If useBuffer Then
    targetDC = This.bufferDC
  Else
    targetDC = This.hDC
    
    'Clear canvas if not using buffer
    If Not This.obj Is Nothing Then Call This.obj.Repaint 
  End If

  Dim graphics As LongPtr
  If GdipCreateFromHDC(targetDC, graphics) <> 0 Then Exit Sub
  
  ' Clear background if using buffer
  Call GdipGraphicsClear(graphics, argb(255, 255, 255, 255))
  
  ' Render all objects using unified method
  Dim i As Long
  For i = 1 To This.objectsIndex
    Call RenderObject(graphics, This.objects(i))
  Next
  
  Call GdipDeleteGraphics(graphics)
  
  ' Copy buffer to screen if double buffering
  If useBuffer Then
    Call BitBlt(This.hDC, 0, 0, newWidth, newHeight, This.bufferDC, 0, 0, SRCCOPY)
  End If
End Sub

'Property to check if double buffering is available
'@returns - True if double buffering is enabled, False otherwise
Public Property Get DoubleBuffered() As Boolean
  DoubleBuffered = (This.bufferBitmap <> 0)
End Property

'Get a new object handle
'@returns - New object handle
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

'Get the alpha channel from an ARGB value
'@param u - ARGB value
'@returns - Alpha channel (0-255)
Private Function getAlpha(ByVal u As Double) As Byte
  If u < 0 Then u = u + 4294967296#
  getAlpha = CByte(Fix(u / 16777216#) Mod 256)
End Function

'Render an object to the graphics context
'@param graphics - Graphics context to render to
'@param obj - Object to render
Private Sub RenderObject(ByVal graphics As LongPtr, ByRef obj As CanvasObject)
  Const UnitPixel As Long = 2
  Const FillModeAlternate As Long = 0
  
  With obj
    ' Handle fill first
    If .ObjectStyle(CanvasObjectStyles.[_cvsFillVisible]) Then
      Dim fillBrush As LongPtr
      If GdipCreateSolidFill(.ObjectStyle(cvsFillColor), fillBrush) = 0 Then
        Select Case .ObjectType
          Case cvtRectangle
            Call GdipFillRectangle(graphics, fillBrush, _
              .ObjectParams(cvfRectLeft), .ObjectParams(cvfRectTop), _
              .ObjectParams(cvfRectWidth), .ObjectParams(cvfRectHeight))
              
          Case cvtEllipse
            Dim ellipseX As Single: ellipseX = .ObjectParams(cvfEllipseCenterX) - .ObjectParams(cvfEllipseWidth) / 2
            Dim ellipseY As Single: ellipseY = .ObjectParams(cvfEllipseCenterY) - .ObjectParams(cvfEllipseHeight) / 2
            Call GdipFillEllipse(graphics, fillBrush, ellipseX, ellipseY, _
              .ObjectParams(cvfEllipseWidth), .ObjectParams(cvfEllipseHeight))
              
          Case cvtArc
            Dim arcX As Single: arcX = .ObjectParams(cvfArcCenterX) - .ObjectParams(cvfArcWidth) / 2
            Dim arcY As Single: arcY = .ObjectParams(cvfArcCenterY) - .ObjectParams(cvfArcHeight) / 2
            Call GdipFillPie(graphics, fillBrush, arcX, arcY, _
              .ObjectParams(cvfArcWidth), .ObjectParams(cvfArcHeight), _
              .ObjectParams(cvfArcStartAngle), .ObjectParams(cvfArcSweepAngle))
              
          Case cvtPolygon
            Dim polyPoints() As Single: polyPoints = .ObjectParams(cvfPolygonPoints)
            Dim polyCount As Long: polyCount = .ObjectParams(cvfPolygonCount)
            Dim gdipPoints() As PointF: gdipPoints = ConvertToPointFArray(polyPoints, polyCount)
            Call GdipFillPolygon(graphics, fillBrush, gdipPoints(0), polyCount, FillModeAlternate)
        End Select
        Call GdipDeleteBrush(fillBrush)
      End If
    End If
    
    ' Handle stroke/outline
    If .ObjectStyle(CanvasObjectStyles.[_cvsLineVisible]) Then
      Dim strokePen As LongPtr
      If GdipCreatePen1(.ObjectStyle(cvsLineColor), .ObjectStyle(cvsLineWidth), UnitPixel, strokePen) = 0 Then
        Select Case .ObjectType
          Case cvtRectangle
            Call GdipDrawRectangle(graphics, strokePen, _
              .ObjectParams(cvfRectLeft), .ObjectParams(cvfRectTop), _
              .ObjectParams(cvfRectWidth), .ObjectParams(cvfRectHeight))
              
          Case cvtEllipse
            ellipseX = .ObjectParams(cvfEllipseCenterX) - .ObjectParams(cvfEllipseWidth) / 2
            ellipseY = .ObjectParams(cvfEllipseCenterY) - .ObjectParams(cvfEllipseHeight) / 2
            Call GdipDrawEllipse(graphics, strokePen, ellipseX, ellipseY, _
              .ObjectParams(cvfEllipseWidth), .ObjectParams(cvfEllipseHeight))
              
          Case cvtArc
            arcX = .ObjectParams(cvfArcCenterX) - .ObjectParams(cvfArcWidth) / 2
            arcY = .ObjectParams(cvfArcCenterY) - .ObjectParams(cvfArcHeight) / 2
            Call GdipDrawArc(graphics, strokePen, arcX, arcY, _
              .ObjectParams(cvfArcWidth), .ObjectParams(cvfArcHeight), _
              .ObjectParams(cvfArcStartAngle), .ObjectParams(cvfArcSweepAngle))
              
          Case cvtPolygon
            polyPoints = .ObjectParams(cvfPolygonPoints)
            polyCount = .ObjectParams(cvfPolygonCount)
            gdipPoints = ConvertToPointFArray(polyPoints, polyCount)
            Call GdipDrawPolygon(graphics, strokePen, gdipPoints(0), polyCount)
        End Select
        Call GdipDeletePen(strokePen)
      End If
    End If
  End With
End Sub

'Resize the double buffer to the new dimensions
'@param newWidth - New width of the buffer
'@param newHeight - New height of the buffer
Private Sub ResizeBuffer(ByVal newWidth As Long, ByVal newHeight As Long)
  If This.bufferBitmap <> 0 Then Call DeleteObject(This.bufferBitmap)
  If This.bufferDC <> 0 Then Call DeleteDC(This.bufferDC)
  
  This.bufferWidth = newWidth
  This.bufferHeight = newHeight
  
  This.bufferDC = CreateCompatibleDC(This.hDC)
  If This.bufferDC <> 0 Then
    This.bufferBitmap = CreateCompatibleBitmap(This.hDC, This.bufferWidth, This.bufferHeight)
    If This.bufferBitmap <> 0 Then
      Call SelectObject(This.bufferDC, This.bufferBitmap)
    End If
  End If
End Sub

'Convert an array of points to an array of PointF structures
'@param points - Array of points
'@param count - Number of points
'@returns - Array of PointF structures
Private Function ConvertToPointFArray(ByRef points() As Single, ByVal count As Long) As PointF()
  Dim gdipPoints() As PointF
  ReDim gdipPoints(0 To count - 1)
  Dim j As Long
  For j = 0 To count - 1
    gdipPoints(j).x = points(LBound(points) + j * 2)
    gdipPoints(j).y = points(LBound(points) + j * 2 + 1)
  Next
  ConvertToPointFArray = gdipPoints
End Function
