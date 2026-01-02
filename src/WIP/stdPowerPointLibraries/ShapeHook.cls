VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShapeHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents cb As CommandBars
Attribute cb.VB_VarHelpID = -1
Private Type TShape
  name As String
  x As Double
  y As Double
  w As Double
  h As Double
  mark As Boolean
End Type
Private Type TCollection
  count As Long
  shapes() As TShape
End Type
Private Type TThis
  view As slide
  viewIndex As Long
  list As TCollection
End Type
Private This As TThis


'Usage:
'  Dim shpHook As ShapeHook
'  Sub commandButt_click()
'    set shpHook = ShapeHook.CreateActive()
'  End Sub
'  Private Sub shpHook_ShapeAdded(ByVal shp As Shape)
'    Debug.Print "Shape added: " & shp.Name
'  End Sub
'  Private Sub shpHook_ShapeRemoved(ByVal shp As Shape)
'    Debug.Print "Shape removed: " & shp.Name
'  End Sub
'  Private Sub shpHook_ShapeMoved(ByVal shp As Shape, fromX As Double, fromY As Double, toX As Double, toY As Double)
'    Debug.Print "Shape moved: " & shp.Name & " from (" & fromX & ", " & fromY & ") to (" & toX & ", " & toY & ")"
'  End Sub

Public Event ShapeAdded(ByVal shp As Shape)
Public Event ShapeRemoved(ByVal name As String)
Public Event ShapeMoved(ByVal shp As Shape, fromX As Double, fromY As Double, toX As Double, toY As Double)
Public Event ShapeResized(ByVal shp As Shape, fromW As Double, fromH As Double, toW As Double, toH As Double)


'Create from the active slide
'@returns - object running shape events
Public Function CreateActive() As ShapeHook
  Dim sl As slide: Set sl = Application.ActiveWindow.view.slide
  Set CreateActive = ShapeHook.CreateFromSlide(sl)
End Function

'Create from a provided slide
'@param view - The slide to hook shape events too
'@returns - object running shape events
Public Function CreateFromSlide(ByVal view As slide)
  Set CreateFromSlide = New ShapeHook
  Call CreateFromSlide.protInit(view)
End Function

'Init on slide
Public Sub protInit(ByVal view As slide)
  Set This.view = view
  This.viewIndex = view.SlideIndex
  Set cb = view.Application.CommandBars
  
  This.list.count = view.shapes.count
  ReDim This.list.shapes(1 To This.list.count)
  
  
  Dim shp As Shape
  For Each shp In This.view.shapes
    Dim i As Long: i = i + 1
    With This.list.shapes(i)
      .name = shp.name 'unique identifier
      .x = shp.Left
      .y = shp.Top
      .w = shp.width
      .h = shp.Height
    End With
  Next
End Sub

'Pseudo-event hook
Private Sub cb_OnUpdate()
  If This.view Is Nothing Then Set This.view = Application.ActivePresentation.Slides(This.viewIndex)
  Dim slideShapeCount As Long: slideShapeCount = This.view.shapes.count
  
  'Mark all shapes as unprocessed
  Dim i As Long
  For i = 1 To This.list.count
    With This.list.shapes(i)
      .mark = False
    End With
  Next
  
  'Process add/remove events
  Dim shp As Shape
  If This.list.count < slideShapeCount Then
    'Shape has been added
    For Each shp In This.view.shapes
      If Not doesShapeExistInTCollection(This.list, shp.name) Then
        RaiseEvent ShapeAdded(shp)
        This.list.count = This.list.count + 1
        ReDim Preserve This.list.shapes(1 To This.list.count)
        With This.list.shapes(This.list.count)
          .name = shp.name
          .x = shp.Left
          .y = shp.Top
          .w = shp.width
          .h = shp.Height
          .mark = True
        End With
        If This.list.count = slideShapeCount Then Exit For
      End If
    Next
  ElseIf This.list.count > This.view.shapes.count Then
    'Shape has been removed
    For i = This.list.count To 1 Step -1
      If Not This.list.shapes(i).mark Then
        If getShapeFromCollection(This.view.shapes, This.list.shapes(i).name) Is Nothing Then
          RaiseEvent ShapeRemoved(This.list.shapes(i).name)
          For j = i To This.list.count - 1
            LSet This.list.shapes(j) = This.list.shapes(j + 1)
          Next
          This.list.count = This.list.count - 1
          ReDim Preserve This.list.shapes(1 To This.list.count)
          If This.list.count = slideShapeCount Then Exit For
        End If
      End If
    Next
  End If

  'Process change events
  For i = 1 To This.list.count
    With This.list.shapes(i)
      If Not .mark Then
        Set shp = getShapeFromCollection(This.view.shapes, .name)
        If Not shp Is Nothing Then
          If .x <> shp.Left Or .y <> shp.Top Then
            RaiseEvent ShapeMoved(shp, .x, .y, shp.Left, shp.Top)
            .x = shp.Left
            .y = shp.Top
          End If
          If .w <> shp.width Or .h <> shp.Height Then
            RaiseEvent ShapeResized(shp, .w, .h, shp.width, shp.Height)
            .w = shp.width
            .h = shp.Height
          End If
          .mark = True
        End If
      End If
    End With
  Next
End Sub

'Get a shape from a TCollection by name
'@param tCol - The collection to search
'@param name - The name of the shape to search for
'@returns - The shape if found, otherwise Nothing
Private Function doesShapeExistInTCollection(ByRef shps As TCollection, ByRef name As String) As Boolean
  Dim i As Long
  For i = 1 To shps.count
    If shps.shapes(i).name = name Then
      doesShapeExistInTCollection = True
      Exit Function
    End If
  Next
End Function

'Get a shape from a Shapes collection by name
'@param shps - The collection to search
'@param name - The name of the shape to search for
'@returns - The shape if found, otherwise Nothing
Private Function getShapeFromCollection(ByRef shps As shapes, ByRef name As String) As Shape
  On Error Resume Next
  Set getShapeFromCollection = shps.Item(name)
End Function
