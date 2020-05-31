VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdShapeEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Example.bas
'
'  Dim WithEvents shpEvents As stdShapeEvents
'
'  Sub latchEvents()
'    Set shpEvents = stdShapeEvents.Create(Sheet2)
'    Call shpEvents.HookSheet(Sheet2)
'  End Sub
'
'  Private Sub shpEvents_Changed(shape As shape)
'    Debug.Print shape.Name & " Changed"
'  End Sub
'
'  Private Sub shpEvents_Deselected(shape As shape)
'    Debug.Print shape.Name & " Deselected"
'  End Sub
'
'  Private Sub shpEvents_Selected(shape As shape)
'    Debug.Print shape.Name & " Selected"
'  End Sub
'
'  Private Sub shpEvents_Created(shape As shape)
'    Debug.Print shape.Name & " Created"
'  End Sub
'
'  Private Sub shpEvents_Deleted(ByVal shpName As string)
'    Debug.Print shpName & " Deleted"
'  End Sub
'
'  Private Sub shpEvents_Renamed(shape As shape, ByVal oldName as string)
'    Debug.Print oldName & " renamed to " & shape.name
'  End Sub
'run latchEvents and observe the hooked events in the immediate window.


'TODO:
' * Add Moved(shape as Shape)



Private old_selection As Object
Private WithEvents bars As CommandBars

'Fake events:
Public Event Selected(shp As shape)
Public Event Deselected(shp As shape)
Public Event Changed(shp As shape)
Public Event Created(shp As shape)
Public Event Deleted(ByVal shpName As String)
Public Event Renamed(shp As shape, ByVal oldName As String)

Public SheetShapesDict As Object

Private Sub bars_OnUpdate()
  Dim xSel As Object
  Set xSel = Selection
  
  If isShape(Selection) Or isShape(old_selection) Then
    If Not SheetShapesDict Is Nothing Then
      'Get active sheet codename:
      Dim shtName As String
      shtName = ActiveSheet.CodeName
  
      'Ensure active sheet codename exists in dictionary (via HookSheet)
      If SheetShapesDict.exists(shtName) Then
        'Ensure shape counts are different
        Dim shp As shape
        If SheetShapesDict(shtName)("=COUNT") <> ActiveSheet.Shapes.Count Then
          If SheetShapesDict(shtName)("=COUNT") < ActiveSheet.Shapes.Count Then
            'Shape has been created
            For Each shp In ActiveSheet.Shapes
              If Not SheetShapesDict(shtName).exists(shp.Name) Then
                SheetShapesDict(shtName)("=COUNT") = SheetShapesDict(shtName)("=COUNT") + 1
                Set SheetShapesDict(shtName)(shp.Name) = shp
                RaiseEvent Created(shp)
              End If
            Next
          Else
            'Shape has been deleted
            Dim shpName As Variant
            For Each shpName In SheetShapesDict(shtName).keys()
              Set shp = getShapeByName(ActiveSheet, shpName)
              
              If Left(shpName, 1) <> "=" Then
                If shp Is Nothing Then
                  SheetShapesDict(shtName)("=COUNT") = SheetShapesDict(shtName)("=COUNT") - 1
                  SheetShapesDict(shtName).Remove shpName
                  RaiseEvent Deleted(shpName)
                End If
              End If
            Next
          End If
        Else
          'Shape might have been renamed
          'Identify new name:
          Dim existingShape As shape
          For Each shp In ActiveSheet.Shapes
            If Not SheetShapesDict(shtName).exists(shp.Name) Then
              Set SheetShapesDict(shtName)(shp.Name) = shp
              Set existingShape = shp
            End If
          Next
          For Each shpName In SheetShapesDict(shtName).keys()
            If Left(shpName, 1) <> "=" Then
              Set shp = getShapeByName(ActiveSheet, shpName)
              If shp Is Nothing Then
                SheetShapesDict(shtName).Remove shpName
                RaiseEvent Renamed(existingShape, shpName)
              End If
            End If
          Next
        End If
      End If
    End If
  
  
  
  
    'If selection is a shape then it could have changed or been selected,
    'otherwise if the old selection contained a shape and the new doesn't then
    'shape has been deselected
    If DetectShape(Selection) Then
      'Use the name to decide if it has been changed or selected
      If GetName(old_selection) = GetName(Selection) Then
        'Raise Changed event - doesn't actually imply the shape changed...
        
        'if hash(shp) <> old_hash then ...
          RaiseEvent Changed(Selection.ShapeRange(1))
        'end if
      Else
        'Raise Selected event
        RaiseEvent Selected(Selection.ShapeRange(1))
      End If
    Else
      'Ensure old selection was a shape
      If DetectShape(old_selection) Then
        'If shapeExists(old_selection.ShapeRange) Then
          'Raise Deselected event
          RaiseEvent Deselected(old_selection.ShapeRange(1))
        'End If
      End If
    End If
  End If
  
  'Keep track of old selection
  Set old_selection = Selection
End Sub

Public Sub HookSheet(ByVal sht As Worksheet)
  If SheetShapesDict Is Nothing Then Set SheetShapesDict = CreateObject("Scripting.Dictionary")
  Set SheetShapesDict(sht.CodeName) = CreateObject("Scripting.Dictionary")
  SheetShapesDict(sht.CodeName)("=COUNT") = sht.Shapes.Count
  

  Dim shp As shape
  For Each shp In sht.Shapes
    Set SheetShapesDict(sht.CodeName)(shp.Name) = shp
  Next
End Sub

Private Function getShapeByName(sht As Worksheet, ByVal sName As String) As shape
  Dim shp As shape
  For Each shp In sht.Shapes
    If shp.Name = sName Then
      Set getShapeByName = shp
      Exit Function
    End If
  Next
  Set getShapeByName = Nothing
End Function

Private Function GetName(ByVal obj As Object) As String
  On Error Resume Next
  GetName = obj.Name
End Function

Private Function DetectShape(ByVal obj As Object) As Boolean
  On Error GoTo endDetect
    DetectShape = obj.ShapeRange.Count > 0
endDetect:
End Function

Private Function isShape(ByVal obj As Object) As Boolean
  Select Case TypeName(obj)
    Case "Rectangle", "Arc", "Drawing", "Picture"
      isShape = True
    Case Else
      isShape = False
  End Select
End Function


Private Sub Class_Initialize()
  Set bars = Application.CommandBars
End Sub