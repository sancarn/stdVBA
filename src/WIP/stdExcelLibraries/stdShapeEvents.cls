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
'  Dim WithEvents shpEvents As CShapeEvents
'  
'  Sub latchEvents()
'    Set shpEvents = New CShapeEvents
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
'run latchEvents and observe the hooked events in the immediate window.






Private old_selection As Object
Private WithEvents bars As commandBars
Attribute bars.VB_VarHelpID = -1

'Fake events:
Public Event Selected(shape As shape)
Public Event Deselected(shape As shape)
Public Event Changed(shape As shape)

Private Sub bars_OnUpdate()
  Dim xSel As Object
  Set xSel = Selection
  
  'If selection is a shape then it could have changed or been selected,
  'otherwise if the old selection contained a shape and the new doesn't then
  'shape has been deselected
  If DetectShape(Selection) Then
    'Use the name to decide if it has been
    If GetName(old_selection) = GetName(Selection) Then
      'Raise Changed event
      RaiseEvent Changed(Selection.ShapeRange(1))
    Else
      'Raise Selected event
      RaiseEvent Selected(Selection.ShapeRange(1))
    End If
  Else
    'Ensure old selection was a shape
    If DetectShape(old_selection) Then
      'Raise Deselected event
      RaiseEvent Deselected(old_selection.ShapeRange(1))
    End If
  End If
  
  'Keep track of old selection
  Set old_selection = Selection
End Sub

Private Function GetName(ByVal obj As Object) As String
  On Error Resume Next
  GetName = obj.Name
End Function

Private Function DetectShape(ByVal obj As Object) As Boolean
  On Error GoTo endDetect
    DetectShape = obj.ShapeRange.Count > 0
endDetect:
End Function


Private Sub Class_Initialize()
  Set bars = Application.commandBars
End Sub