VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdXLEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = true
Public Event SelectionChange(selOld As Object, selNew As Object)

Public Event RangeSelectionChange(rOldRange As Range, rNewRange As Range)
Public Event RangeSelected(Target As Range)
Public Event RangeDeselected(Target As Range)

Public Event ShapeSelectionChange(cOld As Collection, cNew As Collection)
Public Event ShapeSelected(Target As Shape)
Public Event ShapeDeselected(Target As Shape)

Public Event ChartSelected(Target As Shape)
Public Event ChartDeselected(Target As Shape)

Private WithEvents cbs As CommandBars
Attribute cbs.VB_VarHelpID = -1


Private Enum ESelectionType
  EmptySelection
  Axis
  AxisTitle
  Chart
  ChartArea
  ChartObject
  ChartObjects
  Charts
  ChartTitle
  Corners
  DataLabel
  DataLabels
  DataTable
  DisplayUnitLabel
  DownBars
  DropLines
  ErrorBars
  Floor
  Gridlines
  HiLoLines
  LeaderLines
  Legend
  LegendEntry
  LegendKey
  OLEObject
  OLEObjects
  PlotArea
  Point
  Range
  Series
  SeriesLines
  Shape
  ShapeRange
  Sheets
  TickLabels
  Trendline
  UpBars
  Walls
End Enum
Private Type SelectionInfo
  iType As ESelectionType
  oObj As Object
  oWB As Workbook
  oWS As Worksheet
  sID As String
  sPath As String
  
  vMeta As Variant
End Type



Private Sub cbs_OnUpdate()
  Static oldSelection As SelectionInfo 'EmptySelection by default
  Dim newSelection As SelectionInfo: newSelection = getSelection()
  If selectionsAreDifferent(oldSelection, newSelection) Then
    RaiseEvent SelectionChange(oldSelection.oObj, newSelection.oObj)
    
    'Range events
    If oldSelection.iType = Range Or newSelection.iType = Range Then
      RaiseEvent RangeSelectionChange(IIf(oldSelection.iType = Range, oldSelection.oObj, Nothing), IIf(newSelection.iType = Range, newSelection.oObj, Nothing))
      If newSelection.iType = Range Then
        If oldSelection.iType = Range Then
          RaiseEvent RangeDeselected(oldSelection.oObj)
        End If
        RaiseEvent RangeSelected(newSelection.oObj)
      Else
        RaiseEvent RangeDeselected(oldSelection.oObj)
      End If
    End If
    
    'Shape events
    If oldSelection.iType = Shape Or newSelection.iType = Shape Then
      RaiseEvent ShapeSelectionChange(getShapesCol(oldSelection), getShapesCol(newSelection))
      If newSelection.iType = Shape Then
        If oldSelection.iType = Shape Then
          Call RaiseShapeEvent(oldSelection.oObj.ShapeRange(1), False)
        End If
        Call RaiseShapeEvent(newSelection.oObj.ShapeRange(1), True)
      Else
        Call RaiseShapeEvent(oldSelection.oObj.ShapeRange(1), False)
      End If
    End If
    
    'Shape range
    Dim shp As Shape
    If oldSelection.iType = ShapeRange Or newSelection.iType = ShapeRange Then
      RaiseEvent ShapeSelectionChange(getShapesCol(oldSelection), getShapesCol(newSelection))
      If newSelection.iType = ShapeRange Then
        If oldSelection.iType = ShapeRange Then
          For Each shp In oldSelection.oObj.ShapeRange
            Call RaiseShapeEvent(shp, False)
          Next
        End If
        For Each shp In newSelection.oObj.ShapeRange
          Call RaiseShapeEvent(shp, True)
        Next
      Else
        For Each shp In oldSelection.oObj.ShapeRange
          Call RaiseShapeEvent(shp, False)
        Next
      End If
    End If
    
    'Chart
    If oldSelection.iType = ESelectionType.ChartArea Or newSelection.iType = ChartArea Then
      RaiseEvent ShapeSelectionChange(getShapesCol(oldSelection), getShapesCol(newSelection))
      If newSelection.iType = ChartArea Then
        If oldSelection.iType = ChartArea Then
          Call RaiseShapeEvent(oldSelection.oObj.parent.parent.ShapeRange(1), False)
        End If
        Call RaiseShapeEvent(newSelection.oObj.parent.parent.ShapeRange(1), True)
      Else
        Call RaiseShapeEvent(oldSelection.oObj.parent.parent.ShapeRange(1), False)
      End If
    End If
    
    Debug.Print "Selection changed to " & newSelection.sPath
  End If
  oldSelection = newSelection
End Sub

Private Sub Class_Initialize()
  Set cbs = Application.CommandBars
End Sub



Private Function getSelection() As SelectionInfo
  Set getSelection.oWB = getObjAncestor(Selection, "Workbook")
  Set getSelection.oWS = getObjAncestor(Selection, "Worksheet")
  Set getSelection.oObj = Selection
  getSelection.sID = getObjID(Selection)
  getSelection.sPath = getObjPath(Selection)
  getSelection.iType = getObjType(Selection)
End Function
Public Sub RaiseShapeEvent(shp As Shape, bSelected As Boolean)
  If shp.Type = MsoShapeType.msoChart Then
    If bSelected Then
      RaiseEvent ChartSelected(shp)
    Else
      RaiseEvent ChartDeselected(shp)
    End If
  Else
    If bSelected Then
      RaiseEvent ShapeSelected(shp)
    Else
      RaiseEvent ShapeDeselected(shp)
    End If
  End If
End Sub

Private Function getShapesCol(ByRef tInf As SelectionInfo) As Collection
  Set getShapesCol = New Collection
  Select Case tInf.iType
    Case ESelectionType.Shape
      getShapesCol.Add tInf.oObj
    Case ESelectionType.ShapeRange
      Dim shp As Shape: For Each shp In tInf.oObj.ShapeRange
        getShapesCol.Add shp
      Next
  End Select
End Function

Public Function getObjPath(ByRef obj As Object) As String
  Dim sPath As String
  sPath = TypeName(obj) & ">" & getObjID(obj)
  While TypeName(obj) <> "Workbook" And TypeName(obj) <> "Application"
    Set obj = getObjParent(obj)
    sPath = TypeName(obj) & ">" & getObjID(obj) & ">" & sPath
    'DoEvents
  Wend
  getObjPath = sPath
End Function
Public Function getObjID(ByRef obj As Object) As String
  Select Case TypeName(obj)
    Case "Range": getObjID = obj.Address
    Case "Axis": getObjID = obj.Type & "-" & obj.AxisGroup
    Case "ChartGroup": getObjID = "ChartGroup"
    Case "LegendEntry": getObjID = obj.Index
    Case "DataTable": getObjID = "DataTable"
    Case "DrawingObjects"
      Dim sID As String: sID = ""
      Dim shp As Shape: For Each shp In obj.ShapeRange
        sID = sID & getObjID(shp) & ";"
      Next
      getObjID = sID
    Case Else: getObjID = obj.Name
  End Select
End Function
Public Function getObjAncestor(ByRef obj As Object, ByVal sTypeName As String) As Variant
  While TypeName(obj) <> sTypeName
    Set obj = getObjParent(obj)
    'DoEvents
  Wend
  Set getObjAncestor = obj
End Function
Public Function getObjParent(ByRef obj As Object) As Object
  Dim parent As Object
  Set parent = obj.parent
  Select Case TypeName(obj)
    Case "Axis"
      If Not ActiveChart Is Nothing Then Set parent = ActiveChart
    Case "DataTable"
      If Not ActiveChart Is Nothing Then Set parent = ActiveChart
  End Select
  Set getObjParent = parent
End Function

Private Function selectionsAreDifferent(ByRef tOld As SelectionInfo, ByRef tNew As SelectionInfo)
  If tOld.iType = tNew.iType Then
    If tOld.oWB Is tNew.oWB Then
      If tOld.oWS Is tNew.oWS Then
        Select Case tOld.iType
        Case Else
          selectionsAreDifferent = tOld.sPath <> tNew.sPath
        End Select
      Else
        selectionsAreDifferent = True
      End If
    Else
      selectionsAreDifferent = True
    End If
  Else
    selectionsAreDifferent = True
  End If
End Function

'Rectange, Oval, Drawing, Arc, Textbox
Private Function getObjType(obj As Object) As ESelectionType
  Dim iRet As ESelectionType
  Select Case TypeName(obj)
    Case "Range": iRet = ESelectionType.Range
    Case "Shape", "Rectangle", "Oval", "Drawing", "Arc", "Textbox": iRet = ESelectionType.Shape
    Case "ShapeRange", "DrawingObjects": iRet = ESelectionType.ShapeRange
    Case "Chart": iRet = ESelectionType.Chart
    Case "Point": iRet = ESelectionType.Point
    Case "Series": iRet = ESelectionType.Series
    Case "SeriesLines": iRet = ESelectionType.SeriesLines
    Case "OLEObject": iRet = ESelectionType.OLEObject
    Case "OLEObjects": iRet = ESelectionType.OLEObjects
    
    Case "Axis": iRet = ESelectionType.Axis
    Case "AxisTitle": iRet = ESelectionType.AxisTitle
    
    Case "ChartArea": iRet = ESelectionType.ChartArea
    Case "ChartObject": iRet = ESelectionType.ChartObject
    Case "ChartObjects": iRet = ESelectionType.ChartObjects
    Case "Charts": iRet = ESelectionType.Charts
    Case "ChartTitle": iRet = ESelectionType.ChartTitle
    Case "Corners": iRet = ESelectionType.Corners
    Case "DataLabel": iRet = ESelectionType.DataLabel
    Case "DataLabels": iRet = ESelectionType.DataLabels
    Case "DataTable": iRet = ESelectionType.DataTable
    Case "DisplayUnitLabel": iRet = ESelectionType.DisplayUnitLabel
    Case "DownBars": iRet = ESelectionType.DownBars
    Case "DropLines": iRet = ESelectionType.DropLines
    Case "ErrorBars": iRet = ESelectionType.ErrorBars
    Case "Floor": iRet = ESelectionType.Floor
    Case "Gridlines": iRet = ESelectionType.Gridlines
    Case "HiLoLines": iRet = ESelectionType.HiLoLines
    Case "LeaderLines": iRet = ESelectionType.LeaderLines
    Case "Legend": iRet = ESelectionType.Legend
    Case "LegendEntry": iRet = ESelectionType.LegendEntry
    Case "LegendKey": iRet = ESelectionType.LegendKey
    Case "PlotArea": iRet = ESelectionType.PlotArea
    Case "Sheets": iRet = ESelectionType.Sheets
    Case "TickLabels": iRet = ESelectionType.TickLabels
    Case "Trendline": iRet = ESelectionType.Trendline
    Case "UpBars": iRet = ESelectionType.UpBars
    Case "Walls": iRet = ESelectionType.Walls
  End Select
  getObjType = iRet
End Function
