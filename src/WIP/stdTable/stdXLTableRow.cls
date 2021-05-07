VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdXLTableRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements stdITableRow

Private WithEvents sht As Worksheet
Attribute sht.VB_VarHelpID = -1

Private pDataDict As Object
Private pTable As stdITable
Private pRange As Range
Private pIsSelected As Boolean
Private pRows As Collection
Private pDiffs As Collection

Public Function Create(ByVal oParent As stdITable, ByRef vHeaders As Variant, ByVal vData As Variant, Optional ByVal rng As Range = Nothing) As stdXLTableRow
  Set Create = New stdXLTableRow
  
  'Create a dictionary storing the data
  Dim oDict As Object: Set oDict = CreateObject("Scripting.Dictionary")
  Dim i As Long
  For i = 1 To UBound(vHeaders)
    oDict(vHeaders(i)) = vData(i)
  Next
  
  Call Create.protInit(oParent, oDict, rng)
End Function
Public Function CreateFromDict(ByVal oParent As stdITable, ByVal oDict As Object, Optional ByVal rng As Range = Nothing)
  Set CreateFromDict = New stdXLTableRow
  Call CreateFromDict.protInit(oParent, oDict, rng)
End Function

'Obtain the row data as an array
'@returns {Variant()}
Public Function AsArray() As Variant
  Set AsArray = pDataDict.items()
End Function

'Obtain the row data as a dictionary. This dictionary is disconnected from the main dictionary
'@returns {Dictionary<string, variant>}
Public Function AsDict() As Object
  Dim newDict As Object: Set newDict = CreateObject("Scripting.Dictionary")
  Dim keys As Variant: keys = pDataDict.keys()
  Dim iKey As Long: For iKey = 0 To UBound(keys)
    Dim sKey As String: sKey = keys(iKey)
    newDict(sKey) = pDataDict(sKey)
  Next
  Set AsDict = newDict
End Function

'Create a clone of this row, linked to the same range
Public Function Clone() As stdXLTableRow
  Set Clone = CreateFromDict(pTable, AsDict(), pRange)
End Function

'Obtain the id of this row, if an id column name is provided
'@returns {variant} Id of row
Public Property Get id() As Variant
  If pTable.idColumnName <> "" Then
    id = pDataDict(pTable.idColumnName)
  End If
End Property

'Returns true if the row is selected

Public Property Get isSelected() As Boolean
  isSelected = pIsSelected
End Property
Public Sub SelectRow()
  pIsSelected = True
  pRange.Select
End Sub


'*********************************
'*   Table Event implementation
'*********************************
'Implementation of RowChange event
Private Sub sht_Change(ByVal Target As Range)
  With Target
    If .row = pRange.row And pRange.Column <= .Column And .Column <= (pRange.Column + pRange.Columns.Count) Then
      If TypeOf pTable Is stdXLTable Then
        Dim xlTable As stdXLTable: Set xlTable = pTable
        
        'Apply changes / diffs
        Dim area As Range
        For Each area In Target.Areas
          Dim cell As Range
          For Each cell In area.Cells
            Dim sHeaderName As String: sHeaderName = Application.Intersect(cell.EntireColumn, xlTable.HeaderRowRange)
            Dim val: val = cell.value
            Call pDiffs.Add(CreateDiff(sHeaderName, pDataDict(sHeaderName), val))
          Next
        Next
      End If
      
      'Message parent
      cancel = pTable.Message(BeforeRowChange, Me)
      
      'Change rows
      If cancel Then
        Set diff = New Collection
      End If
      
      'Commit change
      Call Save
      
      Call pTable.Message(AfterRowChange, Me)
    End If
  End With
End Sub
'Implementation of RowSelected event
Private Sub sht_SelectionChange(ByVal Target As Range)
  With Target
    If .row = pRange.row And pRange.Column <= .Column And .Column <= (pRange.Column + pRange.Columns.Count) Then
      If Not pIsSelected Then
        pIsSelected = True
        Call pTable.Message(RowSelect, Me)
      End If
    Else
      If pIsSelected Then
        pIsSelected = False
        Call pTable.Message(RowDeselect, Me)
      End If
    End If
  End With
End Sub

Sub Save()
  Dim diff As Variant
  For Each diff In pDiffs
    pDataDict(diff(0)) = diff(2)
  Next
  
  Dim h: h = pTable.Headers
  Dim v: v = pRange.value
  Dim i As Long: For i = 1 To UBound(h)
    v(1, i) = pDataDict(h(i))
  Next
  pRange.value = v
End Sub

'******************************
'*   stdITableRow interface
'******************************
Private Function stdITableRow_AsArray() As Variant
  Set stdITableRow_AsArray = AsArray()
End Function

Private Function stdITableRow_AsDict() As Object
  Set stdITableRow_AsDict = AsDict()
End Function

Private Function stdITableRow_Clone() As stdITableRow
  Set stdITableRow_Clone = Clone()
End Function

Private Property Get stdITableRow_Diff() As Collection
  Set stdITableRow_Diff = diff()
End Property

Private Property Get stdITableRow_id() As Variant
  stdITableRow_id = id()
End Property

Private Property Get stdITableRow_Item(ByVal sHeaderName As String) As Variant
  stdITableRow_Item = Item(sHeaderName)
End Property
Private Property Let stdITableRow_Item(ByVal sHeaderName As String, ByVal RHS As Variant)
  Item(sHeaderName) = RHS
End Property

Private Property Get stdITableRow_parent() As stdITable
  Set stdITableRow_parent = parent
End Property

Private Property Get stdITableRow_Range() As Range
  Set stdITableRow_Range = Range
End Property

Private Sub stdITableRow_Save()
  Call Save
End Sub


'***************************************
'*             PROTECTED
'***************************************
Friend Sub protInit(ByVal oParent As stdXLTable, ByVal oDict As Object, Optional ByVal rng As Range = Nothing)
  Set pTable = oParent
  Set pDataDict = oDict
  Set pRange = rng
End Sub

'***************************************
'*             HELPERS
'***************************************
Private Function CreateDiff(ByVal sHeaderName As String, ByVal vOld, ByVal vNew) As Variant
  CreateDiff = Array(sHeaderName, vOld, vNew)
End Function

