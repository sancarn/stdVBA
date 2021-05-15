VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdXLTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements stdITable

Public Event Selected()
Public Event Deselected()



'Table events
Public Event NewRowAdded(ByVal row As stdITableRow)
Public Event NewColumnAdded(ByVal sHeaderName As String)

'XLTable events
Public Event NewRowRangeSelected()
Public Event NewColumnRangeSelected()
Public Event HeaderChange(ByVal col As ListColumn, ByVal sHeaderName As String, ByRef sNewHeaderName As String)

'Row events
Public Event BeforeRowChange(ByVal row As stdITableRow, ByRef cancel As Boolean)
Public Event AfterRowChange(ByVal row As stdITableRow)
Public Event RowSelected(ByVal row As stdITableRow)   '
Public Event RowDeselected(ByVal row As stdITableRow)



'IMPLEMENTATION DETAIL PROPERTIES
Private WithEvents pSheet As Worksheet  'Latch onto sheet events to provide table events
Attribute pSheet.VB_VarHelpID = -1
Private pRange As Range
Private pIsSelected As Boolean          '
Private pIDColumnName As String         '

'************************************************************
'*   CONSTRUCTORS
'************************************************************
Public Function CreateFromRange(ByVal TableRange As Range, Optional ByVal sIDColumnName As String = "") As stdXLTable
  Set CreateFromRange = New stdXLTable
  Call CreateFromRange.protInit(TableRange, sIDColumnName)
End Function

Public Function CreateFromListObject(ByVal listObj As ListObject, Optional ByVal sIDColumnName As String = "") As stdXLTable
    Set CreateFromListObject = New stdXLTable
    Call CreateFromListObject.protInit(listObj.Range, sIDColumnName)
End Function


'***************************************
'*           PROPERTIES
'***************************************
Public Property Get isSelected() As Boolean
    isSelected = pIsSelected
End Property

Public Property Get idColumnName() As String
  idColumnName = pIDColumnName
End Property
Public Property Let idColumnName(ByVal sHeaderName As String)
  Dim vHeaders: vHeaders = Headers
  Dim i As Long: For i = 1 To UBound(vHeaders)
    If sHeaderName = vHeaders(i) Then
      pIDColumnName = sHeaderName
      Exit Property
    End If
  Next
  
  'No header matched, error
  Err.Raise 160, "idColumnName", "Column name provided is not present in headers array"
End Property
Public Property Get Headers() As Variant
  Dim v: v = HeaderRowRange.value
  Dim vRet(): ReDim vRet(1 To UBound(v, 2))
  Dim i As Long: For i = 1 To UBound(v, 2)
    vRet(i) = v(1, i)
  Next
  Headers = vRet
End Property
Public Property Get Range() As Range
  Set Range = pRange
End Property
Public Property Get HeaderRowRange() As Range
  Set HeaderRowRange = pRange.Rows(1)
End Property
Public Property Get DataBodyRange() As Range
  Set DataBodyRange = pRange.Offset(1).Resize(pRange.Rows.CountLarge - 1)
End Property
Public Property Get NewRowRange() As Range
  Set NewRowRange = pRange.Rows(1).Offset(pRange.Rows.CountLarge)
End Property
Public Property Get NewColumnRange() As Range
  Set NewColumnRange = pRange.Resize(, 1).Offset(, pRange.Columns.CountLarge)
End Property
Public Property Get ListObject() As ListObject
  Set ListObject = Range.Resize(1, 1).ListObject
End Property

'***************************************
'*             EVENTS
'***************************************
Private Sub sheet_SelectionChange(ByVal Target As Range)
    If TypeOf Selection Is Range Then
        If Not Application.intersects(Target, lo.Range) Is Nothing Then
            If pOldSelectedRows Is Nothing Then
                RaiseEvent Selected
                pIsSelected = True
            End If
            
            'Dim selectedRows As Collection: Set selectedRows = getSelectedRows(Target)
            'Dim rowsData As Collection: Set rowsData = getRowsData(selectedRows)
            RaiseEvent RowsSelected(Nothing, Nothing)
            
            
        Else
            RaiseEvent Deselected
            pIsSelected = False
            Set pSelectedRow = Nothing
            
        End If
    Else
        If pIsSelected Then
            RaiseEvent Deselected
            Set pOldSelectedRows = Nothing
            pIsSelected = False
        End If
    End If
End Sub





'***************************************
'*             PROTECTED
'***************************************
Friend Sub protInit(ByVal fullRange As Range, ByVal sIDColumnName As String)
    Set pRange = fullRange
    Set pSheet = fullRange.parent
    pIDColumnName = sIDColumnName
    
    If TypeOf Selection Is Range And Selection.parent Is pSheet Then
        pIsSelected = Not (Application.Intersect(Selection, listObj.Range) Is Nothing)
    End If
End Sub

'***************************************
'*             PRIVATE
'***************************************



Private Property Get stdITable_Column(ByVal sHeaderName As String) As stdITable

End Property

Private Function stdITable_getNewID() As Variant

End Function

Private Property Get stdITable_Headers() As Variant

End Property

Private Property Let stdITable_idColumnName(ByVal RHS As String)

End Property

Private Property Get stdITable_idColumnName() As String

End Property

Private Property Get stdITable_isSelected() As Boolean

End Property

Private Function stdITable_Message(ByVal eSubject As EMessageSubject, ParamArray params() As Variant) As Variant
  Dim retVar As Variant
  Select Case eSubject
    Case EMessageSubject.AfterRowChange
      RaiseEvent AfterRowChange(params(0))
    Case EMessageSubject.BeforeRowChange
      retVar = False
      RaiseEvent BeforeRowChange(params(0), retVar)
    Case EMessageSubject.RowDeselect
      RaiseEvent RowDeselected(params(0))
    Case EMessageSubject.RowSelect
      RaiseEvent RowSelected(params(0))
  End Select
  
  stdITable_Message = retVar
End Function

Private Property Get stdITable_Row(ByVal key As Variant) As Variant

End Property

Private Property Get stdITable_Rows() As Collection

End Property
