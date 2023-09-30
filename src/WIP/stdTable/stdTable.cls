VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Event Selected()
Public Event Deselected()

Public Event HeaderChange(ByVal col As ListColumn, ByVal sHeaderName As String, ByRef sNewHeaderName As String)
Public Event RowChange(ByVal row As ListRow, ByRef rowData As Object)

Public Event RowsSelected(ByVal rows As Collection, ByRef rowsData As Collection)
Public Event RowsDeselected(ByVal rows As Collection, ByRef rowsData As Collection)




Private WithEvents pSheet As Worksheet
Attribute pSheet.VB_VarHelpID = -1

Private pLO As ListObject

Private pIsSelected As Boolean
Private pSelectedRows As Collection


'************************************************************
'*   CONSTRUCTORS
'************************************************************
Public Function Create(ByVal listObj As ListObject) As stdTable
    Set Create = New stdTable
    Call Create.Init(listObj)
End Function


'***************************************
'*           PROPERTIES
'***************************************
Public Property Get IsSelected() As Boolean
    IsSelected = pIsSelected
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
            RaiseEvent RowsSelected(Nothing, rowsData)
            
            
        Else
            RaiseEvent Deselected(getRowObject(pSelectedRow))
            pIsSelected = False
            Set pSelectedRow = Nothing
            
        End If
    Else
        If pIsSelected Then
            RaiseEvent Deselected(pOldSelectedRows)
            Set pOldSelectedRows = Nothing
            pIsSelected = False
        End If
    End If
End Sub





'***************************************
'*             PROTECTED
'***************************************
Friend Sub Init(ByVal listObj As ListObject)
    Set pLO = listObj
    Set pSheet = listObj.Parent
    
    If TypeOf Selection Is Range And Selection.Parent Is pSheet Then
        pIsSelected = Not (Application.Intersect(Selection, listObj.Range) Is Nothing)
    End If
End Sub

'***************************************
'*             PRIVATE
'***************************************
Private Function getListRows(r As Range) As Collection
    Dim res As Collection: Set res = New Collection
    If Not Application.Intersect(r, pLO.Range) Is Nothing Then
        Dim rng As Range
        Set rng = Application.Intersect(r.EntireRow, pLO.Range)
        Dim are As Range
        For Each are In rng.Areas
        
    End If
    Set getListRows = res
End Function


