VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdITableRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = true
'Attribute Item.VB_UserMemId = 0

Enum ETableRowSubject
  RowSelected
  RowDeselected
  RowChange
End Enum

'Return as an array of variants
Public Function AsArray() As Variant: End Function

'Return as a dictionary
Public Function AsDict() As Object: End Function

'Get the id of the row.
Public Property Get id() As Variant: End Property

'Get collection of changes undergone by this row
Public Property Get diff() As Collection: End Property

'Get / Set the item to a value
Public Property Get Item(ByVal sHeaderName As String) As Variant: End Property
Public Property Let Item(ByVal sHeaderName As String, ByVal value As Variant): End Property

'Get whether the row has been changed or not
Friend Property Get Dirty() As Boolean: End Property
Friend Property Let Dirty(bDirty As Boolean): End Property

'Get the parent table
Public Property Get parent() As stdITable: End Property

'Get the range of the table this relates too
'@returns {Range} If not attributed to a range this value is nothing. Else returns the corresponding range.
Public Property Get Range() As Range: End Property

'Clones this row and all it's data to a new object with the same properties
Public Function Clone() As stdITableRow: End Function

'Save row to the database table, Diff must be cleared on save
Public Sub Save(): End Sub

Public Function Message(ByVal eSubject As ETableRowSubject, ParamArray params()) As Variant: End Function
