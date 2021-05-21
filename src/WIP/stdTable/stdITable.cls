VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdITable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum ETableSubject
  RowSelect
  RowDeselect
  
  BeforeRowChange
  AfterRowChange
  
  
End Enum

Public Event Selected()
Public Event Deselected()

'Table events
Public Event HeaderChange(ByVal sOldHeaderName As String, ByRef sNewHeaderName As String)


'Row events
Public Event RowChange(ByVal row As stdITableRow)
Public Event RowSelected(ByVal row As stdITableRow)   '
Public Event RowDeselected(ByVal row As stdITableRow)

'***************************************
'*           PROPERTIES
'***************************************
Public Property Get isSelected() As Boolean: End Property

Public Property Get idColumnName() As String: End Property
Public Property Let idColumnName(ByVal sHeaderName As String): End Property

'Obtain an array of headers
Public Property Get Headers() As Variant: End Property

'stdEnumerable wrappable collection
Public Property Get Rows() As Collection: End Property

'
Public Property Get row(ByVal key As Variant): End Property
Public Property Get Column(ByVal sHeaderName As String) As stdITable: End Property

'Obtain a new ID (can be used to create new rows)
Public Function getNewID() As Variant: End Function

Public Function Message(ByVal eSubject As ETableSubject, ParamArray params()) As Variant: End Function
