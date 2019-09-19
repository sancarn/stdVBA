VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAddressClassicFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Parent-Class of cAddressClassic: the usual VB6-approach - just for comparison (of coding-efforts and performance)

Private mCol As New Collection

Public Function Add(ByVal ID As Long, Name As String, LastName As String, ByVal BirthDay As Date) As cAddressClassic
  Set Add = New cAddressClassic
      Add.ID = ID
      Add.Name = Name
      Add.BirthDay = BirthDay
  
  mCol.Add Add
End Function
 
Public Property Get Item(Key) As vbIAddress
  Set Item = mCol.Item(Key)
End Property

Public Property Get Count() As Long
  Count = mCol.Count
End Property

Public Function Enumerator() As stdole.IUnknown
Attribute Enumerator.VB_UserMemId = -4
  Set Enumerator = mCol.[_NewEnum]
End Function

Public Sub RemoveAll()
  Set mCol = Nothing
End Sub

