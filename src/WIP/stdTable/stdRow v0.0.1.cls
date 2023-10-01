VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private pDictionary As Object

'Constructors
Public Function Create() As stdRow
    Set Create = New stdRow
    Call Create.protInit(CreateObject("Scripting.Dictionary"))
End Function
Public Function CreateFromSelection() As stdRow
    Dim lo As ListObject
    If TypeOf Selection Is Range Then
        If Not Selection.ListObject Is Nothing Then
            Dim lo As ListObject: Set lo = Selection.ListObject
            Dim iRow As Long: iRow = Selection.Row - lo.Range.Row
            If iRow < 1 Then iRow = 1
            Set CreateFromSelection = CreateFromListRow(lo.ListRows(iRow))
        Else
            Set CreateFromSelection = Nothing
        End If
    Else
        Set CreateFromSelection = Nothing
    End If
End Function
Public Function CreateFromListRow(ByVal lr As ListRow) As stdRow
    'Generate dictionary from items
    Dim lo As ListObject: Set lo = lr.Parent
    Dim vHeaders: vHeaders = lo.HeaderRowRange.Value
    Dim vRow: vRow = lr.Range.Value
    Dim oDict As Object: Set oDict = CreateObject("Scripting.Dictionary")
    Dim iKey As Long: For iKey = 1 To UBound(vHeaders, 2)
        oDict(vHeaders(1, iKey)) = vRow(1, iKey)
    Next
    
    'Return row
    Set CreateFromListRow = New stdRow
    Call CreateFromListRow.protInit(oDict)
End Function

'Generate an array representing the data in the row
Public Function AsArray() As Variant
    AsArray = pDictionary.items()
End Function

'Generate a 2D-array representing the data in the row. Can be used to write directly
Public Function AsRowArray() As Variant
    Dim items: items = pDictionary.items()
    Dim vRow(): ReDim vRow(1 To 1, 1 To UBound(items) + 1)
    Dim i As Long: For i = 1 To UBound(items) + 1
        vRow(1, i) = items(i - 1)
    Next
    AsRowArray = vRow
End Function

'Get number of columns in the row
Public Property Get ColumnCount() As Long
    ColumnCount = UBound(pDictionary.items()) + 1
End Property

'Obtain a value from the row based on key header or index
Public Property Get Item(ByVal key As Variant)
    Select Case VarType(key)
        Case vbLong, vbInteger
            Item = pDictionary(pDictionary.keys(key - 1))
        Case Else
            Item = pDictionary(key)
    End Select
End Property
Public Property Let Item(ByVal key As Variant, ByVal v As Variant)
    Select Case VarType(key)
        Case vbLong, vbInteger
            pDictionary(pDictionary.keys(key - 1)) = v
        Case Else
            pDictionary(key) = v
    End Select
End Property

'Obtain a value from the row based on key header or index
Public Function getItem(ByVal key As Variant) As Variant
    getItem = Item(key)
End Function
Public Function setItem(ByVal key As Variant, ByVal v As Variant) As stdRow
    Item(key) = v
    Set setItem = Me
End Function




Friend Sub protInit(ByVal oDictionary As Object)
    Set pDictionary = oDictionary
End Sub
