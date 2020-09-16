VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private WithEvents shTests As Worksheet
Attribute shTests.VB_VarHelpID = -1
Private loTests As ListObject
Public Sub Init()
    Set shTests = getSheet("Test")
    
    If Not loExists(shTests, "Tests") Then
        'Create listobject and conditional formatting
        shTests.Range("A1").value = "Type"
        shTests.Range("B1").value = "Description"
        shTests.Range("C1").value = "Passed"
        
        'Create list object
        Set loTests = shTests.ListObjects.Add(XlListObjectSourceType.xlSrcRange, Range("A1:C2"), False, xlYes)
        loTests.name = "Tests"
        
        Dim fc As FormatCondition
        Set fc = shTests.Range("$A:$C").FormatConditions.Add(xlExpression, Formula1:="=$A1=""Category""")
        fc.Font.Color = RGB(255, 255, 255)
        fc.Font.Bold = True
        fc.Interior.Color = RGB(100, 100, 100)
        
        'Highlight Passes
        Set fc = loTests.ListColumns("Passed").Range.FormatConditions.Add(xlCellValue, XlFormatConditionOperator.xlEqual, True)
        fc.Font.Color = RGB(0, 100, 0)
        fc.Interior.Color = RGB(200, 255, 200)
        
        'Highlight Fails
        Set fc = loTests.ListColumns("Passed").Range.FormatConditions.Add(xlCellValue, XlFormatConditionOperator.xlEqual, False)
        fc.Font.Color = RGB(100, 0, 0)
        fc.Interior.Color = RGB(255, 200, 200)
    Else
        'Get list object
        Set loTests = shTests.ListObjects("Tests")
    End If
    
    Call Refresh
End Sub
Public Sub Refresh()
    If loTests Is Nothing Then
        Me.Init
    Else
        If Not loTests.DataBodyRange Is Nothing Then
            loTests.DataBodyRange.EntireRow.Delete
        End If
    End If
End Sub
Public Sub RefreshGlobalFormats()
    Dim fc As FormatCondition
    For Each fc In shTests.Range("A:C").FormatConditions
        If fc.Formula1 = "=$A1=""Category""" Then
            Call fc.ModifyAppliesToRange(shTests.Range("A:C"))
        End If
    Next
End Sub
Public Sub Topic(ByVal sName As String)
    If loTests Is Nothing Then Me.Init
    
    Dim lr As ListRow
    Set lr = loTests.ListRows.Add()
    lr.Range.value = Array("Category", sName, Empty)
    If loTests.ListRows.Count = 1 Then Call RefreshGlobalFormats
End Sub
Public Sub Assert(ByVal sName As String, ByVal result As Boolean)
    If loTests Is Nothing Then Me.Init
    
    Dim lr As ListRow
    Set lr = loTests.ListRows.Add()
    lr.Range.value = Array("Test", sName, result)
    If loTests.ListRows.Count = 1 Then Call RefreshGlobalFormats
End Sub


'Helpers
Private Function getSheet(ByVal sName As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.name = sName Then
            Set getSheet = ws
            Exit Function
        End If
    Next
    
    Set getSheet = ThisWorkbook.Sheets.Add()
    getSheet.name = sName
End Function

Private Function loExists(ByVal ws As Worksheet, ByVal sName As String) As Boolean
    On Error Resume Next
    loExists = Not ws.ListObjects(sName) Is Nothing
End Function

Private Sub shTests_BeforeDelete()
    Set loTests = Nothing
End Sub
