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
Private pFT As Boolean
Dim dtLastCall As Long

#If Mac Then
   #If MAC_OFFICE_VERSION >= 15 Then
      Private Declare Function GetTickCount Lib "/Applications/Microsoft Excel.app/Contents/Frameworks/MicrosoftOffice.framework/MicrosoftOffice" () As Long
   #Else
      Private Declare Function GetTickCount Lib "Applications:Microsoft Office 2011:Office:MicrosoftOffice.framework:MicrosoftOffice" () As Long
   #End If
#Else ' Win32 or Win64
   #If VBA7 Then
      Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
   #Else
      Private Declare Function GetTickCount Lib "kernel32" () As Long
   #End If
#End If

'Refresh the test spreadsheet
Public Sub Refresh()
    If loTests Is Nothing Then
        Me.protInit
    Else
        If Not loTests.DataBodyRange Is Nothing Then
            loTests.DataBodyRange.EntireRow.Delete
        End If
    End If
    dtLastCall = GetTickCount()
End Sub

'Add a Topic to the test log
'@param {String} sName - Name of the topic to enter into the test log
Public Sub Topic(ByVal sName As String)
    If loTests Is Nothing Then Me.protInit
    
    Dim lr As ListRow
    Set lr = loTests.ListRows.Add()
    lr.Range.value = Array("Topic", sName, Empty, Empty)
    If loTests.ListRows.count = 1 Then Call RefreshGlobalFormats
    dtLastCall = GetTickCount() 'Record time
End Sub

'Add a Test to the test log
'@param {String}  sName  -
'@param {Boolean} result -
Public Sub Assert(ByVal sName As String, ByVal result As Boolean)
    If loTests Is Nothing Then Me.protInit
    
    Dim lr As ListRow
    Set lr = loTests.ListRows.Add()
    lr.Range.value = Array("Test", sName, result, GetTickCount() - dtLastCall)
    If loTests.ListRows.count = 1 Then Call RefreshGlobalFormats
    If shTests Is ActiveSheet Then lr.Range.Select
    dtLastCall = GetTickCount() 'Record time
End Sub


'Reset the timer used to test performance of the test suite
'typically we'd advise against calling this method
Public Sub ResetTimer()
  dtLastCall = GetTickCount()
End Sub

'Activate the test sheet
Public Sub Activate()
  If loTests Is Nothing Then Me.protInit
  loTests.Parent.Activate
End Sub


'TestZone
Public Property Get Range() As Range
  Set Range = TestZone.Range("A1")
End Property

'A boolean indicating whether the current running test is a FullTest or a QuickTest.
'Full tests will test everything even things that take a long time to process. Quick tests (FullTest = false) are designed for a quick check while working on the library.
Public Property Get FullTest() As Boolean
  FullTest = pFT
End Property
Public Property Let FullTest(b As Boolean)
  pFT = b
End Property




'Initialise Test
Friend Sub protInit()
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
        Set fc = shTests.Range("$A:$D").FormatConditions.Add(xlExpression, Formula1:="=$A1=""Topic""")
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



'Obtain a sheet, if it doesn't exist, create it.
'@param {String} Name of worksheet to check for. Note: Not to be confused with the codename.
'@returns {Worksheet} Sheet found or created
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

'Check whether a list object exists
'@param {Worksheet} Worksheet to check for a list object on
'@param {String}    Name of list object to check for
'@returns {Boolean} True if list object exists, otherwise false.
Private Function loExists(ByVal ws As Worksheet, ByVal sName As String) As Boolean
    On Error Resume Next
    loExists = Not ws.ListObjects(sName) Is Nothing
End Function

'Refresh the conditional formatting on the test sheet
Private Sub RefreshGlobalFormats()
    Dim fc As FormatCondition
    For Each fc In shTests.Range("A:D").FormatConditions
        If fc.Formula1 = "=$A1=""Topic""" Then
            Call fc.ModifyAppliesToRange(shTests.Range("A:D"))
        End If
    Next
End Sub

'Events
Private Sub shTests_BeforeDelete()
    Set loTests = Nothing
End Sub
