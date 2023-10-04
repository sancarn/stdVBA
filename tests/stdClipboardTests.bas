Attribute VB_Name = "stdClipboardTests"
'@lang VBA

Sub testAll()
    Test.Topic "stdClipboard"
    
    'Test.Assert() currently removes the clipboard when it adds a row to a table (during `ListRows.add()`)
    'So in this case we'll store all tests in an array, and then loop over the array and test results after running
    'these tests
    Dim vbaTests(1 To 99, 1 To 2) As Variant
    Dim index As Long: index = 0
    
    On Error GoTo 0
    With Test.Range.Cells(1, 1)
        
        .value = "Test"
        
        'Can we extract the value of the clipboard?
        .copy
        AppendTest vbaTests, index, "Ensure we can get clipboard text", stdClipboard.text = (.value & vbCrLf)
        AppendTest vbaTests, index, "Ensure Range::Copy() supports CF_BITMAP", stdClipboard.IsFormatAvailable(CF_BITMAP)
        AppendTest vbaTests, index, "Ensure IPicture is present", TypeOf stdClipboard.Picture Is stdole.IPicture
        AppendTest vbaTests, index, "Ensure more than 1 format exists", stdClipboard.formats.Count > 1
        AppendTest vbaTests, index, "Ensure more than 1 formatID exists", stdClipboard.formatIDs.Count > 1
        AppendTest vbaTests, index, "Ensure number of formats equals the number of formatIDs", stdClipboard.formatIDs.Count = stdClipboard.formats.Count
        
        'Can we set the clipboard to text?
        stdClipboard.text = "Hello world"
        xlsProjectBuilder.Paste Test.Range.Cells(1, 1)
        AppendTest vbaTests, index, "Ensure we can set clipboard text", .value = "Hello world"
        .value = Empty
    End With
    
    'Test files copy
    #If Win64 Then
        'Currently got a crash here, known limitation - not sure why this is occurring.
        AppendTest vbaTests, index, "NOT x64: Ensure setting files ensure clipboard supports CF_HDROP", False
        AppendTest vbaTests, index, "NOT x64: Ensure we can get at the files 1", False
        AppendTest vbaTests, index, "NOT x64: Ensure we can get at the files 2", False
    #Else
        Dim files As Collection
        Set files = New Collection
        files.Add "D:\Programming\Github\VBA-STD-Library\README.md"
        files.Add "D:\Programming\Github\VBA-STD-Library\Links.md"
        Set stdClipboard.files = files
        AppendTest vbaTests, index, "Ensure setting files ensure clipboard supports CF_HDROP", stdClipboard.IsFormatAvailable(CF_HDROP)
        AppendTest vbaTests, index, "Ensure we can get at the files 1", stdClipboard.files.item(1) = "C:\a\b\c"
        AppendTest vbaTests, index, "Ensure we can get at the files 2", stdClipboard.files.item(2) = "C:\a\b\d"
    #End If
    
    'Loop through tests and log
    Dim iTest As Long
    For iTest = LBound(vbaTests, 1) To UBound(vbaTests, 1)
        If vbaTests(iTest, 1) = Empty Then Exit For
        Test.Assert vbaTests(iTest, 1), vbaTests(iTest, 2)
    Next iTest
End Sub

Sub testFileCrash()
    'Test files copy
    Dim files As Collection
    Set files = New Collection
    files.Add "D:\Programming\Github\VBA-STD-Library\README.md"
    files.Add "D:\Programming\Github\VBA-STD-Library\Links.md"
    Set stdClipboard.files = files
End Sub


Private Sub AppendTest(ByRef vbaTests As Variant, ByRef testIndex As Long, ByVal sMsg As String, ByVal bResult As Boolean)
    testIndex = testIndex + 1
    vbaTests(testIndex, 1) = sMsg
    vbaTests(testIndex, 2) = bResult
End Sub
