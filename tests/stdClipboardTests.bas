Attribute VB_Name = "stdClipboardTests"
Sub testAll()
    test.Topic "stdClipboard"
    
    On Error Resume Next
    With xlsProjectBuilder.Range("R1")
        .value = "Test"
        .copy: Call stdClipboard.await
        Test.Assert "Ensure we can get clipboard text", stdClipboard.text = (.value & vbCrLf)
        .copy: Call stdClipboard.await
        Test.Assert "Ensure Range::Copy() supports CF_BITMAP", stdClipboard.IsFormatAvailable(CF_BITMAP)
        .copy: Call stdClipboard.await
        Test.Assert "Ensure IPicture is present", TypeOf stdClipboard.Picture Is stdole.IPicture
        .copy: Call stdClipboard.await
        Test.Assert "Ensure more than 1 format exists", formats.count > 1
        .copy: Call stdClipboard.await
        Test.Assert "Ensure more than 1 formatID exists", formatIDs.count > 1
        .copy: Call stdClipboard.await
        Test.Assert "Ensure number of formats equals the number of formatIDs", formatIDs.count = formats.count
        
        'Can we set the clipboard to text?
        stdClipboard.text = "Hello world"
        xlsProjectBuilder.Paste xlsProjectBuilder.Range("R1")
        Test.Assert "Ensure we can set clipboard text", .value = "Hello world"

        Dim files as collection
        set files = new collection
        files.add "C:\a\b\c"
        files.add "C:\a\b\d"
        set stdClipboard.files = files
        Test.Assert "Ensure setting files ensure clipboard supports CF_HDROP", stdClipboard.IsFormatAvailable(CF_HDROP)
        Test.Assert "Ensure we can get at the files 1", stdClipboard.files.item(1) = "C:\a\b\c"
        Test.Assert "Ensure we can get at the files 2", stdClipboard.files.item(2) = "C:\a\b\d"
        .value = Empty
    End With
End Sub
