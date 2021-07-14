VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xlsProjectBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Rebuild and test the project
Public Sub ProjectRebuildAndQuickTest()
  Call ProjectBuilderRebuild
  Call ProjectBuilderQuickTest
End Sub
'Rebuild and test the project
Public Sub ProjectRebuildAndFullTest()
  Call ProjectBuilderRebuild
  Call ProjectBuilderFullTest
End Sub

'Rebuild the project
Public Sub ProjectBuilderRebuild()
  Call ProjectBuilderDeleteAll
  Call ProjectBuilderBuildAll
End Sub


'Delete all modules from the project
Public Sub ProjectBuilderDeleteAll()
  'Obtain vb project
  Dim loVBProject As ListObject: Set loVBProject = Me.ListObjects("VBProject")
  
  Dim iLR As Long
  For iLR = loVBProject.ListRows.count To 1 Step -1
    Dim lr As ListRow: Set lr = loVBProject.ListRows.item(iLR)
    Dim vRow As Variant: vRow = lr.Range.value
    Call RemoveVBComponent(vRow(1, 2))
    lr.Delete
  Next
End Sub

'Build all the modules from the project
Public Sub ProjectBuilderBuildAll()
  'Get roots:
  Dim vRoots As Variant: vRoots = xlsProjectBuilder.ListObjects("Roots").DataBodyRange.value
  Dim iRow As Long
  For iRow = 1 To UBound(vRoots, 1)
    Dim sRootPath As String: sRootPath = ThisWorkbook.path & "\" & vRoots(iRow, 2)
    Dim loVBProject As ListObject: Set loVBProject = Me.ListObjects("VBProject")
    
    Dim vFile: For Each vFile In getFilesFromFolder(sRootPath, vRoots(iRow, 3))
      'Ensure file is a vb file
      If isVBFile(vFile) Then
        'Import component
        With ImportVBComponent(vFile)
          'Add to components table
          loVBProject.ListRows.Add().Range.value = Array(vFile, .name)
        End With
      End If
    Next
  Next
End Sub

'****************************
'*     HELPER FUNCTIONS     *
'****************************

'Run a quick test
Private Sub ProjectBuilderQuickTest()
  Test.Refresh
  Test.FullTest = False
  Test.Activate
  Application.Run Me.Range("PROJBUILDER_TESTFUNC").value
End Sub

'Run a full test
Private Sub ProjectBuilderFullTest()
  Test.Refresh
  Test.FullTest = True
  Test.Activate
  Application.Run Me.Range("PROJBUILDER_TESTFUNC").value
End Sub


'Obtain all the files in a folder, with an optional recurse param
'@param {String}  sDirPath - The path to the directory where files should be obtained from
'@param {Boolean} Recurse  - Whether to recurse or not. Recurse will obtain all files in all child folders as well.
Private Function getFilesFromFolder(ByVal sDirPath As String, Optional ByVal recurse As Boolean) As Collection
  Set getFilesFromFolder = New Collection
  Dim vFile As Variant, vDir As Variant, cFiles As Collection, cFolders As Collection
  If getFilesAndFolders(sDirPath, cFiles, cFolders) Then
    If recurse Then
      For Each vDir In cFolders
        For Each vFile In getFilesFromFolder(vDir, recurse)
          getFilesFromFolder.Add vFile
        Next
      Next
    End If
    For Each vFile In cFiles
      getFilesFromFolder.Add vFile
    Next
  End If
End Function

'Obtain all the files and folders within a folder. This function does not recurse.
'@param {String} Path of the folder to find children in
'@param {ByRef Collection} cFiles   - Files in folder
'@param {ByRef Collection} cFolders - Folders in Folder
'@returns {Boolean} True if children were found, otherwise false.
Private Function getFilesAndFolders(ByVal sFolderPath As String, ByRef cFiles As Collection, ByRef cFolders As Collection) As Boolean
  Set cFiles = New Collection
  Set cFolders = New Collection
  
  Dim sFileOrDir As String
  sFileOrDir = Dir(sFolderPath & "\*", vbDirectory)
  While sFileOrDir <> ""
    If sFileOrDir <> "." And sFileOrDir <> ".." Then
      Dim sFullPath As String: sFullPath = sFolderPath & "\" & sFileOrDir
      If FileSystem.GetAttr(sFullPath) = vbDirectory Then
        Call cFolders.Add(sFullPath)
      Else
        Call cFiles.Add(sFullPath)
      End If
      getFilesAndFolders = True
    End If
    sFileOrDir = Dir()
  Wend
End Function

'Remove a VBComponent by name
'@param {String} The name of the component to remove
Private Sub RemoveVBComponent(ByVal sName As String)
    If VBEIsTrusted() Then
      Set vbc = getVBComponent(sName)
      If Not vbc Is Nothing Then
          Call ThisWorkbook.VBProject.VBComponents.Remove(vbc)
      End If
    Else
      Err.Raise 1, "", "This tool requires VBEExtensibility"
    End If
End Sub

'Obtain a VBComponent by name
'@param {String} The name of the component to retrieve
'@returns {VBComponent} The VBComponent requested
Private Function getVBComponent(ByVal sName As String) As VBComponent
    On Error GoTo ErrorOccurred
    Set getVBComponent = ThisWorkbook.VBProject.VBComponents.item(sName)
    Exit Function
ErrorOccurred:
    Set getVBComponent = Nothing
End Function

'Import a VBComponent by path
'@param {String} The path of the component to import
'@returns {VBComponent} The imported vb component.
Private Function ImportVBComponent(ByVal sPath As String) As VBComponent
    If VBEIsTrusted() Then
      On Error GoTo ErrorOccurred
      Set ImportVBComponent = ThisWorkbook.VBProject.VBComponents.Import(sPath)
      Exit Function
    Else
      Err.Raise 1, "", "This tool requires VBEExtensibility"
    End If
ErrorOccurred:
    Set ImportVBComponent = Nothing
End Function

'Check if a file is a vbFile or not
'@param {String} Path of string to check
'@returns {Boolean} True if file is a VB importable file, false otherwise.
Private Function isVBFile(ByVal sPath As String) As Boolean
    Select Case LCase(right(sPath, 4))
        Case ".bas", ".cls", ".frm"
            isVBFile = True
        Case Else
            isVBFile = False
    End Select
End Function

'Check if VBE Extensibility is trusted
Private Function VBEIsTrusted() As Boolean
    Dim a1 As Integer
    On Error GoTo Label1
    a1 = ActiveWorkbook.VBProject.VBComponents.count
    VBEIsTrusted = True
    Exit Function
Label1:
    VBEIsTrusted = False
End Function
