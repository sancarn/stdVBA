VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdExcelWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@module
'@description
'@ai AI:0
'@aidesc AI has not touched this code yet. This may change to AI:W when AI edits this file.



Public Function CreateFromFile(ByVal sPath as string) as stdExcelWorkbook
  Dim WB as workbook
  if not stdFile.exists(sPath) then
    set WB = Application.Workbooks.add()
    WB.SaveAs sPath
    CreateFromFile = WB
    Exit Function
  end if
  
  On Error GoTo ErrorOccurred
    'Check if workbook is already open...
    For Each WB In Application.Workbooks
      If WB.FullName = sPath Then
        alreadyOpen = True
        Set WorkbookOpen = WB
        Exit Function
      End If
    Next
    
    'Open workbook
    alreadyOpen = False
    Set CreateFromFile = Workbooks.Open(sPath, False, True, CorruptLoad:=XlCorruptLoad.xlRepairFile)
  On Error GoTo 0
  Exit Function
ErrorOccurred:
  'Seems to be a bug in excel where error 1004 (application or object defined error) is thrown when CorruptLoad:=xlCorruptLoad.xlRepairFile
  If Err.Number = 1004 Then
    Set CreateFromFile = Workbooks.Open(sPath, False, True)
  End If
  Set CreateFromFile = Nothing
End Function


