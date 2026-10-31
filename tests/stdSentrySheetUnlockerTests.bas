Attribute VB_Name = "stdSentrySheetUnlockerTests"
Option Explicit

Public Sub testAll()
  Test.Topic "stdSentry sheet unlocker"
  Dim app As Object, wb As Object, sheet As Object, guard As stdSentry
  Dim failure As Long, description As String
  On Error GoTo Failed
  Set app = CreateObject("Excel.Application")
  Set wb = app.Workbooks.Add
  Set sheet = wb.Worksheets(1)

  With stdSentry.CreateSheetUnlocker(sheet)
    Test.Assert "Unprotected worksheet stays unlocked", Not sheet.ProtectContents
  End With
  Test.Assert "Unprotected worksheet is not protected on exit", Not sheet.ProtectContents

  sheet.Protect "secret", True, True, False, True, True, False, True, False, True, False, False, False, True, True, False
  sheet.EnableSelection = 1 'xlUnlockedCells
  With stdSentry.CreateSheetUnlocker(sheet, "secret")
    Test.Assert "Protected sheet is unlocked", Not sheet.ProtectContents And Not sheet.ProtectDrawingObjects
    sheet.Range("A1").Value = "edited"
    With stdSentry.CreateUnlocker(sheet, "secret")
      Test.Assert "Nested alias remains unlocked", Not sheet.ProtectContents
    End With
    Test.Assert "Inner sentry leaves outer scope unlocked", Not sheet.ProtectContents
  End With
  Test.Assert "Original content and drawing protection restored", sheet.ProtectContents And sheet.ProtectDrawingObjects
  Test.Assert "Original scenario protection restored", Not sheet.ProtectScenarios
  Test.Assert "UserInterfaceOnly restored", sheet.ProtectionMode
  Test.Assert "Selection mode restored", sheet.EnableSelection = 1
  With sheet.Protection
    Test.Assert "Formatting permissions restored", .AllowFormattingCells And Not .AllowFormattingColumns And .AllowFormattingRows
    Test.Assert "Insert permissions restored", Not .AllowInsertingColumns And .AllowInsertingRows And Not .AllowInsertingHyperlinks
    Test.Assert "Delete permissions restored", Not .AllowDeletingColumns And Not .AllowDeletingRows
    Test.Assert "Sort/filter/pivot permissions restored", .AllowSorting And .AllowFiltering And Not .AllowUsingPivotTables
  End With
  Test.Assert "Edits persist", sheet.Range("A1").Value = "edited"

  On Error Resume Next
  Set guard = stdSentry.CreateSheetUnlocker(sheet, "wrong")
  failure = Err.Number
  Err.Clear
  On Error GoTo Failed
  Test.Assert "Wrong password fails at construction", failure <> 0 And guard Is Nothing
  Test.Assert "Wrong password leaves protection intact", sheet.ProtectContents

  On Error Resume Next
  Call FailInScope(sheet)
  failure = Err.Number
  Err.Clear
  On Error GoTo Failed
  Test.Assert "Caller receives original error", failure = 5
  Test.Assert "Protection restored while unwinding error", sheet.ProtectContents And sheet.ProtectionMode

  sheet.Unprotect "secret"
  sheet.Protect "secret", False, False, True
  With stdSentry.CreateSheetUnlocker(sheet, "secret")
    Test.Assert "Scenario-only protection is unlocked", Not sheet.ProtectScenarios
  End With
  Test.Assert "Scenario-only protection restored exactly", sheet.ProtectScenarios And Not sheet.ProtectContents And Not sheet.ProtectDrawingObjects
  wb.Close False
  app.Quit
  Exit Sub
Failed:
  failure = Err.Number: description = Err.Description
  On Error Resume Next
  If Not wb Is Nothing Then wb.Close False
  If Not app Is Nothing Then app.Quit
  On Error GoTo 0
  Err.Raise failure, "stdSentrySheetUnlockerTests", description
End Sub

Private Sub FailInScope(ByVal sheet As Object)
  With stdSentry.CreateSheetUnlocker(sheet, "secret")
    Err.Raise 5, "stdSentrySheetUnlockerTests", "Expected test error"
  End With
End Sub

