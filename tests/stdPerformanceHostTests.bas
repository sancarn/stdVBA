Attribute VB_Name = "stdPerformanceHostTests"
Option Explicit

Public Sub testAll()
  Dim app As Object: Set app = Application
  Test.Topic "stdPerformance optimiser host restoration"
  Dim originalScreen As Boolean: originalScreen = app.ScreenUpdating
  Dim originalEvents As Boolean, originalCalculation As Long
  Dim failure As Long, description As String
  On Error GoTo Failed
  If app.Name = "Microsoft Excel" Then
    originalEvents = app.EnableEvents
    originalCalculation = app.Calculation
    app.EnableEvents = True
    app.Calculation = -4105
  End If
  app.ScreenUpdating = True
  With stdPerformance.CreateOptimiser()
    Test.Assert "Repainting disabled", Not app.ScreenUpdating
    If app.Name = "Microsoft Excel" Then
      Test.Assert "Events disabled", Not app.EnableEvents
      Test.Assert "Calculation made manual", app.Calculation = -4135
    End If
    With stdPerformance.CreateOptimiser(True, True, -4105)
      Test.Assert "Nested repainting enabled", app.ScreenUpdating
    End With
    Test.Assert "Nested scope restores outer repainting", Not app.ScreenUpdating
  End With
  Test.Assert "Repainting restored", app.ScreenUpdating
  If app.Name = "Microsoft Excel" Then
    Test.Assert "Events restored", app.EnableEvents
    Test.Assert "Calculation restored", app.Calculation = -4105
  End If
  On Error Resume Next
  Call FailInScope
  failure = Err.Number
  Err.Clear
  On Error GoTo Failed
  Test.Assert "Original error propagates", failure = 5
  Test.Assert "Repainting restored on unwind", app.ScreenUpdating
  app.ScreenUpdating = originalScreen
  If app.Name = "Microsoft Excel" Then
    app.EnableEvents = originalEvents
    app.Calculation = originalCalculation
  End If
  Exit Sub
Failed:
  failure = Err.Number: description = Err.Description
  On Error Resume Next
  app.ScreenUpdating = originalScreen
  If app.Name = "Microsoft Excel" Then
    app.EnableEvents = originalEvents
    app.Calculation = originalCalculation
  End If
  On Error GoTo 0
  Err.Raise failure, "stdPerformanceHostTests", description
End Sub

Private Sub FailInScope()
  With stdPerformance.CreateOptimiser()
    Err.Raise 5, "stdPerformanceHostTests", "Expected test error"
  End With
End Sub

