Attribute VB_Name = "stdTimerTests"
Option Explicit

'Run explicitly: creates isolated Excel hosts and requires existing AccessVBOM trust.
Public Function testLifecycle() As String
  Dim t As stdTimer, i As Long, count As Long
  count = Application.Workbooks.Count
  For i = 1 To 5
    Set t = stdTimer.Create(10)
    t.StopTimer
    t.StopTimer
    Set t = Nothing
  Next
  Set t = New stdTimer
  t.StopTimer
  t.StopTimer
  Set t = Nothing
  For i = 1 To 3
    Set t = stdTimer.Create(10)
    Set t = Nothing
  Next
  Dim probe As stdTimerProbe: Set probe = New stdTimerProbe
  Set probe.TimerInstance = stdTimer.Create(10, probe)
  Dim deadline As Date: deadline = DateAdd("s", 10, Now)
  Do While probe.Calls = 0 And Now < deadline
    DoEvents
  Loop
  If Not probe.TimerInstance Is Nothing Then probe.TimerInstance.StopTimer
  If probe.Calls <> 1 Then Err.Raise 5, , "Callback shutdown failed"
  If Application.Workbooks.Count <> count Then Err.Raise 5, , "Caller workbooks changed"
  testLifecycle = "PASS: repeated stop, release, callback stop, caller workbooks"
End Function

Public Function testInvalidResolution() As Boolean
  On Error GoTo Expected
  Dim t As stdTimer: Set t = stdTimer.Create(0)
  t.StopTimer
  Exit Function
Expected:
  testInvalidResolution = (Err.Number = 5)
End Function
