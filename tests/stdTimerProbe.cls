VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "stdTimerProbe"
Public TimerInstance As stdTimer
Public Calls As Long
Public Sub Run()
  Calls = Calls + 1
  TimerInstance.StopTimer
  TimerInstance.StopTimer
  Set TimerInstance = Nothing
End Sub
