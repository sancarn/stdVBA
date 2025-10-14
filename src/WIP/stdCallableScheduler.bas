Attribute VB_Name = "stdCallableScheduler"

Private scheduledCallbacks As Collection

'Schedule a callback after a number of seconds
'@param cb - The callback to schedule
'@param seconds - The number of seconds to wait before calling the callback
Public Function ScheduleCallback(ByVal cb As stdICallable, ByVal seconds As Long) as Long
  if scheduledCallbacks is nothing Then Set scheduledCallbacks = New Collection
  Dim onTime As Date: onTime = Now() + TimeSerial(0, 0, 5)
  Call scheduledCallbacks.Add(Array(cb, onTime))
  Call Application.onTime(onTime, "protCallScheduledCallbacks")
  ScheduleCallback = scheduledCallbacks.Count
End Sub

'Call all scheduled callbacks
'@protected
Public Sub protCallScheduledCallbacks()
  Dim vCurTime: vCurTime = Now()

  Dim i As Long
  For i = scheduledCallbacks.Count To 1 Step -1
    Dim cb As stdICallable: Set cb = scheduledCallbacks(i)(0)
    Dim onTime As Date: onTime = scheduledCallbacks(i)(1)
    If DateDiff("s", onTime, vCurTime) >= 0 Then
      Call scheduledCallbacks.Remove(i)
      'Allow the callback to reschedule itself if needed
      Dim iSeconds As Long: iSeconds = cb.Run()
      If iSeconds > 0 Then Call ScheduleCallback(cb, iSeconds)
    End If
  Next
End Sub