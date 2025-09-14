Attribute VB_Name = "stdCallableScheduler"

Private scheduledCallbacks As Collection

'Schedule a callback after a number of seconds
'@param cb - The callback to schedule
'@param seconds - The number of seconds to wait before calling the callback
Public Sub ScheduleCallback(ByVal cb As stdICallable, ByVal seconds As Long)
  With stdError.getSentry("stdCallableScheduler#ScheduleCallback")
    On Error GoTo stdErrorWrapper_ErrorOccurred:
      if scheduledCallbacks is nothing Then Set scheduledCallbacks = New Collection
      Dim onTime As Date: onTime = Now() + TimeSerial(0, 0, seconds)
      Call scheduledCallbacks.Add(Array(cb, onTime))
      Call Application.onTime(onTime, "protCallScheduledCallbacks")
    Exit Sub
stdErrorWrapper_ErrorOccurred:
      Call stdError.Raise(Err.description)
  End With
End Sub

'Call all scheduled callbacks
'@protected
Public Sub protCallScheduledCallbacks()
  With stdError.getSentry("stdCallableScheduler#protCallScheduledCallbacks")
    On Error GoTo stdErrorWrapper_ErrorOccurred:
      Dim i As Long
      For i = scheduledCallbacks.Count To 1 Step -1
        Dim cb As stdICallable: Set cb = scheduledCallbacks(i)(0)
        Dim onTime As Date: onTime = scheduledCallbacks(i)(1)
        If onTime <= Now() Then
          Call scheduledCallbacks.Remove(i)
          Call cb.Run()
        End If
      Next
    Exit Sub
stdErrorWrapper_ErrorOccurred:
      Call stdError.Raise(Err.description)
  End With
End Sub