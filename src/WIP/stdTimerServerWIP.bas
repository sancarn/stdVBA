Attribute VB_Name = "timerTest"
'CODE GENERATED FOR `stdTimer`

#Const DEBUGGING = True

'GetTickCount
'This returns the number of milliseconds that have elapsed since computer was started.
'This will run for 49 days before resetting back to zero.
#If Mac Then
  #If MAC_OFFICE_VERSION >= 15 Then
    #If VBA7 Then ' 64-bit Excel 2016 for Mac
      Declare PtrSafe Sub Sleep Lib "/usr/lib/libc.dylib" Alias "usleep" (ByVal dwMicroseconds As Long)
      Declare PtrSafe Function GetTickCount Lib "/Applications/Microsoft Excel.app/Contents/Frameworks/MicrosoftOffice.framework/MicrosoftOffice" () As Long
    #Else ' 32-bit Excel 2016 for Mac
      Declare Sub Sleep Lib "/usr/lib/libc.dylib" Alias "usleep" (ByVal dwMicroseconds As Long)
      Declare Function GetTickCount Lib "/Applications/Microsoft Excel.app/Contents/Frameworks/MicrosoftOffice.framework/MicrosoftOffice" () As Long
    #End If
  #Else
    #If VBA7 Then ' does not exist, but why take a chance
      Declare PtrSafe Sub Sleep Lib "/usr/lib/libc.dylib" Alias "usleep" (ByVal dwMicroseconds As Long)
      Declare PtrSafe Function GetTickCount Lib "Applications:Microsoft Office 2011:Office:MicrosoftOffice.framework:MicrosoftOffice" () As Long
    #Else ' 32-bit Excel 2011 for Mac
      Private Declare Sub Sleep Lib "/usr/lib/libc.dylib" Alias "usleep" (ByVal dwMicroseconds As Long)
      Declare Function GetTickCount Lib "Applications:Microsoft Office 2011:Office:MicrosoftOffice.framework:MicrosoftOffice" () As Long
    #End If
  #End If
#Else
  #If VBA7 Then ' Excel 2010 or later for Windows
    Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
  #Else ' pre Excel 2010 for Windows
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Declare Function GetTickCount Lib "kernel32" () As Long
  #End If
#End If

Private Type Timer
  current As Long
  frequency As Long
  ID As String
  active As Boolean
  lastTickCount As Long
End Type
Private Timers() As Timer
Private Const MAX_TIME As Long = 86400000 '24 hours
Private Const MIN_TIME As Long = 10

Sub tt()
    Debug.Print AddTimer(1000)
End Sub

'Add a timer to the schedule
'@param {Long} How often the timer should raise an event
'@returns {String} The ID of the timer.
Public Function AddTimer(ByVal iMilliseconds As Long) As String
  Dim ID As String: ID = getGUID()
  If iMilliseconds > MAX_TIME Then iMilliseconds = MAX_TIME
  If iMilliseconds < MIN_TIME Then iMilliseconds = MIN_TIME
  On Error Resume Next: iNext = UBound(Timers) + 1: On Error GoTo 0
  ReDim Preserve Timers(0 To iNext)
  With Timers(iNext)
    .current = 0
    .frequency = iMilliseconds
    .ID = ID
    .active = True
    .lastTickCount = GetTickCount()
  End With
  AddTimer = ID
  If iNext = 0 Then Application.OnTime Now(), "MainLoop" 'initialise main loop asynchronously
End Function

'Stop a timer
'@param {String} The Timer ID to stop
Public Sub StopTimer(ByVal sID As String)
  Dim i As Long
  For i = 0 To UBound(Timers)
    If Timers(i).ID = sID Then
      Timers(i).active = False
    End If
  Next
End Sub

'The main loop, repeats until no timers are active for 5 minutes. Each timer structure handles the state and frequency.
Public Sub MainLoop()
  Set r = Sheet1.Cells(1, 1)
  Dim bActive As Boolean: bActive = True
  While bActive
    bActive = False
    Dim iTimer As Long
    For iTimer = 0 To UBound(Timers)
      With Timers(iTimer)
        Dim iCurrentTick As Long: iCurrentTick = GetTickCount()
        Dim iDiffMs As Long: iDiffMs = getTickDiff(iCurrentTick, .lastTickCount)
        If iDiffMs < 0 Then
          .current = .frequency
          iDiffMs = 0
        End If
        Debug.Print iDiffMs
        .lastTickCount = iCurrentTick
        .current = .current + iDiffMs
        If .active Then
          bActive = True
          If .current > .frequency Then
            'Trigger a change event
            r.Value = .ID
            
            
            #If DEBUGGING Then
                Debug.Print "--------TICK--------"
                Call SaveSetting("stdVBA", "stdTimer", "last_" & .ID, Now())
            #End If
            
            'If last response was more than 5 minutes ago then. Note: Should not be a concern even with timers longer
            'than 5 minutes, because as soon as the change event is triggered the timer records the setting. This will
            'only not trigger if the application has lost state.
            If DateDiff("n", Now(), CDate(GetSetting("stdVBA", "stdTimer", "last_" & .ID))) > 5 Then
              .active = False
            End If
            
            'Resets counter - saves having to deal with overflows
            .current = 0
          End If
        End If
      End With
    Next
    DoEvents
    'TODO: Do we need this? Sleep 10 'not sure if sleeping here makes any sense... How long does executing this loop take on average? Feels like we should try to keep this in sync somehow...
    Sleep 10
  Wend

  'If timer expires then cleanup and destroy application
  Call SaveSetting("stdVBA", "stdTimer", "instance", 0)
  #If DEBUGGING = 0 Then
    Application.DisplayAlerts = False
    Application.Quit
  #End If
End Sub

'Create and return a new GUID
'@returns {String} A new GUID
Private Function getGUID() As String
  Call Randomize 'Ensure random GUID generated
  getGUID = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx"
  getGUID = Replace(getGUID, "y", Hex(Rnd() And &H3 Or &H8))
  Dim i As Long: For i = 1 To 30
    getGUID = Replace(getGUID, "x", Hex$(Int(Rnd() * 16)), 1, 1)
  Next
End Function

'Obtain the difference in 2 tick counts.
'@param {long} A tick count from `GetTickCount()`
'@param {long} A tick count from `GetTickCount()`, typically expected that this count is the "previous" identified tick count.
'@returns {long} The time difference in ms
'@remark Tries to account for the fact that GetTickCount() might loop from 2^32-1 >> 0.
'@remark Also accounts for the fact VBA uses `Signed Long` instead of `Unsigned Long` (which `GetTickCount()` uses)
Private Function getTickDiff(ByVal iTick As Long, ByVal iPrevTick As Long) As Long
  'Detect when exceeded 49 days (very uncommon but worth checking)
  If iTick > 0 And iPrevTick < 0 Then
    getTickDiff = iTick + Not iPrevTick
  Else
    getTickDiff = BitwiseSubtract(iTick, iPrevTick)
  End If
End Function

'Performs x-y using bitwise operations. This is the equivalent of being able to subtract 2 signed longs
'@param {Long} A) Base value
'@param {Long} B) Value to subtract from base Value
'@returns result of (A) - (B)
Public Function BitwiseSubtract(ByVal x As Long, ByVal y As Long) As Long
  While y <> 0
    Dim borrow As Long: borrow = (Not x) And y
    x = x Xor y
    y = shl(borrow)
  Wend
  BitwiseSubtract = x
End Function

'Left shift
'@param {long} Value - Value to shift
'@param {byte=1} Shift - Number of bits to shift by
Public Function shl(ByVal Value As Long, Optional ByVal Shift As Byte = 1) As Long
    shl = Value
    If Shift > 0 Then
        Dim i As Byte
        Dim m As Long
        For i = 1 To Shift
            m = shl And &H40000000
            shl = (shl And &H3FFFFFFF) * 2
            If m <> 0 Then
                shl = shl Or &H80000000
            End If
        Next i
    End If
End Function

