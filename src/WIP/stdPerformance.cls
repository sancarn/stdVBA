VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdPerformance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


#If Mac Then
   #If MAC_OFFICE_VERSION >= 15 Then
      Declare Function GetTickCount Lib "/Applications/Microsoft Excel.app/Contents/Frameworks/MicrosoftOffice.framework/MicrosoftOffice" () As Long
   #Else
      Declare Function GetTickCount Lib "Applications:Microsoft Office 2011:Office:MicrosoftOffice.framework:MicrosoftOffice" () As Long
   #End If
#Else ' Win32 or Win64
   #If VBA7 Then
      Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
      Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
   #Else
      Declare Function GetTickCount Lib "kernel32" () As Long
      Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
   #End If
#End If



Private bEnableEvents as boolean
Private bScreenUpdating as boolean
Public oFinishTimes as object
Private oPerformanceDict as object

'Usage:
'With stdPerformance.Optimiser()
'  '...
'End With
Public Function Optimiser() as stdSentry
  Dim perf as stdPerformance
  set perf = new stdPerformance
  perf.OptimiserOn()
  set Optimiser = stdSentry.Create(stdCallback.CreateFromObjectMethod(perf,"OptimiserOff"))
End Function

Public Function OptimiserOn()
  Me.bScreenUpdating = Application.ScreenUpdating
  Me.bEnableEvents = Application.EnableEvents
  Application.ScreenUpdating = false
  Application.EnableEvents = false
End Function
Public Function OptimiserOff()
  Application.ScreenUpdating = Me.bScreenUpdating
  Application.EnableEvents = Me.bEnableEvents
End Function


Public Function Measure(ByVal sMessage as string)
  if Me.oPerformanceDict is nothing then set Me.oPerformanceDict = CreateObject("Scripting.Dictionary")
  Me.oPerformanceDict(sMessage) = getTime()
  
  set Measure = stdSentry.Create(stdCallback.CreateFromObjectMethod(Me,"MeasureStop"),sMessage)
End Function

Public Function MeasureStop(ByVal sMessage as variant)
  #if Mac then

  #else
    if Me.oFinishTimes is nothing then set Me.oFinishTimes = CreateObject("Scripting.Dictionary")
    Dim crStart as double: crStart = Me.oPerformanceDict(sMessage)
    Dim crEnd as double: crEnd = getTime()
    Me.oFinishTimes(sMessage) = 
    Debug.Print 1000# * (crEnd - crStart) / m_crFrequency
  #end if
End Function




Private Sub Class_Initialize()
  Dim PerfFrequency As LARGE_INTEGER
  QueryPerformanceFrequency PerfFrequency
  m_crFrequency = LI2Double(PerfFrequency)
End Sub

Private Function getTime() as double
  #if Mac then
    getTime = Timer()
  #else
    Dim LI as LARGE_INTEGER
    QueryPerformanceCounter LI
    getTime = 1000# * LI2Double(LI) / m_crFrequency
  #end if
End Function

Private Function LI2Double(LI As LARGE_INTEGER) As Double
  Dim Low As Double
    Low = LI.lowpart
    If Low < 0 Then
        Low = Low + TWO_32
    End If
    LI2Double = LI.highpart * TWO_32 + Low
End Function