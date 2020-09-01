VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdPerfTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long

Private m_CounterStart As LARGE_INTEGER
Private m_CounterEnd As LARGE_INTEGER
Private m_crFrequency As Double
Private pName As String

Private Const TWO_32 = 4294967296# ' = 256# * 256# * 256# * 256#

Public Function Create(ByVal sName As String) As stdPerfTest
    Set Create = New stdPerfTest
    Create.Init (sName)
End Function
Public Sub Init(ByVal sName As String)
    pName = sName
    QueryPerformanceCounter m_CounterStart
End Sub

Private Sub Class_Terminate()
    Debug.Print pName & ": " & TimeElapsed()
End Sub



Private Function LI2Double(LI As LARGE_INTEGER) As Double
    Dim Low As Double
    Low = LI.lowpart
    If Low < 0 Then
        Low = Low + TWO_32
    End If
    LI2Double = LI.highpart * TWO_32 + Low
End Function

Private Sub Class_Initialize()
    Dim PerfFrequency As LARGE_INTEGER
    QueryPerformanceFrequency PerfFrequency
    m_crFrequency = LI2Double(PerfFrequency)
End Sub

Property Get TimeElapsed() As Double
Dim crStart As Double
Dim crStop As Double
    QueryPerformanceCounter m_CounterEnd
    crStart = LI2Double(m_CounterStart)
    crStop = LI2Double(m_CounterEnd)
    TimeElapsed = 1000# * (crStop - crStart) / m_crFrequency
End Property

