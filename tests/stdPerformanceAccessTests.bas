Attribute VB_Name = "stdPerformanceAccessTests"
Option Explicit

Public Sub testAll()
  Test.Topic "stdPerformance Access"
  Test.Assert "Access host", Application.Name = "Microsoft Access"
  Dim outer As stdPerformance, inner As stdPerformance
  Set outer = stdPerformance.CreateOptimiser()
  Test.Assert "Access creates without Excel properties", Not outer Is Nothing
  Set inner = stdPerformance.CreateOptimiser(True)
  Test.Assert "Access supports nested repainting scopes", Not inner Is Nothing
  Set outer = Nothing
  Set inner = Nothing
  Set outer = stdPerformance.CreateOptimiser(AccessEchoOnExit:=False)
  Set outer = Nothing
  Application.Echo True

  'Echo has no getter; verify nesting/restoration state independently of display timing.
  Dim first As Long, second As Long
  first = stdPerformance.protRegisterAccessEcho(False, True)
  second = stdPerformance.protRegisterAccessEcho(True, False)
  Test.Assert "Removing inner restores outer", Not stdPerformance.protUnregisterAccessEcho(second)
  Test.Assert "Removing last restores declared baseline", stdPerformance.protUnregisterAccessEcho(first)
  first = stdPerformance.protRegisterAccessEcho(False, False)
  second = stdPerformance.protRegisterAccessEcho(True, True)
  Test.Assert "Removing outer retains newer active scope", stdPerformance.protUnregisterAccessEcho(first)
  Test.Assert "Final scope restores originally-off baseline", Not stdPerformance.protUnregisterAccessEcho(second)
  On Error Resume Next
  Call FailInScope
  Dim failure As Long: failure = Err.Number
  Err.Clear
  On Error GoTo 0
  Test.Assert "Original error propagates through optimiser", failure = 5
  'No active scopes should remain after error unwinding.
  first = stdPerformance.protRegisterAccessEcho(False, False)
  Test.Assert "Unwinding releases scope bookkeeping", Not stdPerformance.protUnregisterAccessEcho(first)
  Application.Echo True
End Sub

Private Sub FailInScope()
  With stdPerformance.CreateOptimiser()
    Err.Raise 5, "stdPerformanceAccessTests", "Expected test error"
  End With
End Sub

