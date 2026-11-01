Attribute VB_Name = "StdRegexTestRunner"
Option Explicit

Public Function RunTests() As Variant
  On Error GoTo Failed
  Test.checks = 0
  stdRegexTests.testAll
  RunTests = Test.checks
  Exit Function
Failed:
  RunTests = "VBA error " & Err.Number & ": " & Err.Description & " (last check: " & Test.lastCheck & ")"
End Function
