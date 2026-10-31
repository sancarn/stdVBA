Attribute VB_Name = "StdRegexTestRunner"
Option Explicit
Public regexCount As Long
Public namesCount As Long

Public Function RunTests() As Variant
  On Error GoTo Failed
  Test.checks = 0
  stdDateTests.testAll
  RunTests = Test.checks
  Exit Function
Failed:
  RunTests = "VBA error " & Err.Number & ": " & Err.Description & " (last check: " & Test.lastCheck & ")"
End Function
