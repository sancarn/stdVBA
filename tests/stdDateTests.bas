Attribute VB_Name = "stdDateTests"
Option Explicit

Public Sub testAll()
  Dim d As stdDate, country As Variant, day As Long, first As Long, dt As Date
  Dim resolved As Long, abbreviated As Variant, result As Object, n As Long
  For Each country In Array("US", "UK")
    For day = 25 To 31
      dt = DateSerial(2026, 10, day)
      Set d = stdDate.Create(dt)
      d.Country = CStr(country)
      For first = 0 To 7
        resolved = first
        If first = 0 Then resolved = IIf(country = "US", vbSunday, vbMonday)
        Test.Assert "numeric weekday", d.Weekday(first) = VBA.Weekday(dt, resolved)
        For Each abbreviated In Array(False, True)
          Test.Assert "weekday name type", VarType(d.WeekdayName(CBool(abbreviated), first)) = vbString
          Test.Assert "weekday name", d.WeekdayName(CBool(abbreviated), first) = VBA.WeekdayName(VBA.Weekday(dt, resolved), CBool(abbreviated), resolved)
        Next
      Next
      Test.Assert "omitted first day", d.Weekday = d.Weekday(0)
      Test.Assert "omitted name first day", d.WeekdayName = d.WeekdayName(False, 0)
    Next
  Next
  For n = 1 To 20
    Set result = stdDate.RegexMatch("AbC", "(?<v>abc)", "i")
    Test.Assert "repeated named capture", result("v") = "AbC"
    Set result = stdDate.RegexMatch("abc", "(?<v>abc)", "")
    Test.Assert "switch options", result(1) = "abc"
    Test.Assert "reuse two keys", StdRegexTestRunner.regexCount = 2 And StdRegexTestRunner.namesCount = 2
  Next
  Set result = stdDate.RegexMatch("a" & vbLf & "b", "(?<v>a.b)", "s")
  Test.Assert "dotall", result("v") = "a" & vbLf & "b"
  Set result = stdDate.RegexMatch("a" & vbLf & "b", "^(?<v>b)", "m")
  Test.Assert "multiline", result("v") = "b"
  Set result = stdDate.RegexMatch("ab", "(?<v>a b)", "xg")
  Test.Assert "free space and global", result("v") = "ab"
  Test.Assert "no match error", RegexError("ABC", "(?<v>abc)", "") = 5
  Dim beforeRegex As Long, beforeNames As Long
  beforeRegex = StdRegexTestRunner.regexCount
  beforeNames = StdRegexTestRunner.namesCount
  For n = 1 To 100
    Test.Assert "malformed pattern error", RegexError("abc", "([" & n, "i") = 5019
  Next
  Test.Assert "repeated malformed error", RegexError("abc", "([1", "i") = 5019
  Set result = stdDate.RegexMatch("AbC", "(?<v>abc)", "i")
  Test.Assert "valid after errors", result("v") = "AbC"
  Test.Assert "failed keys not retained", StdRegexTestRunner.regexCount = beforeRegex And StdRegexTestRunner.namesCount = beforeNames
End Sub

Private Function RegexError(ByVal value As String, ByVal pattern As String, ByVal options As String) As Long
  On Error Resume Next
  Dim result As Object
  Set result = stdDate.RegexMatch(value, pattern, options)
  RegexError = Err.Number
  On Error GoTo 0
End Function
