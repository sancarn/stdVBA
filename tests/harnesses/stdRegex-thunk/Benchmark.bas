Attribute VB_Name = "ThunkBenchmark"
Option Explicit

#If VBA7 Then
  Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
  Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#Else
  Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
  Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#End If

Private Function Seconds() As Double
  Static frequency As Currency
  Dim value As Currency
  If frequency = 0 Then QueryPerformanceFrequency frequency
  QueryPerformanceCounter value
  Seconds = value / frequency
End Function

Public Function HostDetails() As String
#If Win64 Then
  HostDetails = "64-bit VBA"
#Else
  HostDetails = "32-bit VBA"
#End If
End Function

Public Function RunCase(ByVal caseId As Long, ByVal samples As Long, ByVal targetSeconds As Double) As Variant
  On Error GoTo Failed
  Dim pattern As String, text As String, label As String, piece As String
  Dim nMatches As Long, nCaptures As Long, i As Long, padding As String
  pattern = "(?<id>\d{5}-[A-Z]{2}\d) (?<count>\d+) (?<date>\d{2}/\d{2}/\d{4})"
  piece = "12345-AA1 123 10/02/2019" & vbCrLf
  nCaptures = 3
  Select Case caseId
    Case 0: label = "No matches, 1,000 characters": nMatches = 0: text = String$(1000, "x")
    Case 1: label = "1 match, 3 named captures": nMatches = 1
    Case 2: label = "100 matches, 3 named captures": nMatches = 100
    Case 3: label = "1,000 matches, 3 named captures": nMatches = 1000
    Case 4
      label = "100 matches, no captures": nMatches = 100: nCaptures = 0
      pattern = "\d{5}-[A-Z]{2}\d"
    Case 5
      label = "100 sparse matches, 3 named captures": nMatches = 100
      padding = String$(200, "x") & vbCrLf
    Case Else: Err.Raise 5, , "Unknown benchmark case"
  End Select
  For i = 1 To nMatches
    text = text & padding & piece
  Next

  Dim before As stdRegex, after As stdRegex
  Set before = stdRegex.Create(pattern, "g", useNative:=False)
  Set after = stdRegex.Create(pattern, "g", useNative:=True)
  Dim format As String, formats As Variant
  If nCaptures > 0 Then
    format = "$id,$date,$count\r\n"
    formats = Array("$id-$date", "$count")
  Else
    format = "$&\r\n"
    formats = Array("$&", "[$&]")
  End If
  ValidateResults before, after, text, format, formats, nMatches, nCaptures

  'Warm each path and calibrate batching independently. Compilation, validation,
  'and warm-up are outside the reported times.
  Dim iterations(0 To 7) As Long, elapsed As Double, op As Long
  For op = 0 To 7
    elapsed = TimeOperation(op, before, after, text, format, formats, 2)
    iterations(op) = 1
    Do
      elapsed = TimeOperation(op, before, after, text, format, formats, iterations(op))
      If elapsed >= targetSeconds Or iterations(op) >= 65536 Then Exit Do
      iterations(op) = iterations(op) * 2
    Loop
  Next
  Dim result() As Variant, sample As Long, slot As Long, row As Long
  ReDim result(1 To samples * 8, 1 To 9) As Variant
  For sample = 1 To samples
    For slot = 0 To 7
      op = (slot + sample - 1) Mod 8
      elapsed = TimeOperation(op, before, after, text, format, formats, iterations(op))
      row = row + 1
      result(row, 1) = label
      result(row, 2) = nMatches
      result(row, 3) = nCaptures
      result(row, 4) = Len(text)
      result(row, 5) = Choose(op + 1, "List VBA", "List native", "ListArr VBA", "ListArr native", "MatchAllArr VBA", "MatchAllArr native", "Test VBA", "Test native")
      result(row, 6) = sample
      result(row, 7) = iterations(op)
      result(row, 8) = elapsed * 1000
      result(row, 9) = elapsed * 1000 / iterations(op)
    Next
  Next
  RunCase = result
  Exit Function
Failed:
  RunCase = "VBA error " & Err.Number & ": " & Err.Description
End Function

Private Function TimeOperation(ByVal op As Long, ByVal before As stdRegex, ByVal after As stdRegex, _
    ByRef text As String, ByRef format As String, ByRef formats As Variant, ByVal iterations As Long) As Double
  Dim started As Double, i As Long, values As Variant, names As Object, matches As Collection, output As String, found As Boolean
  started = Seconds()
  For i = 1 To iterations
    Select Case op
      Case 0: output = before.List(text, format): output = vbNullString
      Case 1: output = after.List(text, format): output = vbNullString
      Case 2: values = before.ListArr(text, formats): Erase values
      Case 3: values = after.ListArr(text, formats): Erase values
      Case 4: values = before.MatchAllArr(text, names): Erase values: Set names = Nothing
      Case 5: values = after.MatchAllArr(text, names): Erase values: Set names = Nothing
      Case 6: found = before.Test(text)
      Case 7: found = after.Test(text)
    End Select
  Next
  TimeOperation = Seconds() - started
End Function

Private Sub ValidateResults(ByVal before As stdRegex, ByVal after As stdRegex, ByRef text As String, _
    ByRef format As String, ByRef formats As Variant, ByVal nMatches As Long, ByVal nCaptures As Long)
  If before.List(text, format) <> after.List(text, format) Then Err.Raise 5, , "List outputs differ"
  Dim oldValues As Variant, newValues As Variant, row As Long, col As Long
  oldValues = before.ListArr(text, formats)
  newValues = after.ListArr(text, formats)
  If UBound(oldValues, 1) <> UBound(newValues, 1) Then Err.Raise 5, , "ListArr row counts differ"
  For row = 1 To nMatches
    For col = 1 To 2
      If oldValues(row, col) <> newValues(row, col) Then Err.Raise 5, , "ListArr outputs differ"
    Next
  Next
  Dim matches As Collection, names As Object
  Set matches = after.MatchAll(text)
  newValues = after.MatchAllArr(text, names)
  oldValues = before.MatchAllArr(text)
  If matches.Count <> nMatches Then Err.Raise 5, , "Unexpected match count"
  If nMatches = 0 Then
    If UBound(newValues) <> -1 Then Err.Raise 5, , "Expected empty capture array"
  Else
    If UBound(newValues, 1) <> nMatches Or UBound(newValues, 2) <> nCaptures Then Err.Raise 5, , "Unexpected array bounds"
    For row = 1 To nMatches
      For col = 0 To nCaptures
        If oldValues(row, col) <> newValues(row, col) Then Err.Raise 5, , "Baseline/native capture values differ"
        If matches(row)(col) <> newValues(row, col) Then Err.Raise 5, , "Capture values differ"
      Next
    Next
  End If
End Sub

Public Function Differential() As Variant
  On Error GoTo Failed
  Dim patterns As Variant, texts As Variant, flags As Variant, p As Variant, t As Variant, f As Variant
  patterns = Array("", "a", "a|bc", "a*", "a+?", "a{2,4}", "a{2,4}?", "(ab){2}", "(a|b)*", "(a*)*", "(?<x>a+)(b)?", "(?=a)a", "(?!b)a", "(?<=a)b", "(?<!a)b", "(?<=(a+))b", "(a+)b\1", "[a-z]+", "[^a-z]+", "\b\w+\b", "\B", "^a.$", "(?i:a)", "(?-i:a)", "(?s:.)", "(?m:^b)", "(?<x>a)|(?<x>b)", "(a){0}b", "(a?){2,4}", "(?:ab|a)+b", "((a)(b))?", "(?=(a+))a*b\1")
  texts = Array("", "a", "b", "ab", "aaabaaa", "abab", "bbbb", "a" & vbCrLf & "b", "ABab", "123 _ z", ChrW(228) & ChrW(196), ChrW(&H2028), ChrW(0) & "a", "aaaaaaaaab")
  flags = Array("", "i", "m", "s", "ims")
  Dim before As stdRegex, after As stdRegex, count As Long
  Dim oldResult As String, newResult As String, e1 As Long, e2 As Long
  For Each p In patterns
    For Each f In flags
      Set before = stdRegex.Create(CStr(p), CStr(f), useNative:=False)
      Set after = stdRegex.Create(CStr(p), CStr(f), useNative:=True)
      For Each t In texts
        e1 = 0: e2 = 0
        On Error Resume Next
        Err.Clear: oldResult = before.List(CStr(t), "[$&][$1][$2][$x]"): e1 = Err.Number
        Err.Clear: newResult = after.List(CStr(t), "[$&][$1][$2][$x]"): e2 = Err.Number
        On Error GoTo Failed
        If e1 <> e2 Then Err.Raise 5, , "Error mismatch: " & p & " flags=" & f
        If e1 = 0 And oldResult <> newResult Then Err.Raise 5, , "Output mismatch: " & p & " flags=" & f & " text=" & t & " baseline=" & oldResult & " native=" & newResult
        count = count + 1
      Next
    Next
  Next
  'Force the native workspace to grow, with nested capture/quantifier stacks.
  patterns = Array("(a)+$", "(a?){500}b", "(?=(a+))a+$", "(a|aa)+b")
  texts = Array(String$(1000, "a"), String$(500, "a") & "b")
  For Each p In patterns
    Set before = stdRegex.Create(CStr(p), "", 100000, useNative:=False)
    Set after = stdRegex.Create(CStr(p), "", 100000, useNative:=True)
    For Each t In texts
      On Error Resume Next
      Err.Clear: oldResult = before.List(CStr(t), "[$&][$1][$2]"): e1 = Err.Number
      Err.Clear: newResult = after.List(CStr(t), "[$&][$1][$2]"): e2 = Err.Number
      On Error GoTo Failed
      If e1 <> e2 Then Err.Raise 5, , "Workspace/limit error mismatch: " & p
      If e1 = 0 And oldResult <> newResult Then Err.Raise 5, , "Workspace output mismatch: " & p
      count = count + 1
    Next
  Next
  Set before = stdRegex.Create("selection-check", useNative:=False)
  Set after = stdRegex.Create("selection-check", useNative:=True)
  If before Is after Then Err.Raise 5, , "Engine cache collision"
  Dim again As stdRegex
  Set again = stdRegex.Create("selection-check", useNative:=False)
  If Not again Is before Then Err.Raise 5, , "Default engine cache mismatch"
  Set again = stdRegex.Create("selection-check")
  If Not again Is after Then Err.Raise 5, , "Default must select native"
  Differential = count + 3
  Exit Function
Failed:
  Differential = "VBA error " & Err.Number & ": " & Err.Description
End Function
