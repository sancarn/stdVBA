Attribute VB_Name = "StdRegexBenchmark"
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

  Dim before As stdRegexBaseline, after As stdRegex
  Set before = stdRegexBaseline.Create(pattern, "g")
  Set after = stdRegex.Create(pattern, "g")
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
  Dim iterations(0 To 5) As Long, elapsed As Double, op As Long
  For op = 0 To 5
    elapsed = TimeOperation(op, before, after, text, format, formats, 2)
    iterations(op) = 1
    Do
      elapsed = TimeOperation(op, before, after, text, format, formats, iterations(op))
      If elapsed >= targetSeconds Or iterations(op) >= 65536 Then Exit Do
      iterations(op) = iterations(op) * 2
    Loop
  Next
  Dim result() As Variant, sample As Long, slot As Long, row As Long
  ReDim result(1 To samples * 6, 1 To 9) As Variant
  For sample = 1 To samples
    For slot = 0 To 5
      op = (slot + sample - 1) Mod 6
      elapsed = TimeOperation(op, before, after, text, format, formats, iterations(op))
      row = row + 1
      result(row, 1) = label
      result(row, 2) = nMatches
      result(row, 3) = nCaptures
      result(row, 4) = Len(text)
      result(row, 5) = Choose(op + 1, "List before", "List after", "ListArr before", "ListArr after", "MatchAll", "MatchAllArr")
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

Private Function TimeOperation(ByVal op As Long, ByVal before As stdRegexBaseline, ByVal after As stdRegex, _
    ByRef text As String, ByRef format As String, ByRef formats As Variant, ByVal iterations As Long) As Double
  Dim started As Double, i As Long, values As Variant, names As Object, matches As Collection, output As String
  started = Seconds()
  For i = 1 To iterations
    Select Case op
      Case 0: output = before.List(text, format): output = vbNullString
      Case 1: output = after.List(text, format): output = vbNullString
      Case 2: values = before.ListArr(text, formats): Erase values
      Case 3: values = after.ListArr(text, formats): Erase values
      Case 4: Set matches = after.MatchAll(text): Set matches = Nothing
      Case 5: values = after.MatchAllArr(text, names): Erase values: Set names = Nothing
    End Select
  Next
  TimeOperation = Seconds() - started
End Function

Private Sub ValidateResults(ByVal before As stdRegexBaseline, ByVal after As stdRegex, ByRef text As String, _
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
  If matches.Count <> nMatches Then Err.Raise 5, , "Unexpected match count"
  If nMatches = 0 Then
    If UBound(newValues) <> -1 Then Err.Raise 5, , "Expected empty capture array"
  Else
    If UBound(newValues, 1) <> nMatches Or UBound(newValues, 2) <> nCaptures Then Err.Raise 5, , "Unexpected array bounds"
    For row = 1 To nMatches
      For col = 0 To nCaptures
        If matches(row)(col) <> newValues(row, col) Then Err.Raise 5, , "Capture values differ"
      Next
    Next
  End If
End Sub
