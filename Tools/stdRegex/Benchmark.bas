Attribute VB_Name = "RegexBenchmark"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#Else
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#End If

Private pattern As String, flags As String, subject As String, replacement As String
Private expectedMatches As Collection, expectedReplacement As String
Private engine As Object, engineName As String
Private checksum As Long

Private Function Seconds() As Double
    Static frequency As Currency
    Dim counter As Currency
    If frequency = 0 Then QueryPerformanceFrequency frequency
    QueryPerformanceCounter counter
    Seconds = counter / frequency
End Function

Private Function RepeatText(ByVal text As String, ByVal count As Long) As String
    Dim i As Long
    For i = 1 To count
        RepeatText = RepeatText & text
    Next
End Function

Private Sub Configure(ByVal scenario As String)
    Set expectedMatches = New Collection
    flags = "g"
    replacement = "x"
    Dim i As Long
    Select Case scenario
        Case "Literal scan"
            pattern = "needle"
            subject = String$(1024, "x") & "needle" & String$(1024, "x")
            expectedMatches.Add Array("needle")
            expectedReplacement = String$(1024, "x") & "x" & String$(1024, "x")
        Case "Short validation"
            pattern = "^[A-Z]{2}\d{6}$"
            subject = "GB123456"
            expectedMatches.Add Array("GB123456")
            expectedReplacement = "x"
        Case "Capture extraction"
            pattern = "([A-Z]{3})-(\d{4})"
            subject = RepeatText("ABC-1234; ", 100)
            replacement = "$2/$1"
            For i = 1 To 100: expectedMatches.Add Array("ABC-1234", "ABC", "1234"): Next
            expectedReplacement = RepeatText("1234/ABC; ", 100)
        Case "Alternation and ignore case"
            pattern = "(ERROR|WARN|INFO): ([A-Za-z]+)"
            flags = "gi"
            subject = RepeatText("info: Ready; WARN: Retry; ", 50)
            replacement = "$2"
            For i = 1 To 50
                expectedMatches.Add Array("info: Ready", "info", "Ready")
                expectedMatches.Add Array("WARN: Retry", "WARN", "Retry")
            Next
            expectedReplacement = RepeatText("Ready; Retry; ", 50)
        Case "Multiline anchors"
            pattern = "^([A-Z]+):(\d+)$"
            flags = "gm"
            subject = RepeatText("ITEM:123" & vbLf, 100)
            replacement = "$2/$1"
            For i = 1 To 100: expectedMatches.Add Array("ITEM:123", "ITEM", "123"): Next
            expectedReplacement = RepeatText("123/ITEM" & vbLf, 100)
        Case "No-match scan"
            pattern = "Z\d{4}"
            subject = String$(1024, "a")
            expectedReplacement = subject
        Case "Word boundaries"
            pattern = "\bcat\b"
            subject = "a cat scatter cat!"
            expectedMatches.Add Array("cat")
            expectedMatches.Add Array("cat")
            expectedReplacement = "a x scatter x!"
        Case "Backreference"
            pattern = "^(ab)\1$"
            subject = "abab"
            expectedMatches.Add Array("abab", "ab")
            expectedReplacement = "x"
        Case Else
            Err.Raise 5, "RegexBenchmark", "Unknown scenario: " & scenario
    End Select
End Sub

Private Function NewEngine() As Object
    Select Case engineName
        Case "Sihlfall"
            Dim native As stdRegex
            Set native = New stdRegex
            native.protInit pattern, flags, 2147483647
            Set NewEngine = native
        Case "Old stdRegex"
            Dim legacy As legacyStdRegex
            Set legacy = New legacyStdRegex
            legacy.protInit pattern, flags
            Set NewEngine = legacy
        Case "ASF"
            Dim asf As ASF_RegexEngine
            Set asf = New ASF_RegexEngine
            asf.Init pattern, InStr(flags, "i") > 0, 2147483647, InStr(flags, "m") > 0
            Set NewEngine = asf
    End Select
End Function

Private Function GetMatches() As Collection
    If engineName = "ASF" Then
        Set GetMatches = engine.ExecAll(subject)
        If GetMatches Is Nothing Then Set GetMatches = New Collection
    Else
        Set GetMatches = engine.MatchAll(subject)
    End If
End Function

'All observable numbered captures and the complete replacement string must agree
'with the independently constructed expected output before any timing is accepted.
Private Function Validate(ByVal operation As String) As String
    Dim actual As Boolean, matches As Collection, m As Object, expected As Variant
    Dim i As Long, j As Long, value As String
    If operation = "Compile + first Test" Or operation = "Test" Then
        actual = engine.Test(subject)
        If actual <> (expectedMatches.Count > 0) Then Validate = "Test returned " & actual & "; expected " & (expectedMatches.Count > 0)
    ElseIf operation = "MatchAll" Then
        Set matches = GetMatches()
        If matches.Count <> expectedMatches.Count Then
            Validate = "Match count " & matches.Count & "; expected " & expectedMatches.Count
            Exit Function
        End If
        For i = 1 To matches.Count
            Set m = matches(i)
            expected = expectedMatches(i)
            If engineName = "ASF" Then
                If m.Count <> UBound(expected) + 1 Then Validate = "Capture count differs": Exit Function
            Else
                If m("$COUNT") <> UBound(expected) Then Validate = "Capture count differs": Exit Function
            End If
            For j = 0 To UBound(expected)
                If engineName = "ASF" Then value = m(j + 1) Else value = m(j)
                If value <> expected(j) Then Validate = "Capture value differs at match " & i & " group " & j: Exit Function
            Next
        Next
    ElseIf operation = "Replace" Then
        value = engine.Replace(subject, replacement)
        If value <> expectedReplacement Then Validate = "Replacement differs (actual length " & Len(value) & "; expected " & Len(expectedReplacement) & ")"
    End If
End Function

Private Function Batch(ByVal operation As String, ByVal iterations As Long) As Double
    Dim started As Double, i As Long, matches As Collection, text As String
    started = Seconds()
    Select Case operation
        Case "Compile + first Test"
            For i = 1 To iterations
                Set engine = NewEngine()
                checksum = checksum Xor CLng(engine.Test(subject))
            Next
        Case "Test"
            For i = 1 To iterations
                checksum = checksum Xor CLng(engine.Test(subject))
            Next
        Case "MatchAll"
            For i = 1 To iterations
                Set matches = GetMatches()
                checksum = checksum Xor matches.Count
            Next
        Case "Replace"
            For i = 1 To iterations
                text = engine.Replace(subject, replacement)
                checksum = checksum Xor Len(text)
            Next
    End Select
    Batch = Seconds() - started
End Function

Private Function Quote(ByVal value As String) As String
    Quote = """" & VBA.Replace(value, """", """""") & """"
End Function

'Returns one CSV row. Raw samples are microseconds per call, including wrapper work.
'Five batches follow calibration to >= 30 ms, with a maximum of 16384 calls/batch.
Public Function Measure(ByVal whichEngine As String, ByVal scenario As String, ByVal operation As String) As String
    On Error GoTo Failed
    Dim issue As String, iterations As Long, elapsed As Double, sample As Long, row As String, baseRow As String
    engineName = whichEngine
    Configure scenario
    row = Quote(engineName) & "," & Quote(scenario) & "," & Quote(operation) & "," & Len(subject) & "," & expectedMatches.Count
    baseRow = row
    Set engine = NewEngine()
    issue = Validate(operation)
    If Len(issue) > 0 Then
        Measure = row & ",mismatch," & Quote(issue) & ",0,,,,,"
        Exit Function
    End If
    iterations = 1
    Do
        elapsed = Batch(operation, iterations)
        If elapsed >= 0.03 Or iterations >= 16384 Then Exit Do
        iterations = iterations * 2
    Loop
    row = row & ",ok,," & iterations
    For sample = 1 To 5
        elapsed = Batch(operation, iterations)
        row = row & "," & Trim$(Str$(elapsed * 1000000# / iterations))
    Next
    Measure = row
    Exit Function
Failed:
    Measure = baseRow & ",error," & Quote(CStr(Err.Number) & " " & Err.Description) & ",0,,,,,"
End Function

'Diagnostic only: build 100 capture result containers without doing any regex work.
Public Function MeasureAllocations(ByVal kind As String) As String
    Dim sample As Long, i As Long, started As Double, elapsed As Double, obj As Object
    Dim row As String
    row = Quote(kind) & "," & Quote("100 result containers") & "," & Quote("Allocation baseline") & ",0,100,ok,,1"
    For sample = 1 To 5
        started = Seconds()
        For i = 1 To 100
            If kind = "Dictionary" Then
                Set obj = CreateObject("Scripting.Dictionary")
                obj(0) = "ABC-1234": obj(1) = "ABC": obj(2) = "1234": obj("$COUNT") = 2
            Else
                Set obj = New Collection
                obj.Add "ABC-1234": obj.Add "ABC": obj.Add "1234"
            End If
        Next
        elapsed = Seconds() - started
        row = row & "," & Trim$(Str$(elapsed * 1000000#))
    Next
    MeasureAllocations = row
End Function
