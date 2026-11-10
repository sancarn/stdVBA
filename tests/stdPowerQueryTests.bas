Attribute VB_Name = "stdPowerQueryTests"
Option Explicit

Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal codePage As Long, ByVal flags As Long, ByVal text As LongPtr, ByVal chars As Long, ByVal buffer As LongPtr, ByVal bytes As Long, ByVal defaultChar As LongPtr, ByVal usedDefaultChar As LongPtr) As Long

Private Sub Check(ByVal condition As Boolean, ByVal label As String)
  Test.Assert label, condition
  If Not condition Then Err.Raise vbObjectError + 1200, "stdPowerQueryTests", label
End Sub

Public Sub RecordAccess()
  Dim query As stdPowerQuery, saved As Object, unused As Variant, number As Long
  Set query = stdPowerQuery.Create("#table({""Col1""},{{1},{2}})")
  On Error Resume Next
  query.protInit "42"
  number = Err.Number
  On Error GoTo 0
  Check number <> 0, "Repeated initialization rejected"
  Check Not query.EOF, "Factory positions on first row"
  Check query!Col1 = 1 And query.FieldValue("Col1") = 1 And query("Col1") = 1, "Bang and default field access"
  Check query.FieldType("Col1") = "Int64", "Untyped integer inferred as Int64"
  Set saved = query.Current
  Check query.MoveNext, "Advance to second row"
  Check query!Col1 = 2 And saved("Col1") = 1, "Snapshot survives advance"
  Check Not query.MoveNext, "No extra row"
  Check query.EOF And query.Completed And query.IsClosed And query.Count = 2, "Normal exhaustion state"
  Check query.FieldType("Col1") = "Any", "Untyped schema remains Any at EOF"
  On Error Resume Next
  unused = query!Col1
  number = Err.Number
  On Error GoTo 0
  Check number <> 0, "Field access at EOF rejected"

  Set query = stdPowerQuery.Create("[a=1,A=2]")
  Check query.FieldValue("a") = 1 And query.FieldValue("A") = 2, "Exact case-sensitive fields"
  Set saved = query.Current
  On Error Resume Next
  unused = query.FieldValue("missing")
  number = Err.Number
  On Error GoTo 0
  Check number <> 0 And saved.Count = 2, "Missing field lookup does not mutate record"
  On Error Resume Next
  unused = query.FieldType("missing")
  number = Err.Number
  On Error GoTo 0
  Check number <> 0, "Missing field type rejected"
  query.CloseStream
End Sub

Public Sub SchemaAndValues()
  Dim query As stdPowerQuery, value As Variant
  Set query = stdPowerQuery.Create("#table(type table [Id=Int64.Type,Name=nullable text],{})")
  Check query.EOF And query.Completed And query.Count = 0, "Empty table starts at EOF"
  Check query.Columns.Count = 2, "Empty table retains columns"
  Check query.FieldType("Id") = "Int64" And query.FieldType("Name") = "Text", "Empty table retains declared types"
  Set query = stdPowerQuery.Create("#table(type table [N=number,T=nullable text],{{1,null}})")
  Check query.FieldType("N") = "Number" And query.FieldType("T") = "Text", "Declared types take precedence"
  Check IsNull(query!T), "Typed null preserved"

  Set query = stdPowerQuery.Create("[n=1.25,b=true,t=""text"",nil=null,d=#date(2026,11,8),dt=#datetime(2026,11,8,1,2,3),z=#datetimezone(2026,11,8,1,2,3,5,30),time=#time(1,2,3),duration=#duration(1,2,3,4),bin=#binary({0,255}),list={1},r=[x=2],tab=#table({""x""},{{3}})]")
  Check query.FieldType("n") = "Double" And query.FieldType("b") = "Logical" And query.FieldType("t") = "Text", "Primitive value kinds"
  Check query.FieldType("nil") = "Null" And IsNull(query!nil), "Null value kind"
  Check query.FieldType("d") = "Date" And query.FieldType("dt") = "DateTime" And query.FieldType("z") = "DateTimeZone", "Date value kinds"
  Check VarType(query!d) = vbDate And Right$(query!z, 6) = "+05:30", "Date and zone conversion"
  Check query.FieldType("time") = "Time" And query.FieldType("duration") = "Duration", "Time and duration value kinds"
  Check query!time = "01:02:03" And query!duration = "1.02:03:04", "Time and duration conversion"
  Check query.FieldType("bin") = "Binary" And query.FieldType("list") = "List", "Binary and list value kinds"
  Check query.FieldType("r") = "Record" And query.FieldType("tab") = "Table", "Structured value kinds"
  Check query.FieldValue("list")(1) = 1 And query.FieldValue("r")("x") = 2, "Nested object fields"
  Check query.FieldValue("tab")(1)("x") = 3, "Nested table conversion"
  value = query!bin
  Check value(1) = 255, "Binary payload"
  Set query = stdPowerQuery.Create("[big=Int64.From(""9223372036854775807"")]")
  Check CStr(query!big) = "9223372036854775807" And query.FieldType("big") = "Int64", "Large integer precision"

  Set query = stdPowerQuery.Create("null")
  Check query.Columns(1) = "Value" And Not query.EOF And IsNull(query!Value), "Scalar null is one Value record"
  query.MoveNext
  Check query.EOF And query.Count = 1, "Scalar emits once"
  Set query = stdPowerQuery.Create("{null,42,[x=3]}")
  Check query.FieldType("Value") = "Null", "First list item"
  query.MoveNext
  Check query!Value = 42 And query.FieldType("Value") = "Int64", "Type inference changes with list item"
  query.MoveNext
  Check query.FieldValue("Value")("x") = 3 And query.FieldType("Value") = "Record", "List record remains nested"
  query.MoveNext
  Check query.EOF, "List exhaustion"
  value = "quote""#(lf)" & vbLf & ChrW$(&H3BB) & ChrW$(&HD83D) & ChrW$(&HDE03)
  Set query = stdPowerQuery.Create(stdPowerQuery.MText(CStr(value)))
  Check query!Value = value, "M text escaping and Unicode"
  query.CloseStream
End Sub

Public Sub StreamingAndErrors()
  Dim query As stdPowerQuery, other As stdPowerQuery, number As Long, message As String
  Dim checksum As Double, unused As Variant
  Set query = stdPowerQuery.Create("{1,2}")
  Set other = stdPowerQuery.Create("{3,4}")
  query.CloseStream
  query.CloseStream
  Check query.EOF And Not query.Completed And query.Count = 1, "Idempotent early close"
  other.MoveNext
  Check other!Value = 4, "Independent query lifetime"
  other.CloseStream

  Set query = stdPowerQuery.Create("#table({""x""},{{1},{error ""late""}})")
  Check query!x = 1, "First row available before lazy error"
  On Error Resume Next
  query.MoveNext
  number = Err.Number
  message = Err.Description
  On Error GoTo 0
  Check number <> 0 And InStr(message, "late") > 0, "Lazy error preserves message"
  Check query.EOF And query.IsClosed And Not query.Completed And query.Count = 1, "Failure never reports completion"
  On Error Resume Next
  Set other = stdPowerQuery.Create("error ""first""")
  number = Err.Number
  On Error GoTo 0
  Check number <> 0, "Factory propagates first-record error"
  On Error Resume Next
  Set other = stdPowerQuery.Create("let x =")
  number = Err.Number
  On Error GoTo 0
  Check number <> 0, "Syntax error propagates"
  On Error Resume Next
  Set other = stdPowerQuery.Create("(x)=>x")
  number = Err.Number
  On Error GoTo 0
  Check number <> 0, "Unsupported function result rejected"

  Set query = stdPowerQuery.Create("List.Generate(()=>0, each true, each _+1)")
  Check query!Value = 0, "Infinite list first value"
  query.MoveNext
  Check query!Value = 1, "Infinite list advances incrementally"
  query.CloseStream
  Set query = stdPowerQuery.Create("Table.FromColumns({{1..1100}}, {""Id""})")
  While Not query.EOF
    checksum = checksum + query!Id
    query.MoveNext
  Wend
  Check checksum = 605550 And query.Count = 1100 And query.Completed, "1100 records without skips or duplicates"
  Set query = stdPowerQuery.Create("{[x=1],[a=[b=[c=1]]]}")
  query.MaxDepth = 1
  On Error Resume Next
  query.MoveNext
  number = Err.Number
  On Error GoTo 0
  Check number <> 0 And query.EOF And Not query.Completed, "Depth failure closes query"
End Sub

Public Sub FileFactories()
  Dim query As stdPowerQuery, other As stdPowerQuery, source As String, text As String
  Dim valid As String, bom As String, invalid As String, emptyFile As String, missing As String
  Dim file As Integer, badByte As Byte, number As Long, description As String, runtime As String
  On Error GoTo Failed
  text = ChrW$(&H3BB) & ChrW$(&HD83D) & ChrW$(&HDE03)
  source = "#table({""Name""},{{" & stdPowerQuery.MText(text) & "}})"
  valid = FixturePath(): bom = FixturePath(): invalid = FixturePath(): emptyFile = FixturePath(): missing = FixturePath()
  WriteUtf8 valid, source, False
  WriteUtf8 bom, source, True
  WriteUtf8 emptyFile, "", False
  file = FreeFile
  Open invalid For Binary Access Write As #file
  badByte = &HFF
  Put #file, , badByte
  Close #file
  file = 0
  Set query = stdPowerQuery.CreateFromFile(valid)
  Check query!Name = text, "UTF-8 file without BOM"
  runtime = query.RuntimePath
  Set other = stdPowerQuery.CreateFromFile(bom, runtime)
  Check other!Name = text And other.RuntimePath = runtime, "UTF-8 BOM and explicit runtime path"
  On Error Resume Next
  Set other = stdPowerQuery.CreateFromFile(invalid)
  number = Err.Number
  On Error GoTo Failed
  Check number <> 0, "Invalid UTF-8 rejected"
  On Error Resume Next
  Set other = stdPowerQuery.CreateFromFile(emptyFile)
  number = Err.Number
  On Error GoTo Failed
  Check number <> 0, "Empty file rejected"
  On Error Resume Next
  Set other = stdPowerQuery.CreateFromFile(missing)
  number = Err.Number
  On Error GoTo Failed
  Check number <> 0 And Dir$(missing) = "", "Missing file rejected without creation"
  On Error Resume Next
  Set other = stdPowerQuery.Create("File.Contents(" & stdPowerQuery.MText(valid) & ")")
  number = Err.Number
  description = Err.Description
  On Error GoTo Failed
  Check number <> 0 And InStr(description, "UnpermittedResourceAccessException") > 0, "Source-file factory does not grant M connector access"
  Check ThisWorkbook.Queries.Count = 0 And ThisWorkbook.Connections.Count = 0 And ThisWorkbook.Worksheets(1).UsedRange.Cells.Count = 1, "No workbook query, connection, or worksheet output"
  number = 0
CleanUp:
  On Error Resume Next
  If Not query Is Nothing Then query.CloseStream
  If Not other Is Nothing Then other.CloseStream
  If file <> 0 Then Close #file
  If Len(valid) > 0 Then Kill valid
  If Len(bom) > 0 Then Kill bom
  If Len(invalid) > 0 Then Kill invalid
  If Len(emptyFile) > 0 Then Kill emptyFile
  On Error GoTo 0
  If number <> 0 Then Err.Raise number, "stdPowerQueryTests.FileFactories", description
  Exit Sub
Failed:
  number = Err.Number: description = Err.Description
  Resume CleanUp
End Sub

Private Function FixturePath() As String
  Dim fso As Object
  Set fso = CreateObject("Scripting.FileSystemObject")
  FixturePath = fso.BuildPath(fso.GetSpecialFolder(2), fso.GetTempName)
End Function

Private Sub WriteUtf8(ByVal path As String, ByVal text As String, ByVal bom As Boolean)
  Dim bytes() As Byte, header(0 To 2) As Byte, size As Long, file As Integer
  Dim number As Long, description As String
  On Error GoTo Failed
  file = FreeFile
  Open path For Binary Access Write As #file
  If bom Then
    header(0) = &HEF: header(1) = &HBB: header(2) = &HBF
    Put #file, , header
  End If
  If Len(text) > 0 Then
    size = WideCharToMultiByte(65001, &H80, StrPtr(text), Len(text), 0, 0, 0, 0)
    If size = 0 Then Err.Raise 5, , "UTF-8 fixture encoding failed"
    ReDim bytes(0 To size - 1)
    If WideCharToMultiByte(65001, &H80, StrPtr(text), Len(text), VarPtr(bytes(0)), size, 0, 0) <> size Then Err.Raise 5
    Put #file, , bytes
  End If
  Close #file
  Exit Sub
Failed:
  number = Err.Number: description = Err.Description
  On Error Resume Next
  If file <> 0 Then Close #file
  On Error GoTo 0
  Err.Raise number, "stdPowerQueryTests.WriteUtf8", description
End Sub
