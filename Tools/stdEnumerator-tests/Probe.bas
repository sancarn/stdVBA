Attribute VB_Name = "EnumeratorProbe"
Option Explicit

Public Function RunTests(ByVal source As Object) As String
    On Error GoTo Failed
    Dim lazy As stdEnumerator
7     source.Configure 100
8     Set lazy = stdEnumerator.CreateFromIEnumVariant(source)
9     Check source.Reads() = 0, "Zero construction reads"
10     Check lazy.First(0).Length = 0, "Zero prefix"
11     Check source.Reads() = 0, "Zero prefix reads"
12     Check lazy.First(100).Length = 100, "Counted prefix"
13     Check source.Reads() = 100, "Exactly 100 Next calls"
14     Check lazy.First(100).Length = 100, "Counted replay"
15     Check source.Reads() = 100, "No replay reads"
    On Error Resume Next
    Dim ignored As Variant
18     ignored = lazy.Item(101)
    Dim errorCode As Long
20     errorCode = Err.Number
    On Error GoTo Failed
22     Check errorCode <> 0, "Source failure propagates"
23     Check lazy.Item(1) = 1, "Cache survives failure"
    On Error Resume Next
25     ignored = lazy.Item(101)
26     errorCode = Err.Number
    On Error GoTo Failed
28     Check errorCode <> 0, "Source failure repeats"
29     Check source.Reads() = 101, "Failed cursor not called again"
30     source.Configure 0
31     Set lazy = stdEnumerator.CreateFromIEnumVariant(source)
32     Check lazy.Length = 1000, "Counted exhaustion"
33     Check source.Reads() = 1001, "One terminal Next"
34     Check lazy.Length = 1000, "Repeated exhaustion"
35     Check source.Reads() = 1001, "No reads after exhaustion"
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "a", 1
    dict.Add "b", 2
    Set lazy = stdEnumerator.CreateFromIEnumVariant(dict)
    Check lazy.Join(",") = "a,b", "Dictionary keys"
    Dim entry As Variant, total As Long
    For Each entry In lazy
        total = total + 1
    Next
    Check total = 2, "Native For Each"
    source.Configure 0
    Set lazy = stdEnumerator.CreateFromIEnumVariant(source)
    Check lazy.enumNext(ignored), "Counted procedural start"
    Check lazy.First(3).Join(",") = "1,2,3", "Interleaved prefix"
    Check lazy.enumNext(ignored), "Counted procedural continuation"
    Check ignored = 2, "Counted independent position"
    lazy.enumRefresh
    Check lazy.enumNext(ignored), "Refresh"
    Check ignored = 1, "Refresh replays cached value"
    Check source.Reads() = 3, "Refresh and interleaving do not reread"
    Dim c As New Collection, e As stdEnumerator, i As Long, v As Variant
37     For i = 1 To 1000
38         c.Add i
39     Next
40     Set e = stdEnumerator.CreateFromIEnumVariant(c)
41     Check e.First(0).Length = 0, "First zero"
42     Check e.First(-1).Length = 0, "First negative"
43     Check e.First(100).Length = 100, "Prefix length"
44     Check e.Item(10) = 10, "Cached item"
45     Check e.enumNext(v), "Procedural start"
46     Check v = 1, "Independent cursor"
47     Check e.First(3).Join(",") = "1,2,3", "Replay"
48     Check e.enumNext(v), "Procedural continuation"
49     Check v = 2, "Independent continuation"
50     Check e.Length = 1000, "Full length"
51     Check e.Join(",") = stdEnumerator.CreateFromArray(e.AsArray()).Join(","), "Full replay"
52     Set c = New Collection
53     Set e = stdEnumerator.CreateFromIEnumVariant(c)
54     Check e.Length = 0, "Empty source"
    Dim obj As New Collection
56     obj.Add 42
57     c.Add obj
58     c.Add Null
59     c.Add Empty
60     c.Add Array(1, 2)
61     Set e = stdEnumerator.CreateFromIEnumVariant(c)
    Dim actual As Object
63     Set actual = e.Item(1)
64     Check actual Is obj, "Object identity"
65     Check IsNull(e.Item(2)), "Null value"
66     Check IsEmpty(e.Item(3)), "Empty value"
67     v = e.Item(4)
68     Check v(1) = 2, "Array value"
69     Check e.Length = 4, "Mixed length"
70     Set e = stdEnumerator.CreateFromIEnumVariant(ThisWorkbook.Worksheets)
71     Set actual = e.Item(1)
72     Check actual Is ThisWorkbook.Worksheets(1), "Excel Worksheets"
73     Set e = stdEnumerator.CreateFromIEnumVariant(ThisWorkbook.Worksheets(1).Cells)
74     Check e.First(100).Length = 100, "Huge Excel Cells prefix"
75     RunTests = "PASS: cached lazy enumeration, boundaries, replay, variants, Worksheets and huge Cells prefix"
    Exit Function
Failed:
78     RunTests = "FAIL line " & Erl & ": " & Err.Number & " " & Err.Description
End Function

Private Sub Check(ByVal condition As Boolean, ByVal label As String)
    If Not condition Then Err.Raise 5, "EnumeratorProbe", label
End Sub

Public Function RunSuite() As String
    On Error GoTo Failed
    stdEnumeratorTests.testAll
    RunSuite = "PASS: " & Test.Count & " existing assertions"
    Exit Function
Failed:
    RunSuite = "FAIL: " & Err.Number & " " & Err.Source & " " & Err.Description
End Function
