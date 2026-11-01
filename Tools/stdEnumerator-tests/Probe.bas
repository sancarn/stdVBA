Attribute VB_Name = "EnumeratorProbe"
Option Explicit
Private mSource As Object
Private mEnum As stdEnumerator
Private mCalls As Long

Public Function RunTests(ByVal source As Object) As String
    On Error GoTo Failed
    Dim lazy As stdEnumerator, ignored As Variant, errorCode As Long
    source.Configure 100
    Set lazy = stdEnumerator.CreateFromIEnumVariant(source)
    Check source.Reads() = 0 And source.Opens() = 0, "No constructor reads/cursors"
    Check lazy.First(0).Length = 0, "Zero prefix"
    Check source.Reads() = 0 And source.Opens() = 0, "Zero prefix does not open source"
    Check lazy.First(100).Length = 100, "Prefix"
    Check source.Reads() = 100, "Exactly 100 Next calls"
    Check source.Active() = 0, "Early return releases cursor"
    Check lazy.First(100).Length = 100, "Repeated prefix"
    Check source.Reads() = 200 And source.Opens() = 2, "Repeated operations reread source"
    On Error Resume Next
    ignored = lazy.Item(101)
    errorCode = Err.Number
    On Error GoTo Failed
    Check errorCode <> 0, "Source failure propagates"
    Check source.Active() = 0, "Failure releases operation cursor"
    Check lazy.Item(1) = 1, "New traversal after failed operation"
    source.Configure 0
    Set lazy = stdEnumerator.CreateFromIEnumVariant(source)
    Check lazy.Length = 1000, "Full traversal"
    Check source.Reads() = 1001 And source.Active() = 0, "Terminal read releases cursor"
    Check lazy.Length = 1000, "Repeated length"
    Check source.Reads() = 2002, "Length rereads source"
    source.Configure 0
    Set lazy = stdEnumerator.CreateFromIEnumVariant(source)
    Check lazy.enumNext(ignored), "Procedural start"
    Check lazy.First(3).Join(",") = "1,2,3", "Interleaved prefix"
    Check lazy.enumNext(ignored), "Procedural continuation"
    Check ignored = 2 And source.Reads() = 5, "Independent positions"
    lazy.enumRefresh
    Check source.Active() = 0, "Refresh releases procedural cursor"
    Check lazy.enumNext(ignored), "Refreshed traversal"
    Check ignored = 1 And source.Reads() = 6, "Refresh starts fresh source"
    lazy.enumRefresh
    source.Configure 2
    Set lazy = stdEnumerator.CreateFromIEnumVariant(source)
    Check lazy.enumNext(ignored), "First before failure"
    Check lazy.enumNext(ignored), "Second before failure"
    On Error Resume Next
    lazy.enumNext ignored
    errorCode = Err.Number
    On Error GoTo Failed
    Check errorCode <> 0 And source.Active() = 0, "Procedural failure releases cursor"
    On Error Resume Next
    lazy.enumNext ignored
    errorCode = Err.Number
    On Error GoTo Failed
    Check errorCode <> 0 And source.Reads() = 3, "Failed procedural cursor stays failed"
    lazy.enumRefresh
    Check lazy.enumNext(ignored), "Refresh recovers from failure"
    lazy.enumRefresh
    source.Configure 0, 3000, True
    Set lazy = stdEnumerator.CreateFromIEnumVariant(source)
    Dim prefix As stdEnumerator
    Set prefix = lazy.First(100)
    Check source.Live() = 100, "Prefix retains only returned results"
    Check source.Reads() = 100 And source.Active() = 0, "Prefix stops and releases source"
    Set prefix = Nothing
    Check source.Live() = 0, "Releasing prefix releases its values"
    Dim streamed As Variant, seen As Long
    source.Configure 0, 3000, True
    For Each streamed In lazy
        seen = seen + 1
        If seen = 100 Then Exit For
    Next
    Check source.Reads() = 100, "Native For Each stops at prefix"
    Set streamed = Nothing
    Check source.Live() = 0 And source.Active() = 0, "Native early exit releases cursor and items"
    source.Configure 0, 3000, True
    Set mSource = source
    Set mEnum = stdEnumerator.CreateFromIEnumVariant(source)
    mCalls = 0
    Check mEnum.countBy(stdCallback.CreateFromModule("EnumeratorProbe", "StreamingPredicate")) = 1500, "Streaming predicate count"
    Check mCalls = 3000 And source.Reads() = 3001, "Predicate runs during each read"
    Check source.Live() = 0 And source.Peak() <= 3, "Consumed objects are not cached"
    Check source.Active() = 0, "Count releases source cursor"
    source.Configure 0, 20
    mCalls = 0
    Check mEnum.countBy(stdCallback.CreateFromModule("EnumeratorProbe", "NestedPredicate")) = 10, "Reentrant operation cursors"
    Check source.Reads() = 23 And source.Opens() = 2, "Nested prefix has separate cursor"
    source.Configure 0, 3
    Dim outer As Variant, inner As Variant, pairs As Long
    For Each outer In mEnum
        For Each inner In mEnum
            pairs = pairs + 1
        Next
    Next
    Check pairs = 9 And source.Opens() = 4, "Nested native loops have separate cursors"
    Check source.Active() = 0, "Nested native loops release cursors"
    Set mEnum = Nothing
    Set mSource = Nothing
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
    Dim c As New Collection, e As stdEnumerator, i As Long, v As Variant
    For i = 1 To 1000
        c.Add i
    Next
    Set e = stdEnumerator.CreateFromIEnumVariant(c)
    Check e.First(0).Length = 0, "First zero"
    Check e.First(-1).Length = 0, "First negative"
    Check e.First(100).Length = 100, "Prefix length"
    Check e.Item(10) = 10, "Repeated item traversal"
    Check e.enumNext(v), "Procedural start"
    Check v = 1, "Independent cursor"
    Check e.First(3).Join(",") = "1,2,3", "Replay"
    Check e.enumNext(v), "Procedural continuation"
    Check v = 2, "Independent continuation"
    Check e.Length = 1000, "Full length"
    Check e.Join(",") = stdEnumerator.CreateFromArray(e.AsArray()).Join(","), "Full replay"
    Set c = New Collection
    Set e = stdEnumerator.CreateFromIEnumVariant(c)
    Check e.Length = 0, "Empty source"
    c.Add 1
    Check e.Length = 1, "Later traversals observe source changes"
    c.Remove 1
    Dim obj As New Collection
    obj.Add 42
    c.Add obj
    c.Add Null
    c.Add Empty
    c.Add Array(1, 2)
    Set e = stdEnumerator.CreateFromIEnumVariant(c)
    Dim actual As Object
    Set actual = e.Item(1)
    Check actual Is obj, "Object identity"
    Check IsNull(e.Item(2)), "Null value"
    Check IsEmpty(e.Item(3)), "Empty value"
    v = e.Item(4)
    Check v(1) = 2, "Array value"
    Check e.Length = 4, "Mixed length"
    Set e = stdEnumerator.CreateFromIEnumVariant(ThisWorkbook.Worksheets)
    Set actual = e.Item(1)
    Check actual Is ThisWorkbook.Worksheets(1), "Excel Worksheets"
    Set e = stdEnumerator.CreateFromIEnumVariant(ThisWorkbook.Worksheets(1).Cells)
    Check e.First(100).Length = 100, "Huge Excel Cells prefix"
    RunTests = "PASS: uncached streaming, bounded live values, independent cursors, variants and Excel collections"
    Exit Function
Failed:
    RunTests = "FAIL line " & Erl & ": " & Err.Number & " " & Err.Description
End Function

Public Function StreamingPredicate(ByVal item As Object, ByVal index As Long) As Boolean
    mCalls = mCalls + 1
    Check mSource.Reads() = mCalls, "Predicate interleaves with Next"
    Check mSource.Live() <= 3, "Only current values remain live"
    StreamingPredicate = item.Id() Mod 2 = 0
End Function

Public Function NestedPredicate(ByVal item As Variant, ByVal index As Long) As Boolean
    mCalls = mCalls + 1
    If mCalls = 1 Then Check mEnum.First(2).Join(",") = "1,2", "Nested prefix"
    Check item = mCalls, "Nested operation preserves outer position"
    NestedPredicate = item Mod 2 = 0
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
