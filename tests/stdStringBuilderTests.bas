Attribute VB_Name = "stdStringBuilderTests"
Sub testAll()
    Test.Topic "stdStringBuilder"

    test_stdStringBuilder_Append_001
    test_stdStringBuilder_Append_002
    test_stdStringBuilder_Append_003
    test_stdStringBuilder_Append_004
    test_stdStringBuilder_Append_010
    test_stdStringBuilder_Str_001
    test_stdStringBuilder_Str_002
    test_stdStringBuilder_Str_003
    test_stdStringBuilder_Str_050
    test_stdStringBuilder_Str_051
    test_stdStringBuilder_Str_052
    test_stdStringBuilder_Str_053
    test_stdStringBuilder_Str_054
    test_stdStringBuilder_Str_055
    test_stdStringBuilder_Str_056
    test_stdStringBuilder_Str_057
    test_stdStringBuilder_Str_060
    test_stdStringBuilder_Demo
End Sub

Private Sub test_stdStringBuilder_Append_001()
    Test.Topic "stdStringBuilder: Test appending several times, staying within minimum buffer size"
    Dim sb As stdStringBuilder
    Set sb = stdStringBuilder.Create()
    sb.JoinStr = vbNullString
    Test.Assert "Initially", "" = sb.Str
    sb.Append ("xyz")
    Test.Assert "After first append",  "xyz" = sb.Str
    sb.Append ("abc")
    Test.Assert "After second append", "xyzabc" = sb.Str
End Sub

Private Sub test_stdStringBuilder_Append_002()
    Test.Topic "stdStringBuilder: Appending several times, exceeding minimum buffer size"
    Dim sb As stdStringBuilder
    Set sb = stdStringBuilder.Create()
    sb.JoinStr = vbNullString
    sb.Append (String(100, "a"))
    Test.Assert "After first append", String(100, "a") = sb.Str
    sb.Append (String(200, "x"))
    Test.Assert "After second append", String(100, "a") & String(200, "x") = sb.Str
End Sub

Private Sub test_stdStringBuilder_Append_003()
    Test.Topic "stdStringBuilder: Appending several times, exceeding minimum buffer size"
    Dim sb As stdStringBuilder
    Set sb = stdStringBuilder.Create()
    sb.JoinStr = vbNullString
    Test.Assert "Initially", "" = sb.Str
    sb.Append (String(100, "a"))
    Test.Assert "After first append", String(100, "a") = sb.Str
    sb.Append (String(200, "x"))
    Test.Assert "After second append", String(100, "a") & String(200, "x") = sb.Str
    sb.Append (String(1000, "y"))
    Test.Assert "After third append", String(100, "a") & String(200, "x") & String(1000, "y") = sb.Str
End Sub

Private Sub test_stdStringBuilder_Append_004()
    Test.Topic "stdStringBuilder: Appending vbNullString to empty buffer"
    Dim sb As stdStringBuilder
    Set sb = stdStringBuilder.Create()
    sb.JoinStr = vbNullString
    Test.Assert "Initially", "" = sb.Str
    sb.Append (vbNullString)
    Test.Assert "After append", "" = sb.Str
End Sub

Private Sub test_stdStringBuilder_Append_010()
    Test.Topic "stdStringBuilder: Calling with bracket notation works"
    Dim sb As Object
    Set sb = stdStringBuilder.Create()
    sb.JoinStr = vbNullString
    Test.Assert "Initially", "" = sb.Str
    sb.[abcdef]
    sb.[ghijkl]
    Test.Assert "After appending", "abcdefghijkl" = sb.Str
End Sub

Private Sub test_stdStringBuilder_Str_001()
    Test.Topic "stdStringBuilder: Initially an empty string is returned"
    Dim sb As stdStringBuilder
    Set sb = stdStringBuilder.Create()
    sb.JoinStr = vbNullString
    Test.Assert "Initially", "" = sb.Str
End Sub

Private Sub test_stdStringBuilder_Str_002()
    Test.Topic "stdStringBuilder: The correct string is returned from buffer 0"
    Dim sb As stdStringBuilder
    Set sb = stdStringBuilder.Create()
    sb.JoinStr = vbNullString
    sb.Append (String(15, "a"))
    Test.Assert "After appending", "aaaaaaaaaaaaaaa" = sb.Str
End Sub

Private Sub test_stdStringBuilder_Str_003()
    Test.Topic "stdStringBuilder: The correct string is returned from buffer 1"
    Dim sb As stdStringBuilder
    Set sb = stdStringBuilder.Create()
    sb.JoinStr = vbNullString
    sb.Append (String(15, "a"))
    sb.Append (String(8, "b"))
    Test.Assert "After appending", "aaaaaaaaaaaaaaabbbbbbbb" = sb.Str
End Sub

Private Sub test_stdStringBuilder_Str_050()
    Test.Topic "stdStringBuilder: Str is default get property"
    Dim sb As stdStringBuilder, s As String
    Set sb = stdStringBuilder.Create()
    sb.JoinStr = vbNullString
    sb.Append ("xyzabc")
    s = sb
    Test.Assert "After appending", "xyzabc" = s
End Sub

Private Sub test_stdStringBuilder_Str_051()
    Test.Topic "stdStringBuilder: Str is default let property"
    Dim sb As stdStringBuilder
    Set sb = stdStringBuilder.Create()
    sb.JoinStr = vbNullString
    sb.Append ("xyzabc")
    sb = "hello world"
    Test.Assert "After appending", "hello world" = sb.Str
End Sub

Private Sub test_stdStringBuilder_Str_052()
    Test.Topic "stdStringBuilder: Assigning a short string to Str works if current content is a rather long string"
    Dim sb As stdStringBuilder
    Set sb = stdStringBuilder.Create()
    sb.JoinStr = vbNullString
    sb.Append (String(10000, "a"))
    sb = String(50, "b")
    Test.Assert "After appending", String(50, "b") = sb.Str
End Sub

Private Sub test_stdStringBuilder_Str_053()
    Test.Topic "stdStringBuilder: After a short string has been assigned to a StringBuilder containing a rather long string, appending works"
    Dim sb As stdStringBuilder
    Set sb = stdStringBuilder.Create()
    sb.JoinStr = vbNullString
    sb.Append (String(10000, "a"))
    sb = String(50, "b")
    sb.Append (String(2000, "c"))
    Test.Assert "After appending", String(50, "b") & String(2000, "c") = sb.Str
End Sub

Private Sub test_stdStringBuilder_Str_054()
    Test.Topic "stdStringBuilder: Assigning the empty string works"
    Dim sb As stdStringBuilder
    Set sb = stdStringBuilder.Create()
    sb.JoinStr = vbNullString
    sb.Append (String(10000, "a"))
    sb = ""
    Test.Assert "After assigning", "" = sb.Str
    sb.Append (String(2000, "c"))
    Test.Assert "After appending", String(2000, "c") = sb.Str
End Sub

Private Sub test_stdStringBuilder_Str_055()
    Test.Topic "stdStringBuilder: Assigning the null string works"
    Dim sb As stdStringBuilder
    Set sb = stdStringBuilder.Create()
    sb.JoinStr = vbNullString
    sb.Append (String(10000, "a"))
    sb = vbNullString
    Test.Assert "After assigning", "" = sb.Str
    sb.Append (String(2000, "c"))
    Test.Assert "After appending", String(2000, "c") = sb.Str
End Sub

Private Sub test_stdStringBuilder_Str_056()
    Test.Topic "stdStringBuilder: Correctly appends single characters after a short string has been assigned to a StringBuilder containing a long string"
    Dim sb As stdStringBuilder, i As Integer
    Set sb = stdStringBuilder.Create()
    sb.JoinStr = vbNullString
    sb.Append (String(10000, "a"))
    sb = "abc"
    Test.Assert "After assigning", "abc" = sb.Str
    For i = 1 To 2500
        sb.Append ("d")
    Next
    Test.Assert "After appending", "abc" & String(2500, "d") = sb.Str
End Sub

Private Sub test_stdStringBuilder_Str_057()
    Test.Topic "stdStringBuilder: Correctly appends single characters after a long string has been assigned to a StringBuilder containing a short string"
    Dim sb As stdStringBuilder, i As Integer
    Set sb = stdStringBuilder.Create()
    sb.JoinStr = vbNullString
    sb.Append ("aa")
    Test.Assert "Before assigning", "aa" = sb.Str
    sb.Str = String(10000, "b")
    Test.Assert "After assigning", String(10000, "b") = sb.Str
    For i = 1 To 25000
        sb.Append ("d")
    Next
    Test.Assert "After appending", String(10000, "b") & String(25000, "d") = sb.Str
End Sub

Private Sub test_stdStringBuilder_Str_060()
    Test.Topic "stdStringBuilder: Correctly appends single characters after setting MinimumCapacity = 0"
    Dim sb As stdStringBuilder, i As Integer
    Set sb = stdStringBuilder.Create()
    sb.JoinStr = vbNullString
    sb.MinimumCapacity = 0
    Test.Assert "MinimumCapacity delivers 2 after being set to 0", 2& = sb.MinimumCapacity
    For i = 1 To 1000
        sb.Append ("d")
    Next
    Test.Assert "After appending", String(1000, "d") = sb.Str
End Sub

Private Sub test_stdStringBuilder_Demo()
    Test.Topic "stdStringBuilder: Demo code"
    Dim sb As Object
    Set sb = stdStringBuilder.Create()
    sb.JoinStr = "-"
    sb.Str = "Start"
    sb.TrimBehaviour = RTrim
    sb.InjectionVariables.Add "@1", "cool"
    sb.[This is a really cool multi-line    ]
    sb.[string which can even include       ]
    sb.[symbols like " ' # ! / \ without    ]       
    sb.[causing compiler errors!!           ]
    sb.[also this has @1 variable injection!]
    Test.Assert "Correct result", "Start-This is a really cool multi-line-string which can even include-symbols like "" ' # ! / \ without-causing compiler errors!!-also this has cool variable injection!" = sb.Str
End Sub
