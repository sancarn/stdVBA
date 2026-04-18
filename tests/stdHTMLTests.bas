Attribute VB_Name = "stdHTMLTests"
'@lang VBA

Sub testAll()
    Test.Topic "stdHTML"

    Dim node As stdHTML
    Dim attrValue As Variant

    Set node = stdHTML.CreateFromHTML("<input disabled />")
    attrValue = node.Attr("disabled")
    Test.Assert "Parse minimized attr returns Null", IsNull(attrValue)
    Test.Assert "Serialize minimized attr", node.ToString() = "<input disabled />"

    Set node = stdHTML.CreateFromHTML("<a something=""false"" />")
    attrValue = node.Attr("something")
    Test.Assert "Explicit false value remains string", VarType(attrValue) = vbString
    Test.Assert "Explicit false string value", CStr(attrValue) = "false"

    Set node = stdHTML.CreateFromHTML("<el />")
    node.Attr("enabled") = True
    node.Attr("hidden") = False
    node.Attr("compact") = Null
    node.Attr("temp") = "remove-me"
    node.Attr("temp") = Empty

    Test.Assert "Serialize true as literal true", InStr(1, node.ToString(), " enabled=""true""", vbBinaryCompare) > 0
    Test.Assert "Serialize false as literal false", InStr(1, node.ToString(), " hidden=""false""", vbBinaryCompare) > 0
    Test.Assert "Serialize Null as minimized attribute", InStr(1, node.ToString(), " compact", vbBinaryCompare) > 0
    Test.Assert "Empty removes attribute from serialization", InStr(1, node.ToString(), " temp=", vbBinaryCompare) = 0
    Test.Assert "Empty removes attribute from store", IsEmpty(node.Attr("temp"))

    Set node = stdHTML.CreateFromHTML("<el a=""false"" />")
    Test.Assert "Presence selector supports [attr]", Not (node.QuerySelector("[a]") Is Nothing)
    Test.Assert "Value selector [attr=value] still works", Not (node.QuerySelector("[a=false]") Is Nothing)

    Set node = stdHTML.CreateFromHTML("<el a />")
    Test.Assert "Presence selector matches minimized attr", Not (node.QuerySelector("[a]") Is Nothing)
    Test.Assert "Value selector does not coerce minimized attr", node.QuerySelector("[a=false]") Is Nothing

    node.Attr("a") = Empty
    Test.Assert "Presence selector excludes removed attr", node.QuerySelector("[a]") Is Nothing
End Sub
