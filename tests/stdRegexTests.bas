Attribute VB_Name = "stdRegexTests"
'@lang VBA

Sub testAll()
    test.Topic "stdRegex"
    
    Dim sHaystack as string
    sHaystack = "D-22-4BU - London: London is the capital and largest city of England and the United Kingdom." & vbCrLf & _
                "D-48-8AO - Birmingham: Birmingham is a city and metropolitan borough in the West Midlands, England"  & vbCrLf & _
                "A-22-9AO - Paris: Paris is the capital and most populous city of France. Also contains A-22-9AP." 

    Dim sPattern as string: sPattern = "^(?<Code>(?<Country>[A-Z])-(?:\d{2}-(\d[A-Z]{2})))"
    Dim rx as stdRegex: set rx = stdRegex.Create(sPattern,"m")
    

    Test.Assert "Property Access - Pattern", rx.pattern = sPattern
    Test.Assert "Property Access - Flags", rx.Flags = "m"
    Test.Assert "Test() Method", rx.Test(sHaystack)
    Test.Assert "Test() Method", not stdRegex.Create("^[A-Z]{3}").Test(sHaystack) 'Ensure this returns false
    
    'Match should return a dictionary containing the 1st match only
    Dim oMatch as object: set oMatch = rx.Match(sHaystack)
    Test.Assert "Match returns Dictionary", typename(oMatch) = "Dictionary"
    Test.Assert "Match Dictionary contains named captures 1", oMatch.exists("Code")
    Test.Assert "Match Dictionary contains named captures 2", oMatch.exists("Country")
    Test.Assert "Match Dictionary contains named captures 3", oMatch("Code") = "D-22-4BU"
    Test.Assert "Match Dictionary contains named captures 4", oMatch("Country") = "D"
    Test.Assert "Match Dictionary contains numbered captures 1", oMatch(0) = "D-22-4BU"
    Test.Assert "Match Dictionary contains numbered captures 2", oMatch(1) = "D-22-4BU"
    Test.Assert "Match Dictionary contains numbered captures 3", oMatch(2) = "D"
    Test.Assert "Match Dictionary contains numbered captures 4 & ensure non-capturing group not captured", oMatch(3) = "4BU"
    Test.Assert "Match contains count of submatches", oMatch("$COUNT") = 3
    Test.Assert "Native engine has no COM match collection", Not oMatch.exists("$RAW")

    'MatchAll should return a Collection of Dictionaries, and contain all matches in the haystack
    Dim oMatches as Object: set oMatches = rx.MatchAll(sHaystack)
    Test.Assert "MatchAll returns Collection", typeName(oMatches) = "Collection"
    Test.Assert "MatchAll contains all matches", oMatches.Count = 3
    Test.Assert "MatchAll contains Dictionaries", typeName(oMatches(1)) = "Dictionary"
    Test.Assert "Matchall Dictionaries are populated 1", oMatches(1)("Code") = "D-22-4BU"
    Test.Assert "Matchall Dictionaries are populated 2", oMatches(2)("Code") = "D-48-8AO"
    Test.Assert "Matchall Dictionaries are populated 2", oMatches(3)("Code") = "A-22-9AO"

    'Test Flags letter
    rx.Flags = ""
    Test.Assert "Removing flags sets MatchAll count to 1", rx.MatchAll(sHaystack).Count = 1
    rx.Flags = "m"

    'Test replace
    sHaystack = "Here is some cool data:" & vbCrLf & _
                "12345-STA1  123    10/02/2019" & vbCrLf & _
                "12323-STB9  2123   01/01/2005" & vbCrLf & _
                "and here is some more:" & vbCrLf & _
                "23565-STC2  23     ??/??/????" & vbCrLf & _
                "62346-STZ9  5      01/05/1932"

    Dim sResult as string
    sResult = "Here is some cool data:" & vbCrLf & _
              "12345-STA1,10/02/2019,123" & vbCrLf & _
              "12323-STB9,01/01/2005,2123" & vbCrLf & _
              "and here is some more:" & vbCrLf & _
              "23565-STC2,??/??/????,23" & vbCrLf & _
              "62346-STZ9,01/05/1932,5" 

    set rx = stdRegex.Create("(?<id>\d{5}-ST[A-Z]\d)\s+(?<count>\d+)\s+(?<date>..\/..\/....)","g")
    Test.Assert "Replace", rx.Replace(sHaystack, "$id,$date,$count") = sResult

    'Test List
    sResult = "12345-STA1,10/02/2019,123" & vbCrLf & _
              "12323-STB9,01/01/2005,2123" & vbCrLf & _
              "23565-STC2,??/??/????,23" & vbCrLf & _
              "62346-STZ9,01/05/1932,5"  & vbCrLf
    Test.Assert "List", rx.List(sHaystack, "$id,$date,$count\r\n") = sResult

    'Test ListArr
    Dim vResult as Variant
    vResult = rx.ListArr(sHaystack, Array("$id-$date","$count"))
    Test.Assert "Number of columns match number in array", (Ubound(vResult,2)-Lbound(vResult,2)+1) = 2
    Test.Assert "Number of rows in array = number of rows should be 4", (Ubound(vResult,1)-Lbound(vResult,1)+1) = 4
    Test.Assert "Check data 1,1", vResult(1,1) = "12345-STA1-10/02/2019"
    Test.Assert "Check data 1,2", vResult(1,2) = "123"
    Test.Assert "Check data 2,1", vResult(2,1) = "12323-STB9-01/01/2005"
    Test.Assert "Check data 2,2", vResult(2,2) = "2123"
    testNativeEngine
End Sub

Private Sub testNativeEngine()
    Dim rx As stdRegex, match As Object, matches As Collection, values As Variant

    'Issue 141: backreferences within lookahead must never enter the Office regex engine.
    Set rx = stdRegex.Create("(.)(?=.*\1)", "gim")
    Test.Assert "Issue 141 positive", rx.Test("ss")
    Test.Assert "Issue 141 negative", Not rx.Test("st")
    Test.Assert "Issue 141 case insensitive", rx.Test("sS")
    Set match = rx.Match("ss")
    Test.Assert "Issue 141 captures", match(0) = "s" And match(1) = "s"
    Test.Assert "Issue 141 replace", rx.Replace("ss", "x") = "xs"
    Test.Assert "Backreference", stdRegex.Create("^(ab)\1$").Test("abab")
    Test.Assert "Word boundary", stdRegex.Create("\bcat\b").Test("a cat!")
    Test.Assert "Word boundary negative", Not stdRegex.Create("\bcat\b").Test("scatter")

    Set rx = stdRegex.Create("(?<first>a)(?<second>b)?")
    Set match = rx.Match("a")
    Test.Assert "Unmatched named group", match("second") = ""
    Test.Assert "Unmatched numbered group", match(2) = ""
    Test.Assert "No match returns Nothing", rx.Match("z") Is Nothing
    Test.Assert "No matches returns empty collection", rx.MatchAll("z").Count = 0
    values = rx.ListArr("z", Array("$first"))
    Test.Assert "No matches returns empty array", UBound(values) = -1
    Test.Assert "List retains all matches without g", rx.List("aa", "$first") = "aa"
    Test.Assert "Replace honors missing g", rx.Replace("aa", "$first!") = "a!a"
    rx.flags = "g"
    Test.Assert "Replace honors g", rx.Replace("aa", "$first!") = "a!a!"
    Test.Assert "Modern named replacement", rx.Replace("ab", "$<second>$<first>") = "ba"
    Test.Assert "Replacement dollar escaping", rx.Replace("ab", "$$first") = "$first"
    Test.Assert "Replacement literal dollar", rx.Replace("ab", "$") = "$"
    Test.Assert "Unknown replacement token", rx.Replace("ab", "$unknown") = "$unknown"
    Test.Assert "MatchAll leaves flags unchanged", rx.MatchAll("aa").Count = 2 And rx.flags = "g"

    'Issue 162: dot-all must work at construction without a VBScript lookbehind helper.
    Set rx = stdRegex.Create("a.b", "s")
    Test.Assert "Issue 162 dot-all constructor", rx.Test("a" & vbLf & "b")
    Test.Assert "Issue 162 dot-all adjacent dots", stdRegex.Create("^a..b$", "s").Test("a" & vbCrLf & "b")
    Test.Assert "Issue 162 escaped dot matches literal", stdRegex.Create("^a\.b$", "s").Test("a.b")
    Test.Assert "Issue 162 class dot matches literal", stdRegex.Create("^a[.]b$", "s").Test("a.b")

    Set rx = stdRegex.Create("a.b")
    Test.Assert "Dot excludes newline", Not rx.Test("a" & vbLf & "b")
    rx.flags = "s"
    Test.Assert "s flag", rx.Test("a" & vbLf & "b")
    Test.Assert "s flag replace", rx.Replace("a" & vbLf & "b", "x") = "x"
    Set rx = stdRegex.Create("a\.b", "s")
    Test.Assert "s preserves escaped dot", Not rx.Test("axb")
    Test.Assert "s preserves class dot", Not stdRegex.Create("a[.]b", "s").Test("axb")
    Test.Assert "x flag", stdRegex.Create("a b", "x").Test("ab")
    Test.Assert "x preserves class space", stdRegex.Create("^a[ ]b$", "x").Test("a b")
    Test.Assert "x preserves escaped space", stdRegex.Create("^a\ b$", "x").Test("a b")
    Test.Assert "c flag", stdRegex.Create("a# comment", "c").Test("a")
    Test.Assert "c preserves class hash", stdRegex.Create("a[#]b", "c").Test("a#b")
    Test.Assert "c preserves escaped hash", stdRegex.Create("a\#b", "c").Test("a#b")
    Set rx = stdRegex.Create("^two$", "m")
    Test.Assert "m flag", rx.Test("one" & vbCrLf & "two" & vbCrLf & "three")
    rx.flags = ""
    Test.Assert "Removing m", Not rx.Test("one" & vbCrLf & "two")

    Set rx = stdRegex.Create("setter_initial", "i")
    rx.pattern = "setter_changed"
    Test.Assert "Pattern setter recompiles", rx.Test("SETTER_CHANGED") And Not rx.Test("setter_initial")
    Test.Assert "Pattern setter preserves flags", rx.flags = "i"
    rx.flags = ""
    Test.Assert "Flags setter recompiles", Not rx.Test("SETTER_CHANGED")
    Test.Assert "Pattern getter", rx.pattern = "setter_changed"
    Test.Assert "Invalid pattern preserves state", invalidPatternPreservesState(rx)

    Set rx = stdRegex.Create("(?=a)", "g")
    Test.Assert "Zero width matches advance", rx.MatchAll("aa").Count = 2
    Test.Assert "Zero width replace advances", rx.Replace("aa", "x") = "xaxa"
    Test.Assert "Empty pattern terminates", stdRegex.Create("", "g").MatchAll("a").Count = 2
    Test.Assert "Budget exhaustion raises", stepLimitRaises()
    Test.Assert "Replace budget exhaustion raises", stepLimitRaises("Replace")
    Test.Assert "MatchAll budget exhaustion raises", stepLimitRaises("MatchAll")
    Test.Assert "Create reuses compiled pattern", stdRegex.Create("cache_141") Is stdRegex.Create("cache_141")
    Test.Assert "Cache keys distinguish delimiters", Not (stdRegex.Create("a&g") Is stdRegex.Create("a", "g"))
    Set rx = New stdRegex
    rx.pattern = "new_instance"
    Test.Assert "New instance supports pattern setter", rx.Test("new_instance")
End Sub

Private Function invalidPatternPreservesState(ByVal rx As stdRegex) As Boolean
    On Error GoTo Expected
    rx.pattern = "("
    Exit Function
Expected:
    invalidPatternPreservesState = rx.pattern = "setter_changed" And rx.Test("setter_changed")
End Function

Private Function stepLimitRaises(Optional ByVal operation As String = "Test") As Boolean
    On Error GoTo Expected
    Dim rx As stdRegex, result As Boolean
    Set rx = stdRegex.Create("a+", "", 1)
    Select Case operation
        Case "Test": result = rx.Test("aaaa")
        Case "Replace": rx.Replace "aaaa", "x"
        Case "MatchAll": rx.MatchAll "aaaa"
    End Select
    Exit Function
Expected:
    stepLimitRaises = Err.Number = vbObjectError + 3051
End Function

Sub performanceTest2()
    
End Sub
