Attribute VB_Name = "stdRegexTests"
'@lang VBA

Sub testAll()
    test.Topic "stdRegex"
    
    Dim sHaystack as string
    sHaystack = "D-22-4BU - London: London is the capital and largest city of England and the United Kingdom." & vbCrLf & _
                "D-48-8AO - Birmingham: Birmingham is a city and metropolitan borough in the West Midlands, England"  & vbCrLf & _
                "A-22-9AO - Paris: Paris is the capital and most populous city of France. Also contains A-22-9AP." 

    On Error Resume Next
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
    Test.Assert "Match contains regex match object", TypeName(oMatch("$RAW")) = "IMatchCollection2"

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
    Test.Assert "List", rx.List(sHaystack, "$id,$date,$count\r\n")

    'Test ListArr - ListArr not currently implemented correctly 
    Dim vResult as Variant
    vResult = rx.ListArr(sHaystack, Array("$id-$date","$count"))
    Test.Assert "Number of columns match number in array", (Ubound(vResult,2)-Lbound(vResult,2)+1) = 2
    Test.Assert "Number of rows in array = number of rows should be 4", (Ubound(vResult,1)-Lbound(vResult,1)+1) = 4
    Test.Assert "Check data 1,1", vResult(1,1) = "12345-STA1-10/02/2019"
    Test.Assert "Check data 1,2", vResult(1,2) = "123"
    Test.Assert "Check data 2,1", vResult(2,1) = "12323-STB9-01/01/2005"
    Test.Assert "Check data 2,2", vResult(2,2) = "2123"
    
End Sub

Sub performanceTest2()
    
End Sub
