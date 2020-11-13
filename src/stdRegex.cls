VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdRegex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Adds new features:
'* Named capturing groups
'* TODO: Modern symbols e.g. \A and \Z symbols
'* TODO: Maybe add ListArr() - Produces a list in a table form (as array), ready to import directly into Excel.
'* ...?+ --> (?:...?)+
'
'The class works as follows:
'  * Input regex is transpiled to ES6 convention using options
'  * Regex is executed using VBScript
'  * Regex object returned with new features.

'Usage examples:
'     stdRegex[/hello (?<name>.*)/i].match("hello Andy Chan")("name")



Private p_initialised As Boolean
Private p_pattern As String
Private p_cPattern As String
Private p_flags As String
Private p_namesDict As Object
Private p_regex As Object


'Get pattern for regex
Public Property Get pattern() As String
  If Me Is stdRegex Then
    stdError.Raise ("Property called on class not object")
    Exit Property
  End If

  pattern = p_pattern
End Property

'Set pattern for regex
Public Property Let pattern(val As String)
  If Me Is stdRegex Then
    stdError.Raise ("Property called on class not object")
    Exit Property
  End If
  
  Call init(val, p_flags)
End Property




'Get Flags for regex
Public Property Get flags() As String
  If Me Is stdRegex Then
    stdError.Raise ("Property called on class not object")
    Exit Property
  End If

  flags = p_flags
End Property

'Set Flags for regex
Public Property Let flags(val As String)
  If Me Is stdRegex Then
    stdError.Raise ("Property called on class not object")
    Exit Property
  End If

  Call init(p_pattern, val)
End Property



'Creates a regex object given a pattern and flags.
'@constructor
'@EVAL
'
'@param {string}  Pattern and flags in ECMA Script syntax.
'@return {stdRegex} Regular expression object
'@example
'    stdRegex.Create2("/A\d+/i")
'    stdRegex[/hello world/i]
Public Function Create2(ByVal rx As String) As stdRegex
Attribute Create2.VB_UserMemId = -5

  If Not Me Is stdRegex Then
    stdError.Raise ("Constructor called on object not class")
    Exit Function
  End If

  'regex given like /asdf/flags
  Dim iEnd As Long: iEnd = inStrRev(rx, "/")
  Dim sPattern As String: sPattern = Mid(rx, 2, iEnd - 2)
  Dim sFlags As String: sFlags = Mid(rx, iEnd + 1)
  Debug.Assert False
  Set Create2 = Create(sPattern, sFlags)
End Function

'Creates a regex object given a pattern and flags.
'@constructor
'
'@param {string}  Pattern - The pattern to match
'@param {string}  Flags - Optional flags to apply
'@return {stdRegex} Regular expression object
'@example
'    stdRegex.Create("A\d+","i")
Public Function Create(ByVal pattern As String, Optional ByVal flags As String = "") As stdRegex
  If Not Me Is stdRegex Then
    stdError.Raise ("Constructor called on object not class")
    Exit Function
  End If
  
  Set Create = New stdRegex
  Call Create.init(pattern, flags)
End Function



'Initialises the class from within the static superclass. This method is meant for internal use only. Use at your own risk.
'@protected
'
'@param {string}  Pattern - The pattern to match
'@param {string}  Flags - Optional flags to apply
'@example
'    obj.init("A\d+","i")
Friend Sub init(ByVal pattern As String, ByVal flags As String)
  If Me Is stdRegex Then
    stdError.Raise ("Cannot run init on class")
    Exit Sub
  End If
  
  p_pattern = pattern
  p_flags = flags

  'Named regexp used to detect capturing groups and named capturing groups
  'Captures all capturing groups. All capturing groups with a name has a subcapturing group containing the name.
  'This is used to distribute names across the submatches captured across the regular expression.
  'Not the regex used here has to ignore all non-capturing groups. E.G. (?:...), (?!...), (?=...), (?<=...), (?<!...)
  'In theory must also not match a (\\)*\(
  Static NamesMatcher As Object
  If NamesMatcher Is Nothing Then
    Set NamesMatcher = CreateObject("VBScript.RegExp")
    NamesMatcher.pattern = "\((?!(?:\?:|\?!|\?=|\?<=|\?<!))(?:\?\<(\w+?)\>)?"  'SaveEscape("(")  -->  match regex --> LoadEscape("(")
    NamesMatcher.Global = True
  End If
  
  'Names replacer is used to strip the name syntax from the regular expression before it's injected into VBScript.Regexp.
  Static NamesReplacer As Object
  If NamesReplacer Is Nothing Then
    Set NamesReplacer = CreateObject("VBScript.Regexp")
    NamesReplacer.pattern = "\((?!(?:\?:|\?!|\?=|\?<=|\?<!))(?:\?\<(\w+?)\>)?"  'SaveEscape("(")  -->  Replace regex with "(" --> LoadEscape("(")
    NamesReplacer.Global = True
  End If

  'Ignore White Space option (x) will ignore all undeclared whitespace. It also makes # a comment
  'i.e. `/hello world\ and stuff #this is a comment/` --> `/helloworld andstuff/`
  Static OptionIgnoreWhiteSpace As Object
  If OptionIgnoreWhiteSpace Is Nothing Then
    Set OptionIgnoreWhiteSpace = CreateObject("VBScript.RegExp")
    OptionIgnoreWhiteSpace.pattern = " +"                               'SaveEscape(" ")  -->  Replace " +" with "" --> LoadEscape(" ")
    OptionIgnoreWhiteSpace.Global = True
  End If
  
  'Non standard option to allow for the injection of comments into the end of regex strings using #.* under option c
  Static OptionIgnoreComment As Object
  If OptionIgnoreComment Is Nothing Then
    Set OptionIgnoreComment = CreateObject("VBScript.RegExp")
    OptionIgnoreComment.pattern = "#.*"                                'SaveEscape("#")  -->  Replace "#.*" with "" --> LoadEscape("#")
    OptionIgnoreComment.Global = True
  End If

  'SingleLineMode (s) forces period (.) to match every character (instead of every character apart from `\n`).
  'With this regex you can do:            OptionSingleLineMode.replace(sRegex,"$1(?:.|\s)")
  'to inject the new regex characters into place.
  Static OptionSingleLineMode As Object
  If OptionSingleLineMode Is Nothing Then
    Set OptionSingleLineMode = CreateObject("VBScript.RegExp")
    OptionSingleLineMode.pattern = "(?<!\\)((?:\\\\)*)\."             'SaveEscape(".")  -->  Replace "." with "(?:.|\s)" --> LoadEscape(".")
    OptionSingleLineMode.Global = True
  End If

  'Create dictionary of names in the regex.
  Set p_namesDict = NamesMatcher.Execute(SaveEscape(p_pattern, "("))
  
  
  'Initialise p_regex
  Set p_regex = CreateObject("VBScript.RegExp")
  p_regex.pattern = LoadEscape(NamesReplacer.Replace(SaveEscape(p_pattern, "("), "("), "(")

  'Process regex options
  Dim i As Integer
  For i = 1 To Len(flags)
    Select Case Mid(flags, i, 1)
      Case "i"
        p_regex.ignoreCase = True
      Case "g"
        p_regex.Global = True
      Case "m"
        p_regex.MultiLine = True
      Case "x"
        p_regex.pattern = LoadEscape(OptionIgnoreWhiteSpace.Replace(SaveEscape(p_regex.pattern, " "), ""), " ")
      Case "c"
        p_regex.pattern = LoadEscape(OptionIgnoreComment.Replace(SaveEscape(p_regex.pattern, "#"), ""), "#")
      Case "s"
        p_regex.pattern = LoadEscape(OptionSingleLineMode.Replace(SaveEscape(p_regex.pattern, "."), "(?:.|\s)"), ".")
    End Select
  Next

  p_initialised = True
End Sub

'These functions allow for easy use of
Private Function SaveEscape(ByVal sHaystack As String, ByVal sChar As String) As String
  SaveEscape = VBA.Replace(VBA.Replace(sHaystack, "\\", ";f934b47b-b799-4f52-9b4c-f170a82b52fe;"), "\" & sChar, ";6632df85-1730-4159-a742-5b0c8c526ab8;")
End Function
Private Function LoadEscape(ByVal sHaystack As String, ByVal sChar As String) As String
  LoadEscape = VBA.Replace(VBA.Replace(sHaystack, ";f934b47b-b799-4f52-9b4c-f170a82b52fe;", "\\"), ";6632df85-1730-4159-a742-5b0c8c526ab8;", "\" & sChar)
End Function





'Given a text string, return whether the source regex is present
'
'@param {string}  sHaystack - Text to search for regex in.
'@return {Boolean} True if the regex is present, false if the regex is not present
'@example
'Regex:
'    \d{5}-ST[A-Z]\d
'
'Data:
'    * The site with id 12345-STA1 is one hell of a cool site.
'Result:
'    True
Public Function Test(ByVal sHaystack As String) As Boolean
  If Me Is stdRegex Then
    stdError.Raise ("Method called on class not object")
    Exit Function
  End If
  
  Test = p_regex.Test(sHaystack)
End Function






'Given a text string, return the first match.
'
'@param {string}  sHaystack - Text to search for regex in.
'@return {Dictionary} A dictionary containing the match, submatches, named matches, Count and Raw match object
'@example
'Regex:
'    (?<id>\d{5}-ST[A-Z]\d) - (?<desc>.*)
'
'Data:
'    Some sites were in critical condition
'    * 12345-STA1 - Large crack through pipe.
'    * 12323-STB9 - Acid leakage polluting watercourse.
'    and some others were largely ok:
'    * 23565-STC2
'    * 62346-STZ9
'Result:
'    {
'      0: "12345-STA1 - Large crack through pipe.",
'      "id":"12345-STA1",
'      1: "12345-STA1",
'      "desc": "Large crack through pipe."
'      2: "Large crack through pipe.",
'      "$COUNT":2,
'      "$RAW": {...}
'    }
Public Function Match(ByVal sHaystack As String) As Object
  If Me Is stdRegex Then
    stdError.Raise ("Method called on class not object")
    Exit Function
  End If
  
  'Execute regex on haystack provided
  Dim oMatches As Object
  Set oMatches = p_regex.Execute(sHaystack)
  
  'Create dictionary to store this match's data
  Dim oRet As Object
  Set oRet = CreateObject("Scripting.Dictionary")

  'Basic properties
  oRet(0) = oMatches(0)
  oRet("$COUNT") = oMatches(0).Submatches.Count
  Set oRet("$RAW") = oMatches

  'Loop over submatches and apply to dict
  Dim i As Long
  For i = 1 To p_namesDict.Count
    oRet(i) = oMatches(0).Submatches(i - 1)
    If Not IsEmpty(oMatches(0).Submatches(i - 1)) then
      If Not IsEmpty(p_namesDict(i - 1).Submatches(0)) Then oRet(p_namesDict(i - 1).Submatches(0)) = oMatches(0).Submatches(i - 1)
    End If
  Next i
  
  'Return data
  Set Match = oRet
End Function

'Given a text string, return all strings which match the source regex pattern.
'
'@param {string}  sHaystack - Text to search for regex in.
'@return {Collection<Dictionary>} An array of strings which match the regex.
'@example
'Regex:
'    \d{5}-ST[A-Z]\d
'
'Data:
'    Some sites were in critical condition
'    * 12345-STA1 - Large crack through pipe.
'    * 12323-STB9 - Acid leakage polluting watercourse.
'    and some others were largely ok:
'    * 23565-STC2
'    * 62346-STZ9
'Result:
'    [{0: "12345-STA1", ...}, {0: "12323-STB9", ...}, {0: "23565-STC2", ...}, {0: "62346-STZ9", ...}]
Public Function MatchAll(ByVal sHaystack As String) As Collection
  If Me Is stdRegex Then
    stdError.Raise ("Method called on class not object")
    Exit Function
  End If
  
  'Note: In order to match all, we need to ensure global is true, but will revert our changes afterwards.
  Dim bOriginalGlobal As Boolean: bOriginalGlobal = p_regex.Global
  p_regex.Global = True
  
  'Execute regex on haystack provided
  Dim oMatches As Object
  Set oMatches = p_regex.Execute(sHaystack)
  
  'Revert forced global
  p_regex.Global = bOriginalGlobal
  
  'Create array to hold match data
  Dim oRet As Collection
  Set oRet = new Collection
  
  'Loop over all matches
  Dim i As Long, j As Long
  For i = 1 To oMatches.Count
    'Create dictionary to store this match's data
    Dim oDict As Object
    Set oDict = CreateObject("Scripting.Dictionary")

    'Basic properties
    oDict(0) = oMatches(i - 1)
    oDict("$COUNT") = oMatches(i - 1).Submatches.Count
    Set oDict("$RAW") = oMatches

    'Loop over submatches and apply to dict
    For j = 1 To p_namesDict.Count
      oDict(j) = oMatches(i - 1).Submatches(j - 1)
      If Not IsEmpty(oMatches(i - 1).Submatches(j - 1)) then
        If Not IsEmpty(p_namesDict(j - 1).Submatches(0)) Then oDict(p_namesDict(j - 1).Submatches(0)) = oMatches(i - 1).Submatches(j - 1)
      end if
    Next

    'Push dictionary to array
    Call oRet.add(oDict)
  Next

  'Return data
  Set MatchAll = oRet
End Function


'Given a Haystack and a Replacer, Replace all matches with the source regex with the format given in the Replacer.
'
'@param {string}  sHaystack - Text to search for regex in.
'@param {string}  sReplacer - Format of replacement text.
'@return {string} The haystack with replaced details.
'@example
'Regex:
'    (?<id>\d{5}-ST[A-Z]\d))\s+(?<count>\d+)\s+(?<date>../../....)
'
'Data:
'    Here is some cool data:
'    12345-STA1  123    10/02/2019
'    12323-STB9  2123   01/01/2005
'    and here is some more:
'    23565-STC2  23     ??/??/????
'    62346-STZ9  5      01/05/1932
'Replacer: (to csv with column re-order)
'    $id,$date,$count
'Result:
'    Here is some cool data:
'    12345-STA1,10/02/2019,123
'    12323-STB9,01/01/2005,2123
'    and here is some more:
'    23565-STC2,??/??/????,23
'    62346-STZ9,01/05/1932,5
Public Function Replace(ByVal sHaystack As String, ByVal sReplacer As String) As String
  If Me Is stdRegex Then
    stdError.Raise ("Method called on class not object")
    Exit Function
  End If
  
  'Replace names in sReplacer with indexed replacers ($1,$2,...)
  For i = 1 To p_namesDict.Count
    If Not IsEmpty(p_namesDict(i - 1).Submatches(0)) Then
      Dim sName As String
      sName = p_namesDict(i - 1).Submatches(0)
      
      sReplacer = VBA.Replace(sReplacer, "$" & sName, "$" & i)
    End If
  Next

  'Pass to vbscript regex handler and return
  Replace = p_regex.Replace(sHaystack, sReplacer)
End Function


'Given a Haystack and a Format, Generate a string containing all matches in the format provided.
'
'@param {string}  sHaystack - Text to search for regex in.
'@param {string}  sFormat   - Format of output list.
'@return {string} A list of strings matched.
'@example
'Regex:
'    (?<id>\d{5}-ST[A-Z]\d))\s+(?<count>\d+)\s+(?<date>../../....)
'
'Data:
'    12345-STA1  123    10/02/2019
'    12323-STB9  2123   01/01/2005
'    23565-STC2  23     ??/??/????
'    62346-STZ9  5      01/05/1932
'Format: (to csv with column re-order)
'    $id,$date,$count\n
'Result:
'    12345-STA1,10/02/2019,123
'    12323-STB9,01/01/2005,2123
'    23565-STC2,??/??/????,23
'    62346-STZ9,01/05/1932,5
Public Function List(ByVal sHaystack As String, ByVal sFormat As String) As String
  If Me Is stdRegex Then
    stdError.Raise ("Method called on class not object")
    Exit Function
  End If

  'Get all matches
  Dim matches As Collection
  Set matches = MatchAll(sHaystack)
  
  'Loop over all matches and list
  Dim match As variant, j As Long, sList As String, sTmpFormat As String
  For each match in matches
    'Concatenate to sList
    sList = sList & FormatFromMatch(match,sFormat)
  Next

  'Return sList
  List = sList
End Function

Public Function ListArr(ByVal sHaystack as string, vFormats as variant) as variant
  If Me Is stdRegex Then
    Call stdError.Raise("Method called on class not object")
    Exit Function
  End If

  'Get all matches
  Dim matches As Collection
  Set matches = MatchAll(sHaystack)
  
  Dim numCols as long: numCols = ubound(vFormats)-lbound(vFormats)+1
  
  Dim v() as variant
  Redim v(1 to matches.count, 1 to numCols)

  'Loop over all matches
  Dim match As variant, j As Long, iRow as long, iCol as long
  For each match in matches
    iRow = iRow + 1
    iCol = 0
    For j = lbound(vFormats) to ubound(vFormats)
      iCol = iCol + 1
      v(iRow,iCol) = FormatFromMatch(match,vFormats(j))
    Next
  Next

  'Return sList
  ListArr = v
End Function

Private Function FormatFromMatch(ByVal match as object, ByVal sFormat as string) as string
  'TODO: Make this more efficient, ideally we'd only loop through the format once

  'Get keys of each match
  Dim keys As Variant
  keys = match.keys()

  'Loop over default keys and make replacements
  sFormat = VBA.Replace(sFormat, "$&", match(0))
  sFormat = VBA.Replace(sFormat, "\\", "f934b47b-b799-4f52-9b4c-f170a82b52fe")
  sFormat = VBA.Replace(sFormat, "\r", Chr(13))
  sFormat = VBA.Replace(sFormat, "\n", Chr(10))
  sFormat = VBA.Replace(sFormat, "f934b47b-b799-4f52-9b4c-f170a82b52fe", "\")

  'Loop over remaining keys and make replacements
  For j = 1 To UBound(keys)
    Select Case keys(j)
      Case "$RAW", "$COUNT"
      Case Else
        sFormat = VBA.Replace(sFormat, "$" & keys(j), match(keys(j)))
    End Select
  Next

  FormatFromMatch = sFormat
End Function



Public Function UnitTests() As Boolean
  Dim b As Long: b = 0
  
  'Match tests (9)
  b = b + CLng(stdRegex.Create("(?<site>\d{5})-(?<equipment>\w{2})(\d+)").Match("12345-SP1")("site") = "12345")
  b = b + CLng(stdRegex.Create("(?<site>\d{5})-(?<equipment>\w{2})(\d+)").Match("12345-SP1")("equipment") = "SP")
  b = b + CLng(stdRegex.Create("(?<site>\d{5})-(?<equipment>\w{2})(\d+)").Match("12345-SP1")(0) = "12345-SP1")
  b = b + CLng(stdRegex.Create("(?<site>\d{5})-(?<equipment>\w{2})(\d+)").Match("12345-SP1")(1) = "12345")
  b = b + CLng(stdRegex.Create("(?<site>\d{5})-(?<equipment>\w{2})(\d+)").Match("12345-SP1")(2) = "SP")
  b = b + CLng(stdRegex.Create("(?<site>\d{5})-(?<equipment>\w{2})(\d+)").Match("12345-SP1")(3) = "1")
  b = b + CLng(stdRegex.Create("(?<site>\d{5})-(?<equipment>\w{2})(\d+)").Match("12345-SP1")("$COUNT") = 3)
  b = b + CLng(Not stdRegex.Create("(?<site>\d{5})-(?<equipment>\w{2})(\d+)").Match("12345-SP1")("$RAW") Is Nothing)
  b = b + CLng(stdRegex.Create("(?<site>\d{5})-(?<equipment>\w{2})(\d+)").Match("12345-SP1;62735-EF5")(1) = "12345")
  
  'MatchAll tests (3)
  b = b + CLng(stdRegex.Create("(?<site>\d{5})-(?<equipment>\w{2})(\d+)").MatchAll("12345-SP1;62735-EF5").Length = 2)
  b = b + CLng(stdRegex.Create("(?<site>\d{5})-(?<equipment>\w{2})(\d+)").MatchAll("12345-SP1;62735-EF5").item(1)("site") = "12345")
  b = b + CLng(stdRegex.Create("(?<site>\d{5})-(?<equipment>\w{2})(\d+)").MatchAll("12345-SP1;62735-EF5").item(2)("site") = "62735")
  
  'Replace tests (2)
  b = b + CLng(stdRegex.Create("(?<site>\d{5})-(?<equipment>\w{2})(\d+)").Replace("12345-SP1;62735-EF5", "$site") = "12345;62735-EF5")
  b = b + CLng(stdRegex.Create("(?<site>\d{5})-(?<equipment>\w{2})(\d+)", "g").Replace("12345-SP1;62735-EF5", "$site") = "12345;62735")
  
  'List tests (1)
  b = b + CLng(stdRegex.Create("(?<site>\d{5})-(?<equipment>\w{2})(\d+)").List("12345-SP1;62735-EF5", "$site-") = "12345-62735-")
  
  'Calculate percent complete
  Dim f As Double
  f = (-1 * b) / 15 * 100
  Debug.Print Round(f, 2) & "% successful."
  
  UnitTests = f = 100
End Function
