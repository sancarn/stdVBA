VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdRegex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Spec:
'This class has been designed to extend the feature set provided by VBScript.Regexp (or to eventually replace it entirely)
'Functions implemented on the class
'CONSTRUCTORS
'    [X] Create - With Cache
'    [X] init     #PROTECTED
'
'PROPERTIES
'    [X] Get/Let Pattern
'    [X] Get/Let Flags
'
'INSTANCE METHODS
'    [X] Test
'    [X] Match
'    [X] MatchAll
'    [X] Replace
'    [X] List
'    [X] ListArr
'    [ ] Compile - Can we compile for faster execution?
'    [?] NamedCaptures - See ruby
'    [?] Names - Get the names of the pattern
'
'REGEX ENGINE
'    [X] Implement Named Capturing Groups
'    [-] Flags Supported
'        [X] i - Case insensitive
'        [X] g - Global
'        [X] m - Multiline
'        [X] x - Ignore whitespace
'        [X] c - Ignore comment
'        [X] s - Singleline mode. `.` Matches all characters including whitepsace
'        [ ] n - Explicit capture. Do not capture unnamed groups
'    [ ] More symbols e.g. \A, \Z, ...
'    [ ] Rewrite the regex engine to implement more features, dependency reduction and Mac compatibility? https://swtch.com/~rsc/regexp/regexp1.html and https://deniskyashif.com/2019/02/17/implementing-a-regular-expression-engine/ 
'        [ ] Stream Support? - could be helpful with huge files
'    [ ] Back propogation
'    [ ] Positive and Negative lookahead/lookbehind
'
'OUT-OF-SCOPE
'    * Grep from file - unless there a smarter way than using rx.Match(stdShell.Read()) we won't be adding this to this class
'
'EXAMPLES
'# 1 - Printing codes from some haystack  
'
'   sHaystack = "12345-AA1, 12345-AB1,15233-AC3, 63234-ZD2"
'   set debugPrint = stdCallback.CreateFromModule("Main", "debugPrint")
'   stdEnumerator.Create(stdRegex.Create("(?<Site>\d{5})-(\w{2}\d)").MatchAll(sHaystack)).Map(stdLambda.Create("$1.item(""Site"")")).Unique.ForEach(debugPrint)
'
'   Sub debugPrint(ByVal s as string)
'       Debug.Print s
'   End Sub
'
'# 2 - Dumping array to excel:
'
'   sResult = "Here is some cool data:" & vbCrLf & _ 
'             "12345-STA1,10/02/2019,123" & vbCrLf & _ 
'             "12323-STB9,01/01/2005,2123" & vbCrLf & _ 
'             "and here is some more:" & vbCrLf & _ 
'             "23565-STC2,??/??/????,23" & vbCrLf & _ 
'             "62346-STZ9,01/05/1932,5" 
'   vResult = rx.ListArr(sHaystack, Array("$id-$date","$count"))
'   Range("A1:B1").value = Array("ID-Date","Count")
'   Range("A2").Resize(ubound(vResult,1),2).value = vResult

Private p_initialised As Boolean
Private p_pattern As String
Private p_cPattern As String
Private p_flags As String
Private p_namesDict As Object
Private p_regex As Object


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
    Call stdError.Raise("Constructor called on object not class")
    Exit Function
  End If
  
  'Cache regex objects for faster speed
  static cache as object: if cache is nothing then set cache = CreateObject("Scripting.Dictionary")
  Dim sKey as string: sKey = pattern & "&" & flags
  if not cache.exists(sKey) then
    'Call to [friend] init requires type information, can't use dict directly
    Dim r As stdRegex
    Set r = New stdRegex
    Call r.init(pattern, flags)
    
    'Set cache
    Set cache(sKey) = r
  end if

  'Return cached object
  set Create = cache(sKey)
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







'Get/Set pattern for regex
Public Property Get pattern() As String
  If Me Is stdRegex Then
    stdError.Raise ("Property called on class not object")
    Exit Property
  End If

  pattern = p_pattern
End Property
Public Property Let pattern(val As String)
  If Me Is stdRegex Then
    stdError.Raise ("Property called on class not object")
    Exit Property
  End If
  
  Call init(val, p_flags)
End Property



'Get/Set Flags for regex
Public Property Get flags() As String
  If Me Is stdRegex Then
    stdError.Raise ("Property called on class not object")
    Exit Property
  End If

  flags = p_flags
End Property
Public Property Let flags(val As String)
  If Me Is stdRegex Then
    stdError.Raise ("Property called on class not object")
    Exit Property
  End If

  Call init(p_pattern, val)
End Property



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



'Supplied with a haystack to search, ListArr will return a 2d array of data where the rows represent each match, and the columns represent the formats passed into vFormats parameter
'@param {ByVal String} Haystack to search
'@param {ByVal Array<String>} Array of formats to print into seperate columns. E.G. `Array("$id-$date","$detail")`
'@returns {Array[2d]<String>} All matches as a 2d array. Useful especially for dumping matches to Excel.
Public Function ListArr(ByVal sHaystack as string, ByVal vFormats as variant) as variant
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


'*******************
'* PRIVATE METHODS *
'*******************

'Supplied with a format string and match object, dump the required information from the match object to the string, and return it.
'@param {ByVal Dictionary} Match dictionary with numerical keys for each match, key 0 is whole match and also named matches.
'@param {ByVal string} Format to return data in. Indexed submatches to be extracted using $1, $2, ... . Named submatches to be extracted using $name syntax. \r\n will be 
'@returns {String} String formatted with data from match object.
Private Function FormatFromMatch(ByVal match as object, ByVal sFormat as string) as string
  'TODO: Make this more efficient, ideally we'd only loop through the format once

  'Get keys of each match
  Dim keys As Variant
  keys = match.keys()

  'Loop over default keys and make replacements
  sFormat = VBA.Replace(sFormat, "$&", match(0))
  sFormat = VBA.Replace(sFormat, "\\", "f934b47b-b799-4f52-9b4c-f170a82b52fe")
  sFormat = VBA.Replace(sFormat, "\$", "bfde0637-8207-418d-9d34-68ba974e8110")
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

  sFormat = VBA.Replace(sFormat, "bfde0637-8207-418d-9d34-68ba974e8110","$")
  FormatFromMatch = sFormat
End Function



'Escape chars within a string
'@param {ByVal String} Haystack to search
'@param {ByVal String} Char to escape
'@returns {String} Escaped string
Private Function SaveEscape(ByVal sHaystack As String, ByVal sChar As String) As String
  SaveEscape = VBA.Replace(VBA.Replace(sHaystack, "\\", ";f934b47b-b799-4f52-9b4c-f170a82b52fe;"), "\" & sChar, ";6632df85-1730-4159-a742-5b0c8c526ab8;")
End Function

'Convert a string with escaped chars to those without escaped chars
'@param {ByVal String} Haystack to search
'@param {ByVal String} Char to un-escape
'@returns {String} Unescaped string
Private Function LoadEscape(ByVal sHaystack As String, ByVal sChar As String) As String
  LoadEscape = VBA.Replace(VBA.Replace(sHaystack, ";f934b47b-b799-4f52-9b4c-f170a82b52fe;", "\\"), ";6632df85-1730-4159-a742-5b0c8c526ab8;", "\" & sChar)
End Function
