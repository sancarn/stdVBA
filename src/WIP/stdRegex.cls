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



private p_initialised as boolean
private p_pattern as string
private p_cPattern as string
private p_flags as string
private p_namesDict as object
private p_regex as object


'Get pattern for regex
Public Property Get Pattern() as string
  if Me is stdRegex then
    stdError.raise("Property called on class not object")
    Exit Property
  End If

  Pattern = p_pattern
End Property

'Set pattern for regex
Public Property Let Pattern(val as string)
  if Me is stdRegex then
    stdError.raise("Property called on class not object")
    Exit Property
  End If
  
  Call init(val,p_flags)
End Property




'Get Flags for regex
Public Property Get Flags() as string
  if Me is stdRegex then
    stdError.raise("Property called on class not object")
    Exit Property
  End If

  Flags = p_flags
End Property

'Set Flags for regex
Public Property Let Flags(val as string)
  if Me is stdRegex then
    stdError.raise("Property called on class not object")
    Exit Property
  End If

  Call init(p_pattern,val)
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
Public Function Create2(ByVal rx as string) as stdRegex
  Attribute Create2.VB_UserMemID = -5

  if not Me is stdRegex then
    stdError.raise("Constructor called on object not class")
    Exit Function
  End If

  'regex given like /asdf/flags
  Dim iEnd as long: iEnd = inStrRev(rx,"/")
  Dim sPattern as string: sPattern = mid(rx, 2, iEnd-2)
  Dim sFlags as string: sFlags = mid(rx, iEnd+1)
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
Public Function Create(ByVal pattern as string, ByVal flags as string) as stdRegex
  if not Me is stdRegex then
    stdError.raise("Constructor called on object not class")
    Exit Function
  End If
  
  Set Create = new stdRegex
  Create.init(pattern,flags)
End Function



'Initialises the class from within the static superclass. This method is meant for internal use only. Use at your own risk.
'@protected
'
'@param {string}  Pattern - The pattern to match
'@param {string}  Flags - Optional flags to apply
'@example
'    obj.init("A\d+","i")
Friend Sub init(ByVal pattern as string, ByVal flags as string)
  If Me is stdRegex then
    stdError.raise("Cannot run init on class")
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
    NamesMatcher.pattern = "(?<!\\)(?:\\\\)*\((?!(?:\?:|\?!|\?=|\?<=|\?<!))(?:\?\<(\w+?)\>)?"
    NamesMatcher.Global = True
  End If
  
  'Names replacer is used to strip the name syntax from the regular expression before it's injected into VBScript.Regexp.
  Static NamesReplacer as object
  If NamesReplacer is nothing then
    Set NamesReplacer = CreateObject("VBScript.Regexp")
    NamesReplacer.Pattern = "((?<!\\)(?:\\\\)*\()\?<(?:\w+)>"  'replace with $1
    NamesReplacer.Global = True
  end if

  'Ignore White Space option (x) will ignore all undeclared whitespace. It also makes # a comment
  'i.e. `/hello world\ and stuff #this is a comment/` --> `/helloworld andstuff/`
  Static OptionIgnoreWhiteSpace As Object
  If OptionIgnoreWhiteSpace Is Nothing Then
    Set OptionIgnoreWhiteSpace = CreateObject("VBScript.RegExp")
    OptionIgnoreWhiteSpace.pattern = "(?<!\\)(\\\\)*(?:\s+|#.*)"   'Replace with $1
    OptionIgnoreWhiteSpace.Global = True
  End If

  'SingleLineMode (s) forces period (.) to match every character (instead of every character apart from `\n`).
  'With this regex you can do:            OptionSingleLineMode.replace(sRegex,"$1(?:.|\\s)")
  'to inject the new regex characters into place.
  Static OptionSingleLineMode as object
  If OptionSingleLineMode is nothing then
    Set OptionSingleLineMode = CreateObject("VBScript.RegExp")
    OptionSingleLineMode.pattern = "(?<!\\)((?:\\\\)*)\."
    OptionSingleLineMode.Global = True
  end if

  'Create dictionary of names in the regex.
  set p_namesDict = NamesMatcher.Execute(p_pattern)

  'Initialise p_regex
  set p_regex = CreateObject("VBScript.RegExp")
  p_regex.pattern = NamesReplacer.replace(p_pattern,"$1")

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
        p_regex.pattern = OptionIgnoreWhiteSpace.Replace(p_regex.pattern, "$1")
      Case "s"
        p_regex.pattern = OptionSingleLineMode.Replace(p_regex.pattern, "$1(?:.|\\s)")
    End Select
  Next

  p_initialised = true
End Sub




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
Public Function Test(ByVal sHaystack as string) as Boolean
  If Me is stdRegex then
    stdError.raise("Method called on class not object")
    Exit Function
  End If
  
  Test = p_regex.test(sHaystack)
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
Public Function Match(ByVal sHaystack as string, Optional byval bAllMatches as boolean = false) as Object
  If Me is stdRegex then
    stdError.raise("Method called on class not object")
    Exit Function
  End If
  
  'Execute regex on haystack provided
  Dim oMatches As Object
  Set oMatches = p_regex.Execute(sHaystack)
  
  'Create dictionary to store this match's data
  Set oRet = CreateObject("Scripting.Dictionary")

  'Basic properties
  oRet(0) = oMatches(0)
  oRet("$COUNT") = oMatches(0).Count
  set oRet("$RAW") = oMatches

  'Loop over submatches and apply to dict
  Dim i as long, oRet As Object
  For i = 1 To p_namesDict.Count
    oRet(i) = oMatches(0).Submatches(i - 1)
    If Not IsEmpty(p_namesDict(i - 1).Submatches(0)) Then oRet(p_namesDict(i - 1).Submatches(0)) = oMatches(0).Submatches(i - 1)
  Next i
  
  'Return data
  Set Match = oRet
End Function

'Given a text string, return all strings which match the source regex pattern.
'
'@param {string}  sHaystack - Text to search for regex in.
'@return {stdArray<Dictionary>} An array of strings which match the regex.
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
Public Function MatchAll(ByVal sHaystack as string) as stdArray
  If Me is stdRegex then
    stdError.raise("Method called on class not object")
    Exit Function
  End If

  'Execute regex on haystack provided
  Dim oMatches As Object
  Set oMatches = p_regex.Execute(sHaystack)

  'Create array to hold match data
  Dim oRet as stdArray
  set oRet = stdArray.Create()
  
  'Loop over all matches
  Dim i as long, j as long, oRet As Object
  For i = 1 to oMatches.count
    'Create dictionary to store this match's data
    Dim oDict as object
    set oDict = CreateObject("Scripting.Dictionary")

    'Basic properties
    oDict(0) = oMatches(i - 1)
    oDict("$COUNT") = oMatches(i - 1).Count
    set oDict("$RAW") = oMatches

    'Loop over submatches and apply to dict
    For j = 1 to p_namesDict.count
      oDict(j) = oMatches(i - 1).Submatches(j - 1)
      If Not IsEmpty(p_namesDict(j - 1).Submatches(0)) Then oDict(p_namesDict(j - 1).Submatches(0)) = oMatches(i - 1).Submatches(j - 1)
    next

    'Push dictionary to array
    oRet.push(oDict)
  next

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
Public Function Replace(ByVal sHaystack as string, ByVal sReplacer as string) as String
  If Me is stdRegex then
    stdError.raise("Method called on class not object")
    Exit Function
  End If
  
  'Replace names in sReplacer with indexed replacers ($1,$2,...)
  For i = 1 to p_namesDict.count
    If Not IsEmpty(p_namesDict(i - 1).Submatches(0)) Then 
      Dim sName as string
      sName = p_namesDict(i - 1).Submatches(0)
      
      sReplacer = Replace(sReplacer,"$" & sName, "$" & i )
    end if
  next

  'Pass to vbscript regex handler and return
  Replace = p_regex.Replace(sHaystack,sReplacer)
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
Public Function List(ByVal sHaystack as string, ByVal sFormat as string) as String
  If Me is stdRegex then
    stdError.raise("Method called on class not object")
    Exit Function
  End If

  'Get all matches
  Dim arr as stdArray
  Set arr = MatchAll(sHaystack)
  
  'Realistically there is going to be more optimal ways of doing this, but for now i'm keeping it really simple.

  'Loop over all matches
  Dim i as long, sList as string, sTmpFormat as string
  For i = 1 to arr.length
    'Get keys of each match
    Dim keys as variant
    keys = arr(i).keys()

    'Loop over default keys and make replacements
    sTmpFormat = sFormat
    sTmpFormat = Replace(sTmpFormat,"$&", arr(i)(0))
    sTmpFormat = Replace(sTmpFormat,"\\", "f934b47b-b799-4f52-9b4c-f170a82b52fe")
    sTmpFormat = Replace(sTmpFormat,"\r", char(13))
    sTmpFormat = Replace(sTmpFormat,"\n", char(10))
    sTmpFormat = Replace(sTmpFormat,"f934b47b-b799-4f52-9b4c-f170a82b52fe","\")

    'Loop over remaining keys and make replacements
    for j = 1 to ubound(keys)
      select case keys(j)
        case "$RAW","$COUNT"
        case else
          sTmpFormat = replace(sTmpFormat, "$" & keys(j), arr(i)(keys(j)))
      end select
    next

    'Concatenate to sList
    sList = sList & sTmpFormat
  next

  'Return sList
  List = sList
End Function
