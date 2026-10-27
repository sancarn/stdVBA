VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdRegex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@module
'@description A class which uses sihlfall's pure VBA regex engine. Use this class to
'create regex objects with additional features like named captures, and use them to analyse text.
'@example A simple example of using the class:
'```vb
'Dim rx as stdRegex: set rx = stdRegex.Create("(?<Site>\d{5})-(\w{2}\d)")
'Debug.Print rx.Test("12345-AA1")  '=> True
'Debug.Print rx.Match("12345-AA1").Site  '=> 12345-AA1
'Debug.Print rx.MatchAll("12345-AA1,12345-AB1,15233-AC3,63234-ZD2")  '=> 12345-AA1,12345-AB1,15233-AC3,63234-ZD2
'Debug.Print rx.Replace("12345-AA1,12345-AB1,15233-AC3,63234-ZD2", "<Redacted>")  '=> "<Redacted>,<Redacted>,<Redacted>,<Redacted>"
'Debug.Print rx.List("12345-AA1,12345-AB1,15233-AC3,63234-ZD2", "$site\r\n")  '=> "12345\r\n12345\r\n15233\r\n63234\r\n"
'```

Private Type TThis
  initialised As Boolean
  pattern As String
  flags As String
  isGlobal As Boolean
  isMultiline As Boolean
  dotAll As Boolean
  regex As RegexTy
End Type
Private This As TThis


'Creates a regex object given a pattern and flags.
'@constructor
'
'@param pattern - The pattern to match
'@param flags   - Optional flags to apply
'@param stepsLimit - Maximum engine steps per candidate match (default 10000); exhaustion raises vbObjectError + 3051.
'@returns - Regular expression object
'@example ```vb
'stdRegex.Create("A\d+","i")
'```
Public Function Create(ByVal pattern As String, Optional ByVal flags As String = "", Optional ByVal stepsLimit As Long = 10000) As stdRegex
  If Not Me Is stdRegex Then
    Call CriticalRaise("Constructor called on object not class")
    Exit Function
  End If

  'Cache regex objects for faster speed
  Static cache As Object: If cache Is Nothing Then Set cache = CreateObject("Scripting.Dictionary")
  Dim sKey As String: sKey = Len(pattern) & ":" & pattern & Len(flags) & ":" & flags & ":" & stepsLimit
  If Not cache.exists(sKey) Then
    'Call to [friend] init requires type information, can't use dict directly
    Dim r As stdRegex
    Set r = New stdRegex
    Call r.protInit(pattern, flags, stepsLimit)

    'Set cache
    Set cache(sKey) = r
  End If

  'Return cached object
  Set Create = cache(sKey)
End Function



'Initialises the class from within the static superclass. This method is meant for internal use only. Use at your own risk.
'@protected
'@param pattern - The pattern to match
'@param flags - Optional flags to apply
'@example
'```vb
'obj.protInit("A\d+","i")
'```
Friend Sub protInit(ByVal pattern As String, ByVal flags As String, Optional ByVal stepsLimit As Long = 10000)
  If Me Is stdRegex Then
    CriticalRaise "Cannot run init on class"
    Exit Sub
  End If
  If stepsLimit <= 0 Then Err.Raise 5, "stdRegex", "stepsLimit must be positive"

  'Compile before updating instance state, so a syntax error preserves the old regex.
  Dim compiled As RegexTy, compiledPattern As String
  compiledPattern = PreparePattern(pattern, flags)
  stdRegex.protStaticRegexInitializeRegex compiled, compiledPattern, InStr(flags, "i") > 0
  compiled.stepsLimit = stepsLimit
  This.regex = compiled
  This.pattern = pattern
  This.flags = flags
  This.isGlobal = InStr(flags, "g") > 0
  This.isMultiline = InStr(flags, "m") > 0
  This.dotAll = InStr(flags, "s") > 0
  This.initialised = True
End Sub







'Get/Set pattern for regex
'@returns - Pattern for regex
Public Property Get pattern() As String
  If Me Is stdRegex Then
    CriticalRaise ("Property called on class not object")
    Exit Property
  End If

  pattern = This.pattern
End Property
Public Property Let pattern(val As String)
  If Me Is stdRegex Then
    CriticalRaise ("Property called on class not object")
    Exit Property
  End If

  Call protInit(val, This.flags, This.regex.stepsLimit)
End Property



'Get/Set Flags for regex
'@returns - Flags for regex
Public Property Get flags() As String
  If Me Is stdRegex Then
    CriticalRaise ("Property called on class not object")
    Exit Property
  End If

  flags = This.flags
End Property
Public Property Let flags(val As String)
  If Me Is stdRegex Then
    CriticalRaise ("Property called on class not object")
    Exit Property
  End If

  Call protInit(This.pattern, val, This.regex.stepsLimit)
End Property



'Given a text string, return whether the source regex is present
'@param sHaystack - Text to search for regex in.
'@returns - `True` if the regex pattern is present, `false` otherwise.
'@example
'For a regex of: `\d{5}-ST[A-Z]\d`
'With data:
'```
'* The site with id 12345-STA1 is one hell of a cool site.
'```
'The result is: `True`
Public Function Test(ByVal sHaystack As String) As Boolean
  If Me Is stdRegex Then
    CriticalRaise "Method called on class not object"
    Exit Function
  End If
  Dim state As MatcherStateTy
  state.multiline = This.isMultiline
  state.dotAll = This.dotAll
  Test = stdRegex.protStaticRegexMatchNext(state, This.regex, sHaystack, raiseOnLimit:=True)
End Function



'Given a text string, return the first match.
'@param  sHaystack - Text to search for regex in.
'@returns Object<Dictionary<string|number,string>> - A dictionary containing the match, submatches, named matches and Count
'@example
'For Regex: `(?<id>\d{5}-ST[A-Z]\d) - (?<desc>.*)` and data:
'```
'Some sites were in critical condition
'* 12345-STA1 - Large crack through pipe.
'* 12323-STB9 - Acid leakage polluting watercourse.
'and some others were largely ok:
'* 23565-STC2
'* 62346-STZ9
'```
'Produces result:
'```json
'{
'  0: "12345-STA1 - Large crack through pipe.",
'  "id":"12345-STA1",
'  1: "12345-STA1",
'  "desc": "Large crack through pipe."
'  2: "Large crack through pipe.",
'  "$COUNT":2
'}
'```
Public Function Match(ByVal sHaystack As String) As Object
  If Me Is stdRegex Then
    CriticalRaise "Method called on class not object"
    Exit Function
  End If
  Dim state As MatcherStateTy
  state.multiline = This.isMultiline
  state.dotAll = This.dotAll
  If stdRegex.protStaticRegexMatchNext(state, This.regex, sHaystack, raiseOnLimit:=True) Then
    Set Match = MakeCapturesDict(state.captures, sHaystack)
  End If
End Function



'Given a text string, return all strings which match the source regex pattern.
'@param sHaystack - Text to search for regex in.
'@returns Collection<Object<Dictionary<string|number,string>>> - An array of strings which match the regex.
'@example
'For regex: `\d{5}-ST[A-Z]\d` and data:
'```
'Some sites were in critical condition
'* 12345-STA1 - Large crack through pipe.
'* 12323-STB9 - Acid leakage polluting watercourse.
'and some others were largely ok:
'* 23565-STC2
'* 62346-STZ9
'```
'Will result in an array of matches like:
'```json
'[
'  {0: "12345-STA1", ...},
'  {0: "12323-STB9", ...},
'  {0: "23565-STC2", ...},
'  {0: "62346-STZ9", ...}
']
'```
Public Function MatchAll(ByVal sHaystack As String) As Collection
  If Me Is stdRegex Then
    CriticalRaise "Method called on class not object"
    Exit Function
  End If
  Dim state As MatcherStateTy, results As New Collection
  state.multiline = This.isMultiline
  state.dotAll = This.dotAll
  Do While stdRegex.protStaticRegexMatchNext(state, This.regex, sHaystack, raiseOnLimit:=True)
    results.Add MakeCapturesDict(state.captures, sHaystack)
  Loop
  Set MatchAll = results
End Function



'Given a Haystack and a Replacer, Replace all matches with the source regex with the format given in the Replacer.
'
'@param sHaystack - Text to search for regex in.
'@param sReplacer - Format of replacement text.
'@returns - The haystack with replaced details.
'@example
'For regex: `(?<id>\d{5}-ST[A-Z]\d))\s+(?<count>\d+)\s+(?<date>../../....)` and data:
'```
'Here is some cool data:
'12345-STA1  123    10/02/2019
'12323-STB9  2123   01/01/2005
'and here is some more:
'23565-STC2  23     ??/??/????
'62346-STZ9  5      01/05/1932
'```
'With replacer `$id,$date,$count` will result in:
'```
'Here is some cool data:
'12345-STA1,10/02/2019,123
'12323-STB9,01/01/2005,2123
'and here is some more:
'23565-STC2,??/??/????,23
'62346-STZ9,01/05/1932,5
'```
Public Function Replace(ByVal sHaystack As String, ByVal sReplacer As String) As String
  If Me Is stdRegex Then
    CriticalRaise "Method called on class not object"
    Exit Function
  End If
  sReplacer = PrepareReplacement(sReplacer)
  Replace = stdRegex.protStaticRegexReplace(This.regex, sReplacer, sHaystack, _
    localMatch:=Not This.isGlobal, multiline:=This.isMultiline, dotAll:=This.dotAll)
End Function



'Given a Haystack and a Format, Generate a string containing all matches in the format provided.
'
'@param  sHaystack - Text to search for regex in.
'@param  sFormat   - Format of output list.
'@returns - A list of strings matched.
'@example
'For regex: `(?<id>\d{5}-ST[A-Z]\d))\s+(?<count>\d+)\s+(?<date>../../....)` and data:
'```
'12345-STA1  123    10/02/2019
'12323-STB9  2123   01/01/2005
'23565-STC2  23     ??/??/????
'62346-STZ9  5      01/05/1932
'```
'With format: `$id,$date,$count\n` will result in:
'```
'12345-STA1,10/02/2019,123
'12323-STB9,01/01/2005,2123
'23565-STC2,??/??/????,23
'62346-STZ9,01/05/1932,5
'```
Public Function List(ByVal sHaystack As String, ByVal sFormat As String) As String
  If Me Is stdRegex Then
    CriticalRaise ("Method called on class not object")
    Exit Function
  End If

  'Get all matches
  Dim matches As Collection
  Set matches = MatchAll(sHaystack)

  'Loop over all matches and list
  Dim match As Variant, j As Long, sList As String, sTmpFormat As String
  For Each match In matches
    'Concatenate to sList
    sList = sList & FormatFromMatch(match, sFormat)
  Next

  'Return sList
  List = sList
End Function



'Supplied with a haystack to search, ListArr will return a 2d array of data where the rows represent each match, and the columns represent the formats passed into vFormats parameter
'@param sHaystack - Haystack to search
'@param vFormats as Variant<Array<String>> - Array of formats to print into seperate columns. E.G. `Array("$id-$date","$detail")`
'@returns Variant<Array[2d]<String>> - All matches as a 2d array. Useful especially for dumping matches to Excel.
Public Function ListArr(ByVal sHaystack As String, ByVal vFormats As Variant) As Variant
  If Me Is stdRegex Then
    Call CriticalRaise("Method called on class not object")
    Exit Function
  End If

  'Get all matches
  Dim matches As Collection
  Set matches = MatchAll(sHaystack)

  Dim numCols As Long: numCols = UBound(vFormats) - LBound(vFormats) + 1

  Dim v() As Variant
  If matches.Count = 0 Then
    ListArr = Array()
    Exit Function
  End If
  ReDim v(1 To matches.Count, 1 To numCols)

  'Loop over all matches
  Dim match As Variant, j As Long, iRow As Long, iCol As Long
  For Each match In matches
    iRow = iRow + 1
    iCol = 0
    For j = LBound(vFormats) To UBound(vFormats)
      iCol = iCol + 1
      v(iRow, iCol) = FormatFromMatch(match, vFormats(j))
    Next
  Next

  'Return sList
  ListArr = v
End Function


'*******************
'* PRIVATE METHODS *
'*******************

'Supplied with a format string and match object, dump the required information from the match object to the string, and return it.
'@param match as Object<Dictionary> - Match dictionary with numerical keys for each match, key 0 is whole match and also named matches.
'@param sFormat - Format to return data in. Indexed submatches to be extracted using $1, $2, ... . Named submatches to be extracted using $name syntax. \r\n will be
'@returns - String formatted with data from match object.
Private Function FormatFromMatch(ByVal match As Object, ByVal sFormat As String) As String
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
  Dim j As Long
  For j = 1 To UBound(keys)
    Select Case keys(j)
      Case "$RAW", "$COUNT"
      Case Else
        sFormat = VBA.Replace(sFormat, "$" & keys(j), match(keys(j)))
    End Select
  Next

  sFormat = VBA.Replace(sFormat, "bfde0637-8207-418d-9d34-68ba974e8110", "$")
  FormatFromMatch = sFormat
End Function



'Raise an error with the class name as the source
'@param sMessage - Error message
Private Sub CriticalRaise(ByVal sMessage As String)
  Err.Raise 1, "stdRegex", sMessage
End Sub

Private Function MakeCapturesDict(ByRef captures As CapturesTy, ByRef sHaystack As String) As Object
    Dim j As Long, identifier As String, referenced As Long
    Dim identifierMapEnd As Long, currentEntry As Long
    Dim oRet As Object

    Set oRet = CreateObject("Scripting.Dictionary")

    With This
        oRet("$COUNT") = captures.nNumberedCaptures

        With captures.entireMatch
            If .start > 0 Then oRet(0) = Mid$(sHaystack, .start, .length) Else oRet(0) = vbNullString
        End With
        For j = 1 To captures.nNumberedCaptures
            With captures.numberedCaptures(j - 1)
                If .start > 0 Then oRet(j) = Mid$(sHaystack, .start, .length) Else oRet(j) = vbNullString
            End With
        Next

        identifierMapEnd = BYTECODE_IDENTIFIER_MAP_BEGIN + _
            BYTECODE_IDENTIFIER_MAP_ENTRY_SIZE * .regex.bytecode(BYTECODE_IDX_N_IDENTIFIERS)
        currentEntry = BYTECODE_IDENTIFIER_MAP_BEGIN
        Do While currentEntry < identifierMapEnd
            identifier = Mid$( _
                .regex.pattern, _
                .regex.bytecode(currentEntry + BYTECODE_IDENTIFIER_MAP_ENTRY_START_IN_PATTERN), _
                .regex.bytecode(currentEntry + BYTECODE_IDENTIFIER_MAP_ENTRY_LENGTH_IN_PATTERN) _
            )
            referenced = captures.namedCaptures( _
                .regex.bytecode(currentEntry + BYTECODE_IDENTIFIER_MAP_ENTRY_ID) _
            )
            If referenced <> -1 Then
                oRet(identifier) = oRet(referenced)
            Else
                oRet(identifier) = vbNullString
            End If
            currentEntry = currentEntry + BYTECODE_IDENTIFIER_MAP_ENTRY_SIZE
        Loop
    End With

    Set MakeCapturesDict = oRet
End Function


'Apply the legacy x (spaces) and c (line comments) flags without using regex COM.
'Escaped characters and character classes are copied unchanged.
Private Function PreparePattern(ByVal pattern As String, ByVal flags As String) As String
  Dim i As Long, ch As String, inClass As Boolean, result As String
  Dim ignoreSpaces As Boolean, ignoreComments As Boolean
  ignoreSpaces = InStr(flags, "x") > 0
  ignoreComments = InStr(flags, "c") > 0
  i = 1
  Do While i <= Len(pattern)
    ch = Mid$(pattern, i, 1)
    If ch = "\" Then
      result = result & Mid$(pattern, i, 2)
      i = i + 2
    ElseIf ch = "#" And ignoreComments And Not inClass Then
      Do While i <= Len(pattern)
        If Mid$(pattern, i, 1) = vbCr Or Mid$(pattern, i, 1) = vbLf Then Exit Do
        i = i + 1
      Loop
    Else
      If ch = "[" Then inClass = True
      If ch = "]" Then inClass = False
      If Not (ch = " " And ignoreSpaces And Not inClass) Then result = result & ch
      i = i + 1
    End If
  Loop
  PreparePattern = result
End Function

'Translate legacy $name references to the engine's $<name> syntax in one pass.
'A doubled dollar is already escaped; literal dollar signs remain literal.
Private Function PrepareReplacement(ByVal format As String) As String
  Dim i As Long, j As Long, ch As String, name As String, result As String
  i = 1
  Do While i <= Len(format)
    ch = Mid$(format, i, 1)
    If ch <> "$" Then
      result = result & ch
      i = i + 1
    Else
      ch = Mid$(format, i + 1, 1)
      Select Case ch
        Case "$", "&", "`", "'", "<", "~", "0" To "9"
          result = result & "$" & ch
          i = i + 2
        Case Else
          j = i + 1
          Do While j <= Len(format)
            ch = Mid$(format, j, 1)
            If Not (ch Like "[A-Za-z0-9_]") Then Exit Do
            j = j + 1
          Loop
          name = Mid$(format, i + 1, j - i - 1)
          If Len(name) > 0 Then
            If GetIdentifierId(This.regex.bytecode, This.regex.pattern, name) >= 0 Then
              result = result & "$<" & name & ">"
            Else
              result = result & "$$" & name
            End If
          Else
            result = result & "$$"
          End If
          i = j
      End Select
    End If
  Loop
  PrepareReplacement = result
End Function
