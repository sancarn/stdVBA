VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdRegex2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type QuantifierBounds
    UB As Long
    LB As Long
End Type
Private Type Buffer
    Haystack As String
    index As Long
    length As Long
End Type
Private Type RegexFlags
    caseSensitive As Boolean
    multiline As Boolean
    global As Boolean
End Type
Private Type TGenData
    maxRepitionLength As Long
End Type
Private Type TThis
    pattern As String
    ast As Object
    Flags As RegexFlags
    GenData As TGenData
End Type
Private This As TThis

'Creates a new regex object
'@param Pattern - The regex pattern
'@param Flags - The regex flags
'@returns - The regex object
'@remark - The flags are:
' - `i` - Case insensitive
' - `m` - Multiline
' - `g` - Global
Public Function Create(ByVal pattern As String, ByVal Flags As String) As stdRegex2
    Set Create = New stdRegex2
    Call Create.protInit(pattern, Flags)
End Function

'Initializes the regex object
'@protected
'@param Pattern - The regex pattern
'@param Flags - The regex flags
'@remark - This is a protected method and should not be called directly. Use `Create` instead.
Public Sub protInit(ByVal pattern As String, ByVal Flags As String)
    Set This.ast = Parse(pattern)
End Sub

Public Property Get pattern() As String
    pattern = This.pattern
End Property

Public Property Get Flags() As String
    Flags = iif(This.flags.caseSensitive, "i", "") & iif(This.flags.multiline, "m", "") & iif(This.flags.global, "g", "")
End Property

Public Property Get ast() As Object
    Set ast = This.ast
End Property

'Generate a string which complies with the regex
'@param maxRepitionLength - The maximum number of times a quantifier can repeat when `*` or `+` is used (default is 10)
'@return - A string which complies with the regex
Public Function Generate(Optional ByVal maxRepitionLength as long=10) As String
    This.GenData.maxRepitionLength = maxRepitionLength
    Call GenerateForRegex(Generate, This.ast)
End Function

'Test whether a string matches the regex
'@param str - The string to test
'@return - `True` if the string matches the regex, `False` otherwise
'@TODO: Implement
Private Function Test(ByVal str as string) as Boolean

End Function

'Execute the regex against a string
'@param str - The string to execute the regex against
'@return - A collection of matches
'@TODO: Implement
Private Function Match(ByVal str as string) As Object

End Function

'@TODO: Implement
Private Function MatchFirst(ByVal str as string, Optional ByVal iStart as Long = 0) As Object
    
End Function

'@TODO: Implement
Private Function Replace(ByVal str as string, ByVal replacement as string) As string
    
End Function

'Obtain a list of matches with a single format
'@TODO: Implement
'Private Function List

'Obtain a 2d table of matches with seperate formats for each column
'@TODO: Implement
Private Function List2D(ByVal str as string, ParamArray formats()) as variant

End Function

'@TODO: Implement
'@param buff - The buffer to match against, could link to a file or a string
'@param astRegex - The regex to match
'@param iStart - The index to start matching from
'@return - An individual match which is the first match found after iStart
Private Function matchRegex(ByRef buff as HaystackBuffer, ByVal astRegex as Object, ByVal iStart as Long) as TRegexMatch
    
End Sub





'**********************************************************
'* PARSING
'**********************************************************
'Summary of parsing a regex expression
'Regex = Branch+
'Branch = Piece+                           'Branch ==> ORs
'Piece = atom quantifier
'quantifier = quantZeroPlus | quantOnePlus | quantZeroOrOne | quantBound
'quantBound = "{" \d (?:"," \d?)?  "}"
'atom = "(?:" Regex ")" |
'       "(?=" Regex ")" |
'       "(?<" .* ">" Regex ")" |
'       "(?!" Regex ")" |
'       "(?<!" Regex ")" |
'       "(?<=" Regex ")" |
'       "(" Regex ")" |
'       "(" ")" |
'       BOF | EOF |
'       escapedChar |
'       anyChar |
'       charGroup
'charGroup = "[" .* "]"
'escapedChar = \\.
'anyChar = .
'BOF = ^
'EOF = $

'Returns an AST for a given regex
'@param regex - The regex to parse
'@returns - The AST
Private Function Parse(ByVal regex As String) As Object
    Dim buff As Buffer
    buff.Haystack = regex
    buff.index = 1
    buff.length = Len(regex)
    Set Parse = parserRegex(buff)
End Function

'Parses a regex
'@param buff - The buffer to parse from
'@returns - The AST
Private Function parserRegex(ByRef buff As Buffer, Optional ByVal captureDepth As Long = 0, Optional ByVal captureCount As Long = 0) As Object
    Dim ast As Object
    Set ast = CreateDictionary("type", "regex", "start", buff.index, "end", 0, "value", New Collection)
    Do While parserHasNext(buff) And Not parserPeek(buff, ")")
        ast("value").Add parserBranch(buff, captureDepth, captureCount)
    Loop
    ast("end") = buff.index
    ast("text") = mid(buff.Haystack, ast("start"), ast("end") - ast("start"))
    Set parserRegex = ast
End Function

'Parses a branch
'@param buff - The buffer to parse from
'@returns - The AST
'@remark - A branch is a collection of pieces, which are OR'd together with `|`
Private Function parserBranch(ByRef buff As Buffer, ByVal captureDepth As Long, ByVal captureCount As Long) As Object
    Dim ast As Object
    Set ast = CreateDictionary("type", "branch", "start", buff.index, "end", 0, "value", New Collection)
    While parserHasNext(buff) And Not parserPeek(buff, ")") And Not parserPeek(buff, "|")
        ast("value").Add parserPiece(buff, captureDepth, captureCount)
    Wend
    Call parserOptionalConsume(buff, "|")
    
    ast("end") = buff.index
    ast("text") = mid(buff.Haystack, ast("start"), ast("end") - ast("start"))
    Set parserBranch = ast
End Function

'Parses a piece
'@param buff - The buffer to parse from
'@returns - The AST
'@remark - A piece is an atom with a quantifier
Private Function parserPiece(ByRef buff As Buffer, ByVal captureDepth As Long, ByVal captureCount As Long) As Object
    Dim ast As Object
    Set ast = CreateDictionary("type", "piece", "start", buff.index, "end", 0)
    ast.Add "value", parserAtom(buff, captureDepth, captureCount)
    
    If parserOptionalConsume(buff, "*") Then
        ast.Add "quantifier", "zeroPlus"
        ast.Add "greedy", Not parserOptionalConsume(buff, "?")
    ElseIf parserOptionalConsume(buff, "+") Then
        ast.Add "quantifier", "onePlus"
        ast.Add "greedy", Not parserOptionalConsume(buff, "?")
    ElseIf parserOptionalConsume(buff, "?") Then
        ast.Add "quantifier", "zeroOrOne"
        ast.Add "greedy", False
    ElseIf parserOptionalConsume(buff, "{") Then
        ast.Add "quantifier", "bound"
        ast.Add "greedy", False
        ast.Add "min", parserConsumeInt(buff)
        If parserOptionalConsume(buff, ",") Then
            ast.Add "max", parserConsumeInt(buff)
        Else
            ast.Add "max", ast("min")
        End If
        Call parserConsume(buff, "}")
    Else
        ast.Add "quantifier", "one"
        ast.Add "greedy", False
    End If

    ast("end") = buff.index
    ast("text") = mid(buff.Haystack, ast("start"), ast("end") - ast("start"))
    Set parserPiece = ast
End Function

'Parses an atom
'@param buff - The buffer to parse from
'@returns - The AST
'@remark - An atom is a single character, a group of characters or a whole expression within parentheses
Private Function parserAtom(ByRef buff As Buffer, ByVal captureDepth As Long, ByVal captureCount As Long) As Object
    Dim ast As Object
    Set ast = CreateDictionary("type", "atom", "start", buff.index, "end", 0)

    Dim isEmpty As Boolean: isEmpty = False
    If parserOptionalConsume(buff, "(") Then
        Select Case True
            Case parserOptionalConsume(buff, "?<!")
                ast.Add "atomType", "negative-lookbehind"
            Case parserOptionalConsume(buff, "?<=")
                ast.Add "atomType", "positive-lookbehind"
            Case parserOptionalConsume(buff, "?:")
                ast.Add "atomType", "non-capturing-group"
            Case parserOptionalConsume(buff, "?=")
                ast.Add "atomType", "positive-lookahead"
            Case parserOptionalConsume(buff, "?<")
                ast.Add "atomType", "named-capturing-group"
                ast.Add "name", parserConsumeUntil(buff, ">")
                captureDepth = captureDepth + 1
                ast.Add "captureDepth", captureDepth
                captureCount = captureCount + 1
                ast.Add "captureIndex", captureCount
                Call parserConsume(buff, ">")
            Case parserOptionalConsume(buff, "?!")
                ast.Add "atomType", "negative-lookahead"
            Case Else
                ast.Add "atomType", "capturing-group"
                captureDepth = captureDepth + 1
                ast.Add "captureDepth", captureDepth
                captureCount = captureCount + 1
                ast.Add "captureIndex", captureCount
        End Select
        '
        ast.Add "value", parserRegex(buff, captureDepth, captureCount)
        captureDepth = captureDepth - 1
        Call parserConsume(buff, ")")
    ElseIf parserOptionalConsume(buff, "^") Then
        ast.Add "atomType", "anchor"
        ast.Add "value", "BOF"
    ElseIf parserOptionalConsume(buff, "$") Then
        ast.Add "atomType", "anchor"
        ast.Add "value", "EOF"
    ElseIf parserOptionalConsume(buff, "\") Then
        Call parserHelperEscape(buff, ast)
    ElseIf parserOptionalConsume(buff, ".") Then
        ast.Add "atomType", "charsets"
        ast.Add "value", createCharSets(createCharSet(0, &HFFFF&))
    ElseIf parserOptionalConsume(buff, "[") Then
        ast.Add "atomType", "charsets"
        If parserOptionalConsume(buff, "^") Then
            ast.Add "value", invertCharSets(parserCharGroup(buff))
        Else
            ast.Add "value", parserCharGroup(buff)
        End If
        Call parserConsume(buff, "]")
    ElseIf Not parserPeek(buff, ")") Then
        ast.Add "atomType", "literal"
        ast.Add "value", parserConsumeLength(buff, 1)
    Else
        ast.Add "atomType", "empty"
        ast.Add "value", "EmptyString"
    End If

    ast("end") = buff.index
    ast("text") = mid(buff.Haystack, ast("start"), ast("end") - ast("start"))
    Set parserAtom = ast
End Function

'Parses an escaped character
'@param buff - The buffer to parse from
'@param ast as Object<Dictionary<string, variant>> - The AST
'@remark - An escaped character is a character preceded by a backslash. This function assumes the backslash has already been consumed.
Private Sub parserHelperEscape(ByRef buff As Buffer, ByRef ast As Object)
    If parserOptionalConsume(buff, "b") Then
        ast.Add "atomType", "anchor"
        ast.Add "value", "wordBoundary"
    ElseIf parserOptionalConsume(buff, "B") Then
        ast.Add "atomType", "anchor"
        ast.Add "value", "notWordBoundary"
    ElseIf parserOptionalConsume(buff, "d") Then
        ast.Add "atomType", "charsets"
        ast.Add "value", createCharSets(createCharSetDigit())
    ElseIf parserOptionalConsume(buff, "w") Then
        ast.Add "atomType", "charsets"
        ast.Add "value", createCharSets_W()
    ElseIf parserOptionalConsume(buff, "s") Then
        ast.Add "atomType", "charsets"
        ast.Add "value", createCharSet_S()
    ElseIf parserOptionalConsume(buff, "D") Then
        ast.Add "atomType", "charsets"
        ast.Add "value", invertCharSets(createCharSets(createCharSetDigit()))
    ElseIf parserOptionalConsume(buff, "W") Then
        ast.Add "atomType", "charsets"
        ast.Add "value", invertCharSets(createCharSets(createCharSets_W()))
    ElseIf parserOptionalConsume(buff, "S") Then
        ast.Add "atomType", "charsets"
        ast.Add "value", invertCharSets(createCharSets(createCharSet_S()))
    ElseIf parserOptionalConsume(buff, "r") Then
        ast.Add "atomType", "literal"
        ast.Add "value", vbCr
    ElseIf parserOptionalConsume(buff, "n") Then
        ast.Add "atomType", "literal"
        ast.Add "value", vbLf
    ElseIf parserOptionalConsume(buff, "t") Then
        ast.Add "atomType", "literal"
        ast.Add "value", vbTab
    ElseIf parserOptionalConsume(buff, "f") Then
        ast.Add "atomType", "literal"
        ast.Add "value", vbFormFeed
    ElseIf parserOptionalConsume(buff, "v") Then
        ast.Add "atomType", "literal"
        ast.Add "value", vbVerticalTab
    ElseIf parserOptionalConsume(buff, "x") Then
        ast.Add "atomType", "literal"
        ast.Add "value", Chr("&h" & parserConsumeLength(buff, 2))
    ElseIf parserOptionalConsume(buff, "u") Then
        ast.Add "atomType", "literal"
        ast.Add "value", ChrW("&h" & parserConsumeLength(buff, 4))
    ElseIf parserOptionalConsume(buff, "c") Then
        ast.Add "atomType", "literal"
        Dim sChar As String: sChar = ucase(parserConsumeLength(buff, 1))
        If Asc(sChar) >= Asc("A") And Asc(sChar) <= Asc("Z") Then
            ast.Add "value", Chr(Asc(sChar) - Asc("A") + 1)
        Else
            Err.Raise 5, "parserAtom", "Expected control character at position " & buff.index & " but found '" & mid(buff.Haystack, buff.index, 1) & "'"
        End If
    ElseIf parserOptionalConsume(buff, "k<") Then
        ast.Add "atomType", "backreference-named"
        ast.Add "value", parserConsumeUntil(buff, ">")
        Call parserConsume(buff, ">")
    ElseIf parserPeekLength(buff, 1) Like "[0123456789]" Then
        ast.Add "atomType", "backreference-numeric"
        ast.Add "value", parserConsumeInt(buff)
    ElseIf parserOptionalConsume(buff, "o") Then 'Homebrew: Octal
        ast.Add "atomType", "literal"
        ast.Add "value", ChrW("&o" & parserConsumeLength(buff, 3))
    Else
        ast.Add "atomType", "literal"
        ast.Add "value", parserConsumeLength(buff, 1)
    End If
End Sub

'Parses a character group
'@param buff - The buffer to parse from
'@returns - collection of character sets
'@remark - A character group is a collection of character sets, which are AND'd together
Private Function parserCharGroup(ByRef buff As Buffer) As Collection
    Dim ast As Collection
    Set ast = New Collection
    Do While parserHasNext(buff) And Not parserPeek(buff, "]")
        If parserOptionalConsume(buff, "\") Then
            If parserOptionalConsume(buff, "b") Then
                ast.Add createCharSet(0, 8)
            Else 'delegate to escape helper
                Dim atom As Object: Set atom = CreateDictionary("type", "atom")
                Call parserHelperEscape(buff, atom)
                Select Case atom("type")
                    Case "literal"
                        ast.Add createCharSet(AscW(atom("value")))
                    Case "charsets"
                        Dim charset As Object
                        For Each charset In atom("value")
                            ast.Add charset
                        Next
                End Select
            End If
        Else
            If parserPeek(buff, "-", 1) Then
                Dim minChar As String: minChar = parserConsumeLength(buff, 1)
                Call parserConsume(buff, "-")
                Dim maxChar As String: maxChar = parserConsumeLength(buff, 1)
                ast.Add createCharSet(AscW(minChar), AscW(maxChar))
            Else
                ast.Add createCharSet(AscW(parserConsumeLength(buff, 1)))
            End If
        End If
    Loop

    'Sort collection by character set min character from lowest to highest
    Call sortCollectionOn(ast, "charMin")
    Set parserCharGroup = ast
End Function

'Create a character set
'@param charMin - The minimum character
'@param charMax - The maximum character
'@returns - The character set
Private Function createCharSet(ByVal charMin As Long, Optional ByVal charMax As Long = -1) As Object
    If charMax = -1 Then charMax = charMin

    Dim ast As Object
    Set ast = CreateDictionary("type", "charset", "charMin", charMin, "charMax", charMax)
    Set createCharSet = ast
End Function

'Create a charsets object from a number of character sets
'@param charsets - The character sets to add to the charsets object
'@returns - The charsets object
Private Function createCharSets(ParamArray charsets()) As Collection
    Dim ast As Collection: Set ast = New Collection
    Dim i As Long
    For i = 0 To UBound(charsets)
        ast.Add charsets(i)
    Next
    Call sortCollectionOn(ast, "charMin")

    Set createCharSets = ast
End Function

'Invert a set of character sets
'@param charsets - The character set to invert
'@returns - A collection of character sets which match the inverse of the specified character sets
Private Function invertCharSets(ByVal charsets As Collection) As Collection
    Dim ret As Collection: Set ret = New Collection
    
    Dim oldCharMax As Long: oldCharMax = -1

    Dim charset As Object
    For Each charset In charsets
        Dim charMin As Long: charMin = charset("charMin")
        If charMin > oldCharMax + 1 Then
            ret.Add createCharSet(oldCharMax + 1, charMin - 1)
        End If
        
        oldCharMax = charset("charMax")
    Next
    If oldCharMax < &HFFFF& Then ret.Add createCharSet(oldCharMax + 1, &HFFFF&)
    Set invertCharSets = ret
End Function

'Create a character set for word characters
'@returns - The character set
Private Function createCharSets_W() As Collection
    Static cache As Collection
    If cache Is Nothing Then
        Set cache = createCharSets( _
            createCharSetAlphabetUCase(), _
            createCharSetAlphabetLCase(), _
            createCharSetDigit(), _
            createCharSet(AscW("_")) _
        )
    End If
    Set createCharSets_W = cache
End Function

'Create a character set for whitespace characters
'@returns - The character set
Private Function createCharSet_S() As Collection
    Static cache As Collection
    If cache Is Nothing Then
        Set cache = createCharSets( _
            createCharSet(AscW(" ")), _
            createCharSet(AscW(vbTab)), _
            createCharSet(AscW(vbCr)), _
            createCharSet(AscW(vbLf)) _
        )
    End If
    Set createCharSet_S = cache
End Function

'Create a character set for uppercase alphabet letters
'@returns - The character set
Private Function createCharSetAlphabetUCase() As Object
    Static cache As Object: If cache Is Nothing Then Set cache = createCharSet(AscW("A"), AscW("Z"))
    Set createCharSetAlphabetUCase = cache
End Function

'Create a character set for lowercase alphabet letters
'@returns - The character set
Private Function createCharSetAlphabetLCase() As Object
    Static cache As Object: If cache Is Nothing Then Set cache = createCharSet(AscW("a"), AscW("z"))
    Set createCharSetAlphabetLCase = cache
End Function

'Create a character set for digits
'@returns - The character set
Private Function createCharSetDigit() As Object
    Static cache As Object: If cache Is Nothing Then Set cache = createCharSet(AscW("0"), AscW("9"))
    Set createCharSetDigit = cache
End Function




'Consume until a given string is found
'@param buff - The buffer to consume from
'@param str - The string to consume until
'@returns - The characters consumed
Private Function parserConsumeUntil(ByRef buff As Buffer, ByVal str As String) As String
    Dim index As Long
    index = InStr(buff.index, buff.Haystack, str)
    If index = 0 Then
        Err.Raise 5, "parserConsumeUntil", "Expected '" & str & "' at position " & buff.index & " but found '" & mid(buff.Haystack, buff.index, Len(str)) & "'"
    End If
    parserConsumeUntil = mid(buff.Haystack, buff.index, index - buff.index)
    buff.index = index
End Function

'Consume a string
'@param buff - The buffer to consume from
'@param str - The string to consume
'@throws - Error if the string is not found
Private Sub parserConsume(ByRef buff As Buffer, ByVal str As String)
    If mid(buff.Haystack, buff.index, Len(str)) = str Then
        buff.index = buff.index + Len(str)
    Else
        Err.Raise 5, "parserConsume", "Expected '" & str & "' at position " & buff.index & " but found '" & mid(buff.Haystack, buff.index, Len(str)) & "'"
    End If
End Sub

'Optional consume a string
'@param buff - The buffer to consume from
'@param str - The string to consume
'@returns - `True` if the string was consumed, `False` otherwise
Private Function parserOptionalConsume(ByRef buff As Buffer, ByVal str As String) As Boolean
    If mid(buff.Haystack, buff.index, Len(str)) = str Then
        buff.index = buff.index + Len(str)
        parserOptionalConsume = True
    Else
        parserOptionalConsume = False
    End If
End Function

'Consume a fixed number of characters
'@param buff - The buffer to consume from
'@param length - The number of characters to consume
Private Function parserConsumeLength(ByRef buff As Buffer, ByVal length As Long) As String
    parserConsumeLength = mid(buff.Haystack, buff.index, length)
    buff.index = buff.index + length
End Function

'Peek whether a given string is at the current position
'@param buff - The buffer to peek from
'@param str - The string to peek for
'@returns - `True` if the string is at the current position, `False` otherwise
Private Function parserPeek(ByRef buff As Buffer, ByVal str As String, Optional ByVal offset As Long = 0) As Boolean
    parserPeek = mid(buff.Haystack, buff.index + offset, Len(str)) = str
End Function

'Peek at a fixed number of characters
'@param buff - The buffer to peek from
'@param length - The number of characters to peek
'@returns - The characters peeked
Private Function parserPeekLength(ByRef buff As Buffer, ByVal length As String) As String
    parserPeekLength = mid(buff.Haystack, buff.index, length)
End Function

'Parse an integer
'@param buff - The buffer to parse from
'@returns - The integer
Private Function parserConsumeInt(ByRef buff As Buffer) As Long
    Dim s As String: s = ""
    Do While parserPeekLength(buff, 1) Like "[0123456789]"
        s = s & parserConsumeLength(buff, 1)
    Loop
    parserConsumeInt = CLng(s)
End Function

'Check whether there are more characters yet to parse
'@param buff - The buffer to check
'@returns - `True` if there are more characters yet to parse, `False` otherwise
Private Function parserHasNext(ByRef buff As Buffer) As Boolean
    parserHasNext = buff.index <= buff.length
End Function

'Create a dictionary
'@param args - The key/value pairs to add to the dictionary
'@returns - The dictionary
'@remark - This is a helper function to make creating dictionaries easier and will help with future mac compatibility.
Private Function CreateDictionary(ParamArray args()) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 0 To UBound(args) Step 2
        dict.Add args(i), args(i + 1)
    Next
    Set CreateDictionary = dict
End Function

'Sort a collection on a given key
'@param col - The collection to sort
'@param key - The key to sort on
Private Sub sortCollectionOn(ByRef ast As Collection, ByVal key As String)
    'Convert to array for fast sort
    Dim items() As Object: ReDim items(1 To ast.count)
    Dim i As Long
    For Each Item In ast
        i = i + 1
        Set items(i) = Item
    Next

    'Sort array
    For i = 1 To UBound(items)
        Dim j As Long
        For j = i + 1 To UBound(items)
            If items(i)(key) > items(j)(key) Then
                Dim temp As Object: Set temp = items(i)
                Set items(i) = items(j)
                Set items(j) = temp
            End If
        Next
    Next

    'Convert back to collection
    Set ast = New Collection
    For i = 1 To UBound(items)
        ast.Add items(i)
    Next
End Sub

'**********************************
'* GENERATING TEXT FROM AST
'**********************************

'Generate a string which complies with a object of type "regex"
'@param result - The string to append the generated string to
'@param astRegex - The regex object to generate a string for
'@returns - The generated string for the specified regex
'@remark - During parsing a regex is composed of branches, each branch is ORed together. As such when generating data which complies
'we need to select 1 branch at random and generate for that branch.
Private Function GenerateForRegex(ByRef result As String, ByVal astRegex As Object) As String
    Dim iStart As Long: iStart = Len(result)
    Dim branchCount As Long: branchCount = astRegex("value").count
    Dim selectBranch As Long: selectBranch = RandBetween(1, branchCount)
    Dim astBranch As Object: Set astBranch = astRegex("value")(selectBranch)
    Call GenerateForBranch(result, astBranch)
    GenerateForRegex = mid(result, iStart + 1)
End Function

'Generate a string which complies with a object of type "branch"
'@param result - The string to append the generated string to
'@param astBranch - The branch object to generate a string for
'@remark - During parsing a branch is composed of pieces, each piece following the last. All pieces in a branch are required
'to match for the branch to match. As such when generating data which complies we need to generate for all pieces in a branch.
Private Sub GenerateForBranch(ByRef result As String, ByVal astBranch As Object)
    Dim astPiece As Object
    For Each astPiece In astBranch("value")
        Call GenerateForPiece(result, astPiece)
    Next
End Sub

'Generate a string which complies with a object of type "piece"
'@param result - The string to append the generated string to
'@param astPiece - The piece object to generate a string for
'@remark - During parsing a piece is composed of an atom and a quantifier. The quantifier determines how many times the atom must match.
'As such when generating data which complies we need to generate for the atom the number of times specified by the quantifier.
Private Sub GenerateForPiece(ByRef result As String, ByVal astPiece As Object)
    Dim astAtom As Object: Set astAtom = astPiece("value")
    Dim bounds As QuantifierBounds: bounds = getQuantifierBounds(astPiece)
    Dim i As Long
    For i = 1 To RandBetween(bounds.LB, bounds.UB)
        Call GenerateForAtom(result, astAtom)
    Next
End Sub

'Generate a string which complies with a object of type "atom"
'@param result - The string to append the generated string to
'@param astAtom - The atom object to generate a string for
'@remark - During parsing an atom is composed of a type and a value. The type determines how the value is interpreted.
'As such when generating data which complies we need to generate for the atom based on the type.
Private Sub GenerateForAtom(ByRef result As String, ByVal astAtom As Object)
    Select Case astAtom("atomType")
        Case "literal"
            result = result & astAtom("value")
        Case "capturing-group", "non-capturing-group", "named-capturing-group"
            Call GenerateForRegex(result, astAtom("value"))
        Case "positive-lookahead", "negative-lookahead", "positive-lookbehind", "negative-lookbehind"
            'Do nothing - not necessarily correct but good enough for now
        Case "charsets"
            Call GenerateForCharsets(result, astAtom("value"))
        Case "anchor"
            'Do nothing
        Case "backreference-named"
            'TODO: Implement
        Case "backreference-numeric"
            'TODO: Implement
        Case "empty"
            'Do nothing
    End Select
End Sub

'Generate a string which complies with a object of type "charsets"
'@param result - The string to append the generated string to
'@param astCharsets - The charsets object to generate a string for
'@remark - During parsing a charsets is composed of a list of charset objects. Each charset object represents a character set.
'As such when generating data which complies we need to select a charset at random and generate for that charset.
Private Sub GenerateForCharsets(ByRef result As String, ByVal astCharsets As Collection)
    Dim charsetCount As Long: charsetCount = astCharsets.count
    Dim selectCharset As Long: selectCharset = RandBetween(1, charsetCount)
    Dim astCharset As Object: Set astCharset = astCharsets(selectCharset)
    Call GenerateForCharset(result, astCharset)
End Sub

'Generate a string which complies with a object of type "charset"
'@param result - The string to append the generated string to
'@param astCharset - The charset object to generate a string for
'@remark - During parsing a charset is composed of a list of charset-item objects. Each charset-item object represents a character set item.
'As such when generating data which complies we need to select a charset-item at random and generate for that charset-item.
Private Sub GenerateForCharset(ByRef result As String, ByVal astCharset As Object)
    'Check if charset is charset or a negation
    Select Case astCharset("type")
        Case "charset"
            'TODO: Stabilise such that characters have the same probability of being selected
            If astCharset("charMin") = astCharset("charMax") Then
                result = result & ChrW(astCharset("charMin"))
            Else
                result = result & ChrW(RandBetween(astCharset("charMin"), astCharset("charMax")))
            End If
    End Select
End Sub

'Get the bounds for a quantifier
'@param astPiece - The piece object to get the quantifier bounds for
'@return - The quantifier bounds
'@remark - During parsing a piece is composed of an atom and a quantifier. The quantifier determines how many times the atom must match.
Private Function getQuantifierBounds(ByVal astPiece As Object) As QuantifierBounds
    With getQuantifierBounds
        Select Case astPiece("quantifier")
            Case "zeroPlus"
                .LB = 0
                .UB = This.GenData.maxRepitionLength
            Case "onePlus"
                .LB = 1
                .UB = This.GenData.maxRepitionLength
            Case "zeroOrOne"
                .LB = 0
                .UB = 1
            Case "bound"
                .LB = astPiece("min")
                .UB = astPiece("max")
            Case "one"
                .LB = 1
                .UB = 1
        End Select
    End With
End Function

'Generate a random number between min and max
'@param min - The minimum value
'@param max - The maximum value
'@return - A random number between min and max
Private Function RandBetween(ByVal min As Long, ByVal max As Long) As Long
    RandBetween = Int((max - min + 1) * Rnd + min)
End Function
