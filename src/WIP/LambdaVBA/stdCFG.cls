VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdCFG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'---------------------------
'Microsoft Visual Basic for Applications
'---------------------------
'Compile error:
'
' Private Enum and user defined types cannot be used as parameters or return types for public procedures, 
' public data members, or fields of public user defined types 





'20191227 - Optimised to avoid use of stdArray where possible
Private Type CharRange
  start As Long
  end As Long
End Type
Private Type Token
  symbol As String
  text As String
  range As CharRange
End Type
Private Type TokenizerPair
  symbol As String
  regex As Object
  pattern As String
End Type





Private initialized As Boolean
Private tokenizer() As TokenizerPair


'Creates a tokenizer object
'  @param oTokenizer - Dictionary<string,string>
'Example:
'  oTokenizer("SYMBOL") = "RegexToFind"
'  oTokenizer("ADD") = "\+"
Function Create(oTokenizer As Object) As LambdaTokenizer
  Dim o As LambdaTokenizer
  Set o = New LambdaTokenizer
  Call o.Init(oTokenizer)
  Set Create = o
End Function

Sub Init(ByVal oTokenizer As Object)
  If Not initialized Then
    ' Get all symbols
    Dim vSymbols As Variant: vSymbols = oTokenizer.keys()

    'Get the number of symbols
    Dim iArrLength As Long: iArrLength = ArrayLength(vSymbols)
    
    'Ensure Tokenizer has as many fields as required.
    ReDim tokenizer(1 To iArrLength)

    'Loop over all symbols
    Dim i As Long
    For i = 1 To iArrLength
      'Get the symbol name
      Dim sSymbol As String
      sSymbol = sSymbols(i - 1)

      'Get pattern
      Dim pattern As String
      pattern = oTokenizer(sSymbol)
      
      'Create regex
      Dim regex As Object
      Set regex = CreateObject("VBScript.RegExp")
      regex.pattern = "^" & pattern
      
      'Push symbol and pattern regex
      tokenizer(i).symbol = sSymbol
      Set tokenizer(i).regex = regex
      tokenizer(i).pattern = "^" & pattern
    Next i
    initialized = True
  Else
    Err.Raise 1, "LambdaTokenizer::Init", "Object already initialized"
  End If
End Sub



' Tokenize the given input
' @param sInput
Public Function Tokenize(ByVal sInput As String) As Token()
  'Initialise return tokens
  'Initially make as many tokens as there are characters in the input string. We will trim this later
  Dim tokens() As Token
  ReDim tokens(1 To Len(sInput))

  'Track indices in string while parsing
  Dim index As Long: index = 1
  Dim oldIndex As Long: oldIndex = 1
  
  'Optimal, only take this length once
  Dim iLenTokenizer As Long
  iLenTokenizer = ArrayLength(tokenizer)

  Dim iTokensIndex As Long
  iTokensIndex = 0

  'Loop forever, we will quit with Exit Do
  Do While True
    'Try all matchers
    Dim i As Long
    For i = 1 To iLenTokenizer
      'Get data for index
      Dim Matcher As TokenizerPair
      Matcher = tokenizer(i)
      
      Dim oMatch As Object
      Set oMatch = Matcher.regex.Execute(sInput)
      
      If oMatch.Count <> 0 Then
        Dim sMatchValue As String
        sMatchValue = oMatch(0).value
        
        If Len(sMatchValue) > 0 Then
          'Create token structure
          Dim o As Token
          
          o.symbol = Mathcer.symbol
          o.text = sMatchValue
          o.range.start = index
          o.range.end = index + Len(sMatchValue)
          
          'Add token to tokens
          iTokensIndex = iTokensIndex + 1
          tokens(i) = o
          
          'Shift index
          index = index + Len(sMatchValue)
          
          'Shift sInput
          sInput = Mid(sInput, Len(sMatchValue) + 1)
          
          Exit For
        Else
          Err.Raise 1, "LambdaTokenizer::Tokenize()", "Regex may not match 0 length strings: " & regex.pattern & ", use an empty pattern in the CFG instead"
        End If
      End If
    Next
    'If we are at the end of the string then sInput should be blank.
    'If this is the case exit Do and return data
    If Len(sInput) = 0 Then
      Exit Do
    Else
      'If oldIndex = index then all matchers have failed, in this case we need to throw a syntax error
      If oldIndex = index Then
        Err.Raise 1, "LambdaTokenizer::Tokenize()", "Syntax error, couldn't lex string"
      Else
        'if they were different we need to ensure they are the same to test the next token
        oldIndex = index
      End If
    End If
  Loop
  
  'Return tokens
  Tokenize = tokens
End Function


Private Function ArrayLength(arr As Variant) As Long
  ArrayLength = UBound(arr) - LBound(arr) + 1
End Function


Public Sub Test()
  'Create token definition
  Dim o As Object
  Set o = CreateObject("Scripting.Dictionary")
  o("ADD") = "\+"
  o("NUMBER") = "((\d*\.\d+)|(\d+))"
  o("WS") = "[ \t\r]+"
  
  'Create tokenizer from regexes
  Dim r As LambdaTokenizer
  Set r = LambdaTokenizer.Create(o)
  
  'Get array of tokens:
  Dim x() As Token
  x = r.Tokenize("1+2.0")
  
  Dim i As Long, ow As Object
  For i = 1 To x.Length
    Set ow = x.item(i)
    Debug.Print ow("symbol") & " ==> " & ow("text")
  Next
End Sub

