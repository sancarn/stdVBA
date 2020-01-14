VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LambdaTokenizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private initialized As Boolean
Private tokenizer As stdArray
Private Declare Sub VariantCopy Lib "oleaut32.dll" (ByRef pvargDest As Variant, ByRef pvargSrc As Variant)

'Creates a tokenizer object
'  @param oTokenizer - Dictionary<string,string>
'Example:
'  oTokenizer("SYMBOL") = "RegexToFind"
'  oTokenizer("ADD") = "\+"
Function Create(oTokenizer As Object) As LambdaCNF
    Dim o As LambdaTokenizer
    Set o = New LambdaTokenizer
    Call o.Init(oTokenizer)
    Set Create = o
End Function
Sub Init(ByVal oTokenizer As Object)
    If Not initialized Then
        'Create new array of rules
        Set tokenizer = stdArray.Create()
        
        Dim vSymbol As Variant
        For Each vSymbol In oTokenizer.keys()
            'Get pattern
            Dim pattern As String
            pattern = oTokenizer(vSymbol)
            
            'Create regex
            Dim regex As Object
            Set regex = CreateObject("VBScript.RegExp")
            regex.pattern = "^" & pattern
            
            'Push symbol and pattern regex
            tokenizer.push stdArray.Create(vSymbol, regex)
        Next
        initialized = True
    Else
        Err.Raise 1, "LambdaTokenizer::Init", "Object already initialized"
    End If
End Sub



' Tokenize the given input
' @param sInput
Public Function Tokenize(ByVal sInput As String) As stdArray
    Dim tokens As stdArray: Set tokens = stdArray.Create()
    Dim index As Long: index = 1
    Dim oldIndex As Long: oldIndex = 1
    
    Do While True
        'Try all matchers
        Dim i As Long
        For i = 1 To tokenizer.Length 'note: js uses 0 to tokenizer.length so might have to use i-1 somewhere
            'Get data for index
            Dim symbol As String, regex As Object, arr As stdArray
            Set arr = tokenizer.item(i)
            Set regex = arr.item(2)
            
            Dim oMatch As Object
            Set oMatch = regex.Execute(sInput)
            
            If oMatch.Count <> 0 Then
                Dim sMatchValue As String
                sMatchValue = oMatch(0).value
                
                If Len(sMatchValue) > 0 Then
                    'Create token structure
                    Dim o As Object
                    Set o = CreateObject("scripting.dictionary")
                    o("symbol") = arr.item(1)
                    o("text") = sMatchValue
                    Set o("range") = CreateObject("scripting.dictionary")
                    o("range")("start") = index
                    o("range")("end") = index + Len(sMatchValue)
                    
                    'Add token to tokens
                    Call tokens.push(o)
                    
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
    Set Tokenize = tokens
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
    Set x = r.Tokenize("1+2.0")
    
    Dim i As Long, ow As Object
    For i = 1 To x.Length
        Set ow = x.item(i)
        Debug.Print ow("symbol") & " ==> " & ow("text")
    Next
End Sub

