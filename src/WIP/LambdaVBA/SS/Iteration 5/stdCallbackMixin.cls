'Token for Lambda
Private Type TToken
  pType As String
  pValue As String
End Type






Private Function TokeniseLambda(ByRef Tokens() As TToken)
  Dim Lambda() As Byte
  Lambda = pCallback.lambda
  
  ReDim Tokens(0 To UBound(Lambda) - LBound(Lambda))
  
  Dim TokenHash(255) As Boolean, i As Integer
  For i = 32 To 47: TokenHash(i) = True: Next
  For i = 58 To 64: TokenHash(i) = True: Next
  For i = 91 To 96: TokenHash(i) = True: Next
  For i = 123 To 126: TokenHash(i) = True: Next
  
  Dim s As String, sToken As String, ch As Integer
  Dim token As TToken
  For i = 0 To UBound(Lambda) Step 2
    ch = Lambda(i + 1)
    If ch = 10 Then ch = 58 'Assume vbLf = ":"
    'Do not pass any disallow all control characters
    'TODO: if ch < 32 then Continue For
    If TokenHash(ch) Then
      'Previous token
      token.pType = "Literal"
      token.pValue = s
      TokeniseLambda.add token
      
      'Current Token
      token.pType = "Symbol"
      token.pValue = Asc(ch)
      TokeniseLambda.add token
      
      'Next token
      s = ""
    Else
      'Concatenate string literal
      s = s & Asc(ch)
    End If
  Next
  
  'Add final token
  If s <> "" Then
    token.pType = "Literal"
    token.pValue = s
    TokeniseLambda.add token
  End If
End Function


