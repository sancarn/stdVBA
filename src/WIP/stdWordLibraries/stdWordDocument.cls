VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdWordDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'A really handy function to return all ranges
Public Function FindAll(ByVal sPattern As String, ByVal rHaystack As Range, Optional ByVal MatchCase As Boolean, Optional ByVal MatchWholeWord As Boolean, Optional ByVal MatchWildcards As Boolean, Optional ByVal IgnoreSpace As Boolean) As Collection
    'Create copy of rHaystack (will be replaced by find method)
    Dim rng As Range
    Set rng = rHaystack
    
    'Create collection
    Dim col As Collection
    Set col = New Collection
    
    'Setup find object
    With rng.Find
        .MatchCase = MatchCase
        .MatchWholeWord = MatchWholeWord
        .MatchWildcards = MatchWildcards
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Forward = True
        .MatchPrefix = False
        .MatchSuffix = False
        .MatchPhrase = False
        .IgnoreSpace = True
        .IgnorePunct = False
        
        'Inject pattern
        .Text = sPattern

        'Find all and add to collection
        While .Execute()
            col.Add rng
        Wend
    End With
    
    'Return collection
    Set FindAll = col
End Function

'A really handy function to return all ranges
Public Function FindAllRegex(ByVal sPattern As String, ByVal rHaystack As Range, Optional ByVal MatchCase As Boolean, Optional ByVal MatchWholeWord As Boolean, Optional ByVal MatchWildcards As Boolean, Optional ByVal IgnoreSpace As Boolean) As Collection
    'unimplemented, would be nice to have.
End Function