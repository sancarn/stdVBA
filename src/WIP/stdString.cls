VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "STD_Types_String"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Declare Function SafeArrayRedim Lib "oleaut32" (ByVal saPtr As Long, saBound As Long) As Long

Public Function Split(Expression As String, Optional Delimiter As String = " ", Optional ByVal Limit As Long = -1) as String()
    ' general variables that we need
    Dim lngA As Long, lngCount As Long, lngDelLen As Long, lngExpLen As Long, lngResults() As Long
    ' some dummy variables that we happen to need
    Dim Compare As VbCompareMethod, SafeArrayBound(1) As Long
    ' length information
    lngExpLen = LenB(Expression)
    lngDelLen = LenB(Delimiter)
    ' validate lengths and limit (limit must be larger than 0 or it must be unlimited)
    If lngExpLen > 0 And lngDelLen > 0 And (Limit > 0 Or Limit = -1&) Then
        ' now look up for the first position
        lngA = InStrB(1, Expression, Delimiter, Compare)
        ' InStrB is very fast, but it may give "between characters" results
        Do Until (lngA And 1) Or (lngA = 0)
            ' this is why we look for odd positions (1, 3, 5, 7 etc. are a valid position)
            lngA = InStrB(lngA + 1, Expression, Delimiter, Compare)
        Loop
        ' unlimited or limited?
        If Limit = -1& Then
            ' unlimited, reserve space for maximum possible amount of returned items
            ReDim lngResults(0 To (lngExpLen \ lngDelLen))
            ' index positions until none is found
            Do While lngA > 0
                ' remember this position
                lngResults(lngCount) = lngA
                ' look for the next one
                lngA = InStrB(lngA + lngDelLen, Expression, Delimiter, Compare)
                Do Until (lngA And 1) Or (lngA = 0)
                    lngA = InStrB(lngA + 1, Expression, Delimiter, Compare)
                Loop
                ' increase found counter
                lngCount = lngCount + 1
            Loop
        Else
            ' limited, reserve space for the wanted amount of items
            ReDim lngResults(0 To Limit - 1)
            ' index positions until none is found or until counter hits limitation
            Do While lngA > 0 And lngCount < Limit
                ' remember this position
                lngResults(lngCount) = lngA
                ' look for the next one
                lngA = InStrB(lngA + lngDelLen, Expression, Delimiter, Compare)
                Do Until (lngA And 1) Or (lngA = 0)
                    lngA = InStrB(lngA + 1, Expression, Delimiter, Compare)
                Loop
                ' increase found counter
                lngCount = lngCount + 1
            Loop
        End If
        ' set results to actual findings
        ReDim Preserve Split(0 To lngCount)
        ' see if we found any results
        If lngCount = 0 Then
            ' nope, just set the only item to be the whole string
            Split(0) = Expression
        Else
            ' get the first item
            Split(0) = LeftB$(Expression, lngResults(0) - 1)
            ' get the other items except the last one
            For lngCount = 0 To lngCount - 2
                Split(lngCount + 1) = MidB$(Expression, lngResults(lngCount) + lngDelLen, lngResults(lngCount + 1) - lngResults(lngCount) - lngDelLen)
            Next lngCount
            ' get the last item
            Split(lngCount + 1) = RightB$(Expression, lngExpLen - lngResults(lngCount) - lngDelLen + 1)
        End If
    Else
        ' clean any possible data that exists in the passed string array (like if it is multidimensional)
        If Not Not Split Then Erase Split
        ' mysterious IDE error fix
        Debug.Assert App.hInstance
        ' reset to one element, one dimension
        ReDim Split(0 To 0)
        ' custom redimension: remove the items (this duplicates the VB6 Split behavior)
        SafeArrayRedim Not Not Split, SafeArrayBound(0)
    End If
End Function