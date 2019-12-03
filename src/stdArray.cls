VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@TODO:
'* Implement Exceptions throughout all Array functions.
'* Fully implement pInitialised where necessary.
'* Build Methods Slice; Splice; Sort
'* Add methods from ruby
'* Documentation of methods
'* Optimisation: E.G. Variant arrays are slow

#If VBA6 Then
  Private Declare PtrSafe Sub VariantCopy Lib "oleaut32.dll" (ByRef pvargDest As Variant, ByRef pvargSrc As Variant)
 
#Else
  Private Declare Sub VariantCopy Lib "oleaut32.dll" (ByRef pvargDest As Variant, ByRef pvargSrc As Variant)
#End If

Public Enum SortDirection
    Ascending = 1
    Descending = 2
End Enum
Private Enum AllocationEnum
    isNotCalculated = 0
    isTrue = 1
    isFalse = 2
End Enum

Private pAlloc As AllocationEnum
Private pArr As Variant
Private pInitialised As Boolean

Public Function Create(ParamArray params() As Variant) As stdArray
    Set Create = New stdArray
    Dim v() As Variant
    If UBound(params) = -1 And LBound(params) = 0 Then
        Create.arr = v
    Else
      ReDim v(0 To UBound(params))
      v = params
      Create.arr = v
    End If
    
    Create.Init
End Function

Public Function FromArray(ByVal arr As Variant) As stdArray
    Set FromArray = New stdArray
    FromArray.arr = arr
    FromArray.Init
End Function

Property Get arr() As Variant
  arr = pArr
End Property

Property Let arr(val As Variant)
  If Not pInitialised Then
    pArr = val
    
  Else
    Err.Raise 0, "STD.Types.Array::letArr()", "Cannot set array of already initialised array."
  End If
End Property

Friend Sub Init()
  If Not pInitialised Then
    pInitialised = True
  End If
End Sub

Public Sub push(ByVal el As Variant)
  If pInitialised Then
    If pAlloc = isNotCalculated Then pAlloc = AllocationCheck(pArr)
    If pAlloc = isTrue Then
      ReDim Preserve pArr(UBound(pArr) + 1)
      VariantCopy pArr(UBound(pArr)), el
    Else
      ReDim pArr(0)
      VariantCopy pArr(0), el
      pAlloc = isTrue
    End If
  Else
    'Error
  End If
End Sub

Public Function pop() As Variant
  If pInitialised Then
    If IsArrayAllocated(pArr) Then
      pop = pArr(UBound(pArr))
      ReDim Preserve pArr(UBound(pArr) - 1)
    End If
  Else
    'Error
  End If
End Function

Public Function Shift() As Variant
  If pInitialised Then
    If IsArrayAllocated(pArr) Then
      Shift = pArr(LBound(pArr))
      ReDim Preserve pArr(LBound(pArr) + 1, UBound(pArr)) 'This may not work?
    End If
  Else
    'Error
  End If
End Function
Public Function unshift(val As Variant) As stdArray
  If pInitialised Then
    If IsArrayAllocated(pArr) Then
      ReDim Preserve pArr(LBound(pArr) - 1, UBound(pArr)) 'This may not work?
      pArr(LBound(pArr)) = val
    End If
  Else
    'Error
  End If
  Set unshift = Me
End Function

Public Function Slice() As stdArray

End Function

Public Function Splice() As stdArray

End Function

Public Function Concat(arr As stdArray) As stdArray
  Dim x As stdArray
  Set x = Me.Clone()
  
  If Not arr Is Nothing Then
    Dim i As Long
    For i = 1 To arr.Length
      Call x.push(arr.item(i))
    Next
  End If
  
  Set Concat = x
End Function

Public Function Join(Optional ByVal delimeter As String = ",") As String
  If pInitialised Then
    If IsArrayAllocated(pArr) Then
      If Length() > 0 Then
        Dim sOutput As String
        sOutput = pArr(LBound(pArr))
        
        Dim i As Long
        For i = LBound(pArr) + 1 To UBound(pArr)
          sOutput = sOutput & delimeter & pArr(i)
        Next
        Join = sOutput
      Else
        Join = ""
      End If
    End If
  Else
    'Error
  End If
End Function

Public Function Clone() As stdArray
  If pInitialised Then
    Set Clone = New stdArray
    Clone.arr = pArr
    Call Clone.Init
  Else
    'Error
  End If
End Function


'Iterative Functions (All require stdCallback):
Public Function every(ByVal cb As stdCallback, Optional ByVal metadata As Object = Nothing) As Boolean
  If pInitialised Then
    If IsArrayAllocated(pArr) Then
      Dim i As Integer
      For i = LBound(pArr) To UBound(pArr)
        If Not IIf(metadata Is Nothing, cb(pArr(i)), cb(metadata, pArr(i))) Then
          every = False
          Exit Function
        End If
      Next
      evert = True
    Else
      every = False
    End If
  Else
    'Error
  End If
End Function

Public Function some(ByVal cb As stdCallback, Optional ByVal metadata As Object = Nothing) As Boolean
  If pInitialised Then
    If IsArrayAllocated(pArr) Then
      Dim i As Integer, test As Boolean
      For i = LBound(pArr) To UBound(pArr)
        If IIf(metadata Is Nothing, cb(pArr(i)), cb(metadata, pArr(i))) Then
          some = True
          Exit Function
        End If
      Next
      some = False
    Else
      some = False
    End If
  Else
    'Error
  End If
End Function

Public Sub forEach(ByVal cb As stdCallback, Optional ByVal metadata As Object = Nothing)
  If pInitialised Then
    If IsArrayAllocated(pArr) Then
      Dim i As Integer
      For i = LBound(pArr) To UBound(pArr)
        If metadata Is Nothing Then
            'TODO:
            'cb pArr(i)
        Else
            'cb metadata, pArr(i)
        End If
      Next
    End If
  Else
    'Error
  End If
End Sub

Public Function Map(ByVal cb As stdCallback, Optional ByVal metadata As Object = Nothing) As stdArray
  If pInitialised Then
    Dim pMap As stdArray
    Set pMap = Me.Clone()
    If IsArrayAllocated(pArr) Then
      Dim i As Integer
      For i = 1 To Length
        'BUGFIX: Sometimes required, not sure when
        Dim v As Variant
        If IsObject(item(i)) Then
            Set v = item(i)
        Else
            v = item(i)
        End If
        
        'Call callback based on metadata
        If metadata Is Nothing Then
            pMap.item(i) = cb(v)
        Else
            pMap.item(i) = cb.Run(metadata, v)
        End If
      Next
    End If
    Set Map = pMap
  Else
    'Error
  End If
End Function

'OPTIMISE: Needs optimisation. Currently very sub-optimal
Public Function Unique(Optional ByVal cb As stdCallback = Nothing, Optional ByVal metadata As Variant) As stdArray
  Dim ret As stdArray
  Set ret = stdArray.Create()
  
  Dim o As Object
  Set o = CreateObject("Scripting.Dictionary")
  Dim vKey As Variant
  
  'If cb is nothing then we can optimise this a bit further...
  If cb Is Nothing Then
    For i = 1 To Length
      VariantCopy vKey, Me.item(i)
      If IsEmpty(o(vKey)) Then
        VariantCopy o(vKey), Me.item(i)
      End If
    Next
  Else
    For i = 1 To Length
      'Filter based on key whose content's comes from callback
      If metadata Is Nothing Then
        VariantCopy vKey, cb(Me.item(i))
      Else
        VariantCopy vKey, cb(metadata, Me.item(i))
      End If
      
      'Get unique
      If IsEmpty(o(vKey)) Then
        VariantCopy o(vKey), Me.item(i)
      End If
    Next
  End If
  
  'Loop return objects and return
  For Each vKey In o.keys()
    ret.push vKey
  Next
  
  'Return data
  Set Unique = ret
End Function

Public Function Reduce(ByVal cb As stdCallback, Optional ByVal initialValue As Variant) As Variant
  If pInitialised Then
    If IsArrayAllocated(pArr) Then
      Reduce = initialValue
      Dim i As Integer
      For i = LBound(pArr) To UBound(pArr)
        'BUGFIX: Sometimes required, not sure when
        Dim el As Variant
        If IsObject(pArr(i)) Then
            Set el = pArr(i)
        Else
            el = pArr(i)
        End If
        
        'Reduce
        Reduce = cb(Reduce, el)
      Next
    Else
      Reduce = initialValue
    End If
  Else
    'Error
  End If
End Function

Public Function Filter(ByVal cb As stdCallback, Optional ByVal metadata As Object = Nothing) As stdArray
    Dim ret As stdArray
    Set ret = stdArray.Create()
    Set Filter = ret
    
    'If initialised...
    If pInitialised Then
        Dim i As Long, v As Variant
        'Loop over array
        For i = 1 To Me.Length
            'If callback succeeds (include metadata if given)
            If metadata Is Nothing Then
                If cb(Me.item(i)) Then
                    Call ret.push(Me.item(i))
                End If
            Else
                If cb(metadata, Me.item(i)) Then
                    Call ret.push(Me.item(i))
                End If
            End If
        Next i
    Else
        'error
    End If
End Function

Public Function Count(Optional ByVal v As Variant, Optional ByVal metadata As Object = Nothing) As Long
    Dim i As Long, lCount As Long
    lCount = 0
    'Count by stdCallback condition
    If TypeName(v) = "stdCallback" Then
        For i = 1 To Me.Length
            If IIf(metadata Is Nothing, v(Me.item(i)), v(metadata, Me.item(i))) Then
                lCount = lCount + 1
            End If
        Next i
    
    'Count value in an array
    ElseIf Not IsEmpty(v) Then
        For i = 1 To Me.Length
            If IsObject(Me.item(i)) Then
                If Me.item(i) Is v Then
                    lCount = lCount + 1
                End If
            Else
                If Me.item(i) = v Then
                    lCount = lCount + 1
                End If
            End If
        Next i
    
    'Return count of all items
    Else
        lCount = Me.Length
    End If
    
    'Return count
    Count = lCount
End Function

Public Function groupBy(ByVal cb As stdCallback, Optional ByVal metadata As Object) As Object
    'Array to store result in
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    'Loop over items
    Dim i As Long
    For i = 1 To Me.Length
        'Get grouping key
        Dim key As Variant
        If metadata Is Nothing Then
          key = cb(Me.item(i))
        Else
          key = cb(metadata, Me.item(i))
        End If
        
        'If key is not set then set it
        If IsEmpty(result(key)) Then Set result(key) = stdArray.Create()
        
        'Push item to key
        result(key).push Me.item(i)
    Next
    
    'Return result
    Set groupBy = result
End Function

Public Function Sort(ByVal cb As stdCallback) As stdArray
  MsgBox "NOT IMPLEMENTED"
  Debug.Assert False
End Function




Public Sub SortInPlace(Optional ByVal direction As SortDirection = Ascending, Optional ByVal compareMode As VbCompareMethod = vbTextCompare, Optional ByVal NoAlerts As Boolean = False, Optional ByRef ErrorLevel As Boolean)
  ErrorLevel = QSortInPlace(pArr, -1, -1, direction = Descending, compareMethod, NoAlerts)
End Sub

Public Sub Resize(Optional ByVal LowerBound As Integer = -1, Optional ByVal UpperBound As Integer = -1)
    If LowerBound = -1 Then LowerBound = LBound(pArr)
    If UpperBound = -1 Then UpperBound = UBound(pArr)
    ReDim Preserve pArr(LowerBound To UpperBound)
End Sub

Public Sub Reverse()
    ReverseArrayInPlace2 pArr
End Sub

Public Function indexOf(ByVal el As Variant, Optional ByVal start As Integer)
  If IsArrayAllocated(pArr) Then
    If Not start Then start = LBound(pArr)
    If start < 0 Then start = UBound(pArr) - start + 1
    
    'find the index of an item in the array. If array doesn't contain item, return -1
    Dim i As Integer
    For i = start To UBound(pArr)
      If pArr(i) = el Then
        indexOf = i
        Exit Function
      End If
    Next
  End If
  
  indexOf = -1
End Function

Public Function lastIndexOf(ByVal el As Variant, Optional ByVal start As Integer)
  If Not start Then start = UBound(pArr)
  If start < 0 Then start = UBound(pArr) - start + 1
  
  'find the index of an item in the array. If array doesn't contain item, return -1
  Dim i As Integer
  For i = start To LBound(pArr) Step -1
    If pArr(i) = el Then
      lastIndexOf = i
      Exit Function
    End If
  Next
  lastIndexOf = -1
End Function

Public Function includes(ByVal el As Variant, Optional ByVal startFrom As Integer = -1) As Boolean
  includes = indexOf(el) > startFrom
End Function

Public Function Length() As Long
    'get the total number of elements in the array
    If pInitialised Then
      On Error Resume Next
        Length = UBound(pArr) - LBound(pArr) + 1
      On Error GoTo 0
    Else
      'STD.Error.Raise(...)
    End If
End Function

Public Property Get item(ByVal i As Integer) As Variant
    'item(1) = 1st element
    'item(2) = 2nd element
    'etc.
    If IsObject(pArr(LBound(pArr) + i - 1)) Then
      Set item = pArr(LBound(pArr) + i - 1)
    Else
      item = pArr(LBound(pArr) + i - 1)
    End If
End Property
Public Property Set item(ByVal i As Integer, v As Object)
  Set pArr(LBound(pArr) + i - 1) = v
End Property
Public Property Let item(ByVal i As Integer, v As Variant)
  pArr(LBound(pArr) + i - 1) = v
End Property


Public Sub test()
  On Error Resume Next
  
  Debug.Print "1-Generic tests:"
  Dim v As stdArray:
  Set v = stdArray.Create(1)
  Debug.Print "1-1:" & (v.item(0) = 1)
  Call v.push(2)
  Debug.Print "1-2:" & (v.item(0) = 1)
  Debug.Print "1-3:" & (v.item(1) = 2)
  
  'Test object pushing and access
  Dim o As Object
  Set o = CreateObject("Scripting.Dictionary")
  o("hello") = "world"
  Call v.push(o)
  Debug.Print "1-4:" & (v.item(2)("hello") = "world")
  
  Debug.Print "2-Length tests:"
  Debug.Print "2-1:" & (stdArray.Create().Length = 0)
  Debug.Print "2-2:" & (stdArray.Create(1).Length = 1)
  Debug.Print "2-3:" & (stdArray.Create(1, 2).Length = 2)
  
  
  If Not IsEmpty(stdCallback) Then
    Debug.Print "9-stdCallback tests"
    
  End If
End Sub










'##########################################################################################################
'##                                     SORT IN PLACE - CHIP PEARSON                                     ##
'##########################################################################################################

Private Function QSortInPlace( _
    ByRef InputArray As Variant, _
    Optional ByVal lb As Long = -1&, _
    Optional ByVal ub As Long = -1&, _
    Optional ByVal Descending As Boolean = False, _
    Optional ByVal compareMode As VbCompareMethod = vbTextCompare, _
    Optional ByVal NoAlerts As Boolean = False) As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' QSortInPlace
    '
    ' This function sorts the array InputArray in place -- this is, the original array in the
    ' calling procedure is sorted. It will work with either string data or numeric data.
    ' It need not sort the entire array. You can sort only part of the array by setting the LB and
    ' UB parameters to the first (LB) and last (UB) element indexes that you want to sort.
    ' LB and UB are optional parameters. If omitted LB is set to the LBound of InputArray, and if
    ' omitted UB is set to the UBound of the InputArray. If you want to sort the entire array,
    ' omit the LB and UB parameters, or set both to -1, or set LB = LBound(InputArray) and set
    ' UB to UBound(InputArray).
    '
    ' By default, the sort method is case INSENSTIVE (case doens't matter: "A", "b", "C", "d").
    ' To make it case SENSITIVE (case matters: "A" "C" "b" "d"), set the CompareMode argument
    ' to vbBinaryCompare (=0). If Compare mode is omitted or is any value other than vbBinaryCompare,
    ' it is assumed to be vbTextCompare and the sorting is done case INSENSITIVE.
    '
    ' The function returns TRUE if the array was successfully sorted or FALSE if an error
    ' occurred. If an error occurs (e.g., LB > UB), a message box indicating the error is
    ' displayed. To suppress message boxes, set the NoAlerts parameter to TRUE.
    '
    ''''''''''''''''''''''''''''''''''''''
    ' MODIFYING THIS CODE:
    ''''''''''''''''''''''''''''''''''''''
    ' If you modify this code and you call "Exit Procedure", you MUST decrment the RecursionLevel
    ' variable. E.g.,
    '       If SomethingThatCausesAnExit Then
    '           RecursionLevel = RecursionLevel - 1
    '           Exit Function
    '       End If
    '''''''''''''''''''''''''''''''''''''''
    '
    ' Note: If you coerce InputArray to a ByVal argument, QSortInPlace will not be
    ' able to reference the InputArray in the calling procedure and the array will
    ' not be sorted.
    '
    ' This function uses the following procedures. These are declared as Private procedures
    ' at the end of this module:
    '       IsArrayAllocated
    '       IsSimpleDataType
    '       IsSimpleNumericType
    '       QSortCompare
    '       NumberOfArrayDimensions
    '       ReverseArrayInPlace
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim Temp As Variant
    Dim Buffer As Variant
    Dim CurLow As Long
    Dim CurHigh As Long
    Dim CurMidpoint As Long
    Dim Ndx As Long
    Dim pCompareMode As VbCompareMethod
    
    '''''''''''''''''''''''''
    ' Set the default result.
    '''''''''''''''''''''''''
    QSortInPlace = False
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' This variable is used to determine the level
    ' of recursion  (the function calling itself).
    ' RecursionLevel is incremented when this procedure
    ' is called, either initially by a calling procedure
    ' or recursively by itself. The variable is decremented
    ' when the procedure exits. We do the input parameter
    ' validation only when RecursionLevel is 1 (when
    ' the function is called by another function, not
    ' when it is called recursively).
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Static RecursionLevel As Long
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Keep track of the recursion level -- that is, how many
    ' times the procedure has called itself.
    ' Carry out the validation routines only when this
    ' procedure is first called. Don't run the
    ' validations on a recursive call to the
    ' procedure.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    RecursionLevel = RecursionLevel + 1
    
    If RecursionLevel = 1 Then
        ''''''''''''''''''''''''''''''''''
        ' Ensure InputArray is an array.
        ''''''''''''''''''''''''''''''''''
        If IsArray(InputArray) = False Then
            If NoAlerts = False Then
                MsgBox "The InputArray parameter is not an array."
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' InputArray is not an array. Exit with a False result.
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''
            RecursionLevel = RecursionLevel - 1
            Exit Function
        End If
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Test LB and UB. If < 0 then set to LBound and UBound
        ' of the InputArray.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If lb < 0 Then
            lb = LBound(InputArray)
        End If
        If ub < 0 Then
            ub = UBound(InputArray)
        End If
        
        Select Case NumberOfArrayDimensions(InputArray)
            Case 0
                ''''''''''''''''''''''''''''''''''''''''''
                ' Zero dimensions indicates an unallocated
                ' dynamic array.
                ''''''''''''''''''''''''''''''''''''''''''
                If NoAlerts = False Then
                    MsgBox "The InputArray is an empty, unallocated array."
                End If
                RecursionLevel = RecursionLevel - 1
                Exit Function
            Case 1
                ''''''''''''''''''''''''''''''''''''''''''
                ' We sort ONLY single dimensional arrays.
                ''''''''''''''''''''''''''''''''''''''''''
            Case Else
                ''''''''''''''''''''''''''''''''''''''''''
                ' We sort ONLY single dimensional arrays.
                ''''''''''''''''''''''''''''''''''''''''''
                If NoAlerts = False Then
                    MsgBox "The InputArray is multi-dimensional." & _
                          "QSortInPlace works only on single-dimensional arrays."
                End If
                RecursionLevel = RecursionLevel - 1
                Exit Function
        End Select
        '''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Ensure that InputArray is an array of simple data
        ' types, not other arrays or objects. This tests
        ' the data type of only the first element of
        ' InputArray. If InputArray is an array of Variants,
        ' subsequent data types may not be simple data types
        ' (e.g., they may be objects or other arrays), and
        ' this may cause QSortInPlace to fail on the StrComp
        ' operation.
        '''''''''''''''''''''''''''''''''''''''''''''''''''
        If IsSimpleDataType(InputArray(LBound(InputArray))) = False Then
            If NoAlerts = False Then
                MsgBox "InputArray is not an array of simple data types."
                RecursionLevel = RecursionLevel - 1
                Exit Function
            End If
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' ensure that the LB parameter is valid.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        Select Case lb
            Case Is < LBound(InputArray)
                If NoAlerts = False Then
                    MsgBox "The LB lower bound parameter is less than the LBound of the InputArray"
                End If
                RecursionLevel = RecursionLevel - 1
                Exit Function
            Case Is > UBound(InputArray)
                If NoAlerts = False Then
                    MsgBox "The LB lower bound parameter is greater than the UBound of the InputArray"
                End If
                RecursionLevel = RecursionLevel - 1
                Exit Function
            Case Is > ub
                If NoAlerts = False Then
                    MsgBox "The LB lower bound parameter is greater than the UB upper bound parameter."
                End If
                RecursionLevel = RecursionLevel - 1
                Exit Function
        End Select
    
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' ensure the UB parameter is valid.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        Select Case ub
            Case Is > UBound(InputArray)
                If NoAlerts = False Then
                    MsgBox "The UB upper bound parameter is greater than the upper bound of the InputArray."
                End If
                RecursionLevel = RecursionLevel - 1
                Exit Function
            Case Is < LBound(InputArray)
                If NoAlerts = False Then
                    MsgBox "The UB upper bound parameter is less than the lower bound of the InputArray."
                End If
                RecursionLevel = RecursionLevel - 1
                Exit Function
            Case Is < lb
                If NoAlerts = False Then
                    MsgBox "the UB upper bound parameter is less than the LB lower bound parameter."
                End If
                RecursionLevel = RecursionLevel - 1
                Exit Function
        End Select
    
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' if UB = LB, we have nothing to sort, so get out.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If ub = lb Then
            QSortInPlace = True
            RecursionLevel = RecursionLevel - 1
            Exit Function
        End If
    
    End If ' RecursionLevel = 1
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Ensure that CompareMode is either vbBinaryCompare  or
    ' vbTextCompare. If it is neither, default to vbTextCompare.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If (compareMode = vbBinaryCompare) Or (compareMode = vbTextCompare) Then
        pCompareMode = compareMode
    Else
        pCompareMode = vbTextCompare
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Begin the actual sorting process.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    CurLow = lb
    CurHigh = ub
    
    If lb = 0 Then
        CurMidpoint = ((lb + ub) \ 2) + 1
    Else
        CurMidpoint = (lb + ub) \ 2 ' note integer division (\) here
    End If
    Temp = InputArray(CurMidpoint)
    
    Do While (CurLow <= CurHigh)
        
        Do While QSortCompare(V1:=InputArray(CurLow), V2:=Temp, compareMode:=pCompareMode) < 0
            CurLow = CurLow + 1
            If CurLow = ub Then
                Exit Do
            End If
        Loop
        
        Do While QSortCompare(V1:=Temp, V2:=InputArray(CurHigh), compareMode:=pCompareMode) < 0
            CurHigh = CurHigh - 1
            If CurHigh = lb Then
               Exit Do
            End If
        Loop
    
        If (CurLow <= CurHigh) Then
            Buffer = InputArray(CurLow)
            InputArray(CurLow) = InputArray(CurHigh)
            InputArray(CurHigh) = Buffer
            CurLow = CurLow + 1
            CurHigh = CurHigh - 1
        End If
    Loop
    
    If lb < CurHigh Then
        QSortInPlace InputArray:=InputArray, lb:=lb, ub:=CurHigh, _
            Descending:=Descending, compareMode:=pCompareMode, NoAlerts:=True
    End If
    
    If CurLow < ub Then
        QSortInPlace InputArray:=InputArray, lb:=CurLow, ub:=ub, _
            Descending:=Descending, compareMode:=pCompareMode, NoAlerts:=True
    End If
    
    '''''''''''''''''''''''''''''''''''''
    ' If Descending is True, reverse the
    ' order of the array, but only if the
    ' recursion level is 1.
    '''''''''''''''''''''''''''''''''''''
    If Descending = True Then
        If RecursionLevel = 1 Then
            ReverseArrayInPlace2 InputArray, lb, ub
        End If
    End If
    
    RecursionLevel = RecursionLevel - 1
    QSortInPlace = True
End Function

Private Function QSortCompare(V1 As Variant, V2 As Variant, _
    Optional compareMode As VbCompareMethod = vbTextCompare) As Long
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' QSortCompare
    ' This function is used in QSortInPlace to compare two elements. If
    ' V1 AND V2 are both numeric data types (integer, long, single, double)
    ' they are converted to Doubles and compared. If V1 and V2 are BOTH strings
    ' that contain numeric data, they are converted to Doubles and compared.
    ' If either V1 or V2 is a string and does NOT contain numeric data, both
    ' V1 and V2 are converted to Strings and compared with StrComp.
    '
    ' The result is -1 if V1 < V2,
    '                0 if V1 = V2
    '                1 if V1 > V2
    ' For text comparisons, case sensitivity is controlled by CompareMode.
    ' If this is vbBinaryCompare, the result is case SENSITIVE. If this
    ' is omitted or any other value, the result is case INSENSITIVE.
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim d1 As Double
    Dim d2 As Double
    Dim S1 As String
    Dim S2 As String
    
    Dim Compare As VbCompareMethod
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ' Test CompareMode. Any value other than
    ' vbBinaryCompare will default to vbTextCompare.
    ''''''''''''''''''''''''''''''''''''''''''''''''
    If compareMode = vbBinaryCompare Or compareMode = vbTextCompare Then
        Compare = compareMode
    Else
        Compare = vbTextCompare
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''
    ' If either V1 or V2 is either an array or
    ' an Object, raise a error 13 - Type Mismatch.
    '''''''''''''''''''''''''''''''''''''''''''''''
    If IsArray(V1) = True Or IsArray(V2) = True Then
        Err.Raise 13
        Exit Function
    End If
    If IsObject(V1) = True Or IsObject(V2) = True Then
        Err.Raise 13
        Exit Function
    End If
    
    If IsSimpleNumericType(V1) = True Then
        If IsSimpleNumericType(V2) = True Then
            '''''''''''''''''''''''''''''''''''''
            ' If BOTH V1 and V2 are numeric data
            ' types, then convert to Doubles and
            ' do an arithmetic compare and
            ' return the result.
            '''''''''''''''''''''''''''''''''''''
            d1 = CDbl(V1)
            d2 = CDbl(V2)
            If d1 = d2 Then
                QSortCompare = 0
                Exit Function
            End If
            If d1 < d2 Then
                QSortCompare = -1
                Exit Function
            End If
            If d1 > d2 Then
                QSortCompare = 1
                Exit Function
            End If
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''
    ' Either V1 or V2 was not numeric data type.
    ' Test whether BOTH V1 AND V2 are numeric
    ' strings. If BOTH are numeric, convert to
    ' Doubles and do a arithmetic comparison.
    ''''''''''''''''''''''''''''''''''''''''''''
    If IsNumeric(V1) = True And IsNumeric(V2) = True Then
        d1 = CDbl(V1)
        d2 = CDbl(V2)
        If d1 = d2 Then
            QSortCompare = 0
            Exit Function
        End If
        If d1 < d2 Then
            QSortCompare = -1
            Exit Function
        End If
        If d1 > d2 Then
            QSortCompare = 1
            Exit Function
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''
    ' Either or both V1 and V2 was not numeric
    ' string. In this case, convert to Strings
    ' and use StrComp to compare.
    ''''''''''''''''''''''''''''''''''''''''''''''
    S1 = CStr(V1)
    S2 = CStr(V2)
    QSortCompare = StrComp(S1, S2, Compare)

End Function

Private Function NumberOfArrayDimensions(arr As Variant) As Integer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' NumberOfArrayDimensions
    ' This function returns the number of dimensions of an array. An unallocated dynamic array
    ' has 0 dimensions. This condition can also be tested with IsArrayEmpty.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Ndx As Integer
    Dim res As Integer
    On Error Resume Next
    ' Loop, increasing the dimension index Ndx, until an error occurs.
    ' An error will occur when Ndx exceeds the number of dimension
    ' in the array. Return Ndx - 1.
    Do
        Ndx = Ndx + 1
        res = UBound(arr, Ndx)
    Loop Until Err.Number <> 0
    
    NumberOfArrayDimensions = Ndx - 1

End Function


Private Function ReverseArrayInPlace(InputArray As Variant, _
    Optional NoAlerts As Boolean = False) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ReverseArrayInPlace
    ' This procedure reverses the order of an array in place -- this is, the array variable
    ' in the calling procedure is sorted. An error will occur if InputArray is not an array,
     'if it is an empty, unallocated array, or if the number of dimensions is not 1.
    '
    ' NOTE: Before calling the ReverseArrayInPlace procedure, consider if your needs can
    ' be met by simply reading the existing array in reverse order (Step -1). If so, you can save
    ' the overhead added to your application by calling this function.
    '
    ' The function returns TRUE if the array was successfully reversed, or FALSE if
    ' an error occurred.
    '
    ' If an error occurred, a message box is displayed indicating the error. To suppress
    ' the message box and simply return FALSE, set the NoAlerts parameter to TRUE.
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Temp As Variant
    Dim Ndx As Long
    Dim Ndx2 As Long
    Dim OrigN As Long
    Dim NewN As Long
    Dim NewArr() As Variant
    
    ''''''''''''''''''''''''''''''''
    ' Set the default return value.
    ''''''''''''''''''''''''''''''''
    ReverseArrayInPlace = False
    
    '''''''''''''''''''''''''''''''''
    ' Ensure we have an array
    '''''''''''''''''''''''''''''''''
    If IsArray(InputArray) = False Then
       If NoAlerts = False Then
            MsgBox "The InputArray parameter is not an array."
        End If
        Exit Function
    End If
    
    ''''''''''''''''''''''''''''''''''''''
    ' Test the number of dimensions of the
    ' InputArray. If 0, we have an empty,
    ' unallocated array. Get out with
    ' an error message. If greater than
    ' one, we have a multi-dimensional
    ' array, which is not allowed. Only
    ' an allocated 1-dimensional array is
    ' allowed.
    ''''''''''''''''''''''''''''''''''''''
    Select Case NumberOfArrayDimensions(InputArray)
        Case 0
            '''''''''''''''''''''''''''''''''''''''''''
            ' Zero dimensions indicates an unallocated
            ' dynamic array.
            '''''''''''''''''''''''''''''''''''''''''''
            If NoAlerts = False Then
                MsgBox "The input array is an empty, unallocated array."
            End If
            Exit Function
        Case 1
            '''''''''''''''''''''''''''''''''''''''''''
            ' We can reverse ONLY a single dimensional
            ' arrray.
            '''''''''''''''''''''''''''''''''''''''''''
        Case Else
            '''''''''''''''''''''''''''''''''''''''''''
            ' We can reverse ONLY a single dimensional
            ' arrray.
            '''''''''''''''''''''''''''''''''''''''''''
            If NoAlerts = False Then
                MsgBox "The input array multi-dimensional. ReverseArrayInPlace works only " & _
                       "on single-dimensional arrays."
            End If
            Exit Function
    
    End Select
    
    '''''''''''''''''''''''''''''''''''''''''''''
    ' Ensure that we have only simple data types,
    ' not an array of objects or arrays.
    '''''''''''''''''''''''''''''''''''''''''''''
    If IsSimpleDataType(InputArray(LBound(InputArray))) = False Then
        If NoAlerts = False Then
            MsgBox "The input array contains arrays, objects, or other complex data types." & vbCrLf & _
                "ReverseArrayInPlace can reverse only arrays of simple data types."
            Exit Function
        End If
    End If
    
    ReDim NewArr(LBound(InputArray) To UBound(InputArray))
    NewN = UBound(NewArr)
    For OrigN = LBound(InputArray) To UBound(InputArray)
        NewArr(NewN) = InputArray(OrigN)
        NewN = NewN - 1
    Next OrigN
    
    For NewN = LBound(NewArr) To UBound(NewArr)
        InputArray(NewN) = NewArr(NewN)
    Next NewN
    
    ReverseArrayInPlace = True
End Function


Private Function ReverseArrayInPlace2(InputArray As Variant, _
    Optional lb As Long = -1, Optional ub As Long = -1, _
    Optional NoAlerts As Boolean = False) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ReverseArrayInPlace2
    ' This reverses the order of elements in InputArray. To reverse the entire array, omit or
    ' set to less than 0 the LB and UB parameters. To reverse only part of tbe array, set LB and/or
    ' UB to the LBound and UBound of the sub array to be reversed.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim n As Long
    Dim Temp As Variant
    Dim Ndx As Long
    Dim Ndx2 As Long
    Dim OrigN As Long
    Dim NewN As Long
    Dim NewArr() As Variant
    
    ''''''''''''''''''''''''''''''''
    ' Set the default return value.
    ''''''''''''''''''''''''''''''''
    ReverseArrayInPlace2 = False
    
    '''''''''''''''''''''''''''''''''
    ' Ensure we have an array
    '''''''''''''''''''''''''''''''''
    If IsArray(InputArray) = False Then
        If NoAlerts = False Then
            MsgBox "The InputArray parameter is not an array."
        End If
        Exit Function
    End If
    
    ''''''''''''''''''''''''''''''''''''''
    ' Test the number of dimensions of the
    ' InputArray. If 0, we have an empty,
    ' unallocated array. Get out with
    ' an error message. If greater than
    ' one, we have a multi-dimensional
    ' array, which is not allowed. Only
    ' an allocated 1-dimensional array is
    ' allowed.
    ''''''''''''''''''''''''''''''''''''''
    Select Case NumberOfArrayDimensions(InputArray)
        Case 0
            '''''''''''''''''''''''''''''''''''''''''''
            ' Zero dimensions indicates an unallocated
            ' dynamic array.
            '''''''''''''''''''''''''''''''''''''''''''
            If NoAlerts = False Then
                MsgBox "The input array is an empty, unallocated array."
            End If
            Exit Function
        Case 1
            '''''''''''''''''''''''''''''''''''''''''''
            ' We can reverse ONLY a single dimensional
            ' arrray.
            '''''''''''''''''''''''''''''''''''''''''''
        Case Else
            '''''''''''''''''''''''''''''''''''''''''''
            ' We can reverse ONLY a single dimensional
            ' arrray.
            '''''''''''''''''''''''''''''''''''''''''''
            If NoAlerts = False Then
                MsgBox "The input array multi-dimensional. ReverseArrayInPlace works only " & _
                       "on single-dimensional arrays."
            End If
            Exit Function
    
    End Select
    
    '''''''''''''''''''''''''''''''''''''''''''''
    ' Ensure that we have only simple data types,
    ' not an array of objects or arrays.
    '''''''''''''''''''''''''''''''''''''''''''''
    If IsSimpleDataType(InputArray(LBound(InputArray))) = False Then
        If NoAlerts = False Then
            MsgBox "The input array contains arrays, objects, or other complex data types." & vbCrLf & _
                "ReverseArrayInPlace can reverse only arrays of simple data types."
            Exit Function
        End If
    End If
    
    If lb < 0 Then
        lb = LBound(InputArray)
    End If
    If ub < 0 Then
        ub = UBound(InputArray)
    End If
    
    For n = lb To (lb + ((ub - lb - 1) \ 2))
        Temp = InputArray(n)
        InputArray(n) = InputArray(ub - (n - lb))
        InputArray(ub - (n - lb)) = Temp
    Next n
    
    ReverseArrayInPlace2 = True
End Function


Private Function IsSimpleNumericType(v As Variant) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' IsSimpleNumericType
    ' This returns TRUE if V is one of the following data types:
    '        vbBoolean
    '        vbByte
    '        vbCurrency
    '        vbDate
    '        vbDecimal
    '        vbDouble
    '        vbInteger
    '        vbLong
    '        vbSingle
    '        vbVariant if it contains a numeric value
    ' It returns FALSE for any other data type, including any array
    ' or vbEmpty.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If IsSimpleDataType(v) = True Then
        Select Case VarType(v)
            Case vbBoolean, _
                    vbByte, _
                    vbCurrency, _
                    vbDate, _
                    vbDecimal, _
                    vbDouble, _
                    vbInteger, _
                    vbLong, _
                    vbSingle
                IsSimpleNumericType = True
            Case vbVariant
                If IsNumeric(v) = True Then
                    IsSimpleNumericType = True
                Else
                    IsSimpleNumericType = False
                End If
            Case Else
                IsSimpleNumericType = False
        End Select
    Else
        IsSimpleNumericType = False
    End If
End Function

Private Function IsSimpleDataType(v As Variant) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' IsSimpleDataType
    ' This function returns TRUE if V is one of the following
    ' variable types (as returned by the VarType function:
    '    vbBoolean
    '    vbByte
    '    vbCurrency
    '    vbDate
    '    vbDecimal
    '    vbDouble
    '    vbEmpty
    '    vbError
    '    vbInteger
    '    vbLong
    '    vbNull
    '    vbSingle
    '    vbString
    '    vbVariant
    '
    ' It returns FALSE if V is any one of the following variable
    ' types:
    '    vbArray
    '    vbDataObject
    '    vbObject
    '    vbUserDefinedType
    '    or if it is an array of any type.
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    On Error Resume Next
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Test if V is an array. We can't just use VarType(V) = vbArray
    ' because the VarType of an array is vbArray + VarType(type
    ' of array element). E.g, the VarType of an Array of Longs is
    ' 8195 = vbArray + vbLong.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If IsArray(v) = True Then
        IsSimpleDataType = False
        Exit Function
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' We must also explicitly check whether V is an object, rather
    ' relying on VarType(V) to equal vbObject. The reason is that
    ' if V is an object and that object has a default proprety, VarType
    ' returns the data type of the default property. For example, if
    ' V is an Excel.Range object pointing to cell A1, and A1 contains
    ' 12345, VarType(V) would return vbDouble, the since Value is
    ' the default property of an Excel.Range object and the default
    ' numeric type of Value in Excel is Double. Thus, in order to
    ' prevent this type of behavior with default properties, we test
    ' IsObject(V) to see if V is an object.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If IsObject(v) = True Then
        IsSimpleDataType = False
        Exit Function
    End If
    '''''''''''''''''''''''''''''''''''''
    ' Test the value returned by VarType.
    '''''''''''''''''''''''''''''''''''''
    Select Case VarType(v)
        Case vbArray, vbDataObject, vbObject, vbUserDefinedType
            '''''''''''''''''''''''
            ' not simple data types
            '''''''''''''''''''''''
            IsSimpleDataType = False
        Case Else
            ''''''''''''''''''''''''''''''''''''
            ' otherwise it is a simple data type
            ''''''''''''''''''''''''''''''''''''
            IsSimpleDataType = True
    End Select

End Function

Private Function IsArrayAllocated(arr As Variant) As Boolean
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' IsArrayAllocated
    ' Returns TRUE if the array is allocated (either a static array or a dynamic array that has been
    ' sized with Redim) or FALSE if the array has not been allocated (a dynamic that has not yet
    ' been sized with Redim, or a dynamic array that has been Erased).
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim n As Long
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    ' If Arr is not an array, return FALSE and get out.
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    If IsArray(arr) = False Then
        IsArrayAllocated = False
        Exit Function
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Try to get the UBound of the array. If the array has not been allocated,
    ' an error will occur. Test Err.Number to see if an error occured.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    On Error Resume Next
    n = UBound(arr, 1)
    If Err.Number = 0 Then
        '''''''''''''''''''''''''''''''''''''
        ' No error. Array has been allocated.
        '''''''''''''''''''''''''''''''''''''
        IsArrayAllocated = True
    Else
        '''''''''''''''''''''''''''''''''''''
        ' Error. Unallocated array.
        '''''''''''''''''''''''''''''''''''''
        IsArrayAllocated = False
    End If

End Function

Private Function AllocationCheck(arr As Variant) As AllocationEnum
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' IsArrayAllocated
    ' Returns TRUE if the array is allocated (either a static array or a dynamic array that has been
    ' sized with Redim) or FALSE if the array has not been allocated (a dynamic that has not yet
    ' been sized with Redim, or a dynamic array that has been Erased).
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim n As Long
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    ' If Arr is not an array, return FALSE and get out.
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    If IsArray(arr) = False Then
        AllocationCheck = isFalse
        Exit Function
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Try to get the UBound of the array. If the array has not been allocated,
    ' an error will occur. Test Err.Number to see if an error occured.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    On Error Resume Next
    n = UBound(arr, 1)
    If Err.Number = 0 Then
        '''''''''''''''''''''''''''''''''''''
        ' No error. Array has been allocated.
        '''''''''''''''''''''''''''''''''''''
        AllocationCheck = isTrue
    Else
        '''''''''''''''''''''''''''''''''''''
        ' Error. Unallocated array.
        '''''''''''''''''''''''''''''''''''''
        AllocationCheck = isFalse
    End If

End Function

'##########################################################################################################
'##                                  END   SORT IN PLACE - CHIP PEARSON                                  ##
'##########################################################################################################
