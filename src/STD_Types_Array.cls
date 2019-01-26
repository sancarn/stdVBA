VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "STD_Types_Array"
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

Public Enum SortDirection
    Ascending = 1
    Descending = 2
End Enum

Private pArr As Variant
Private pInitialised as Boolean

Public Function Create(ParamArray params() As Variant) As clsDynamicArray
    set Create = New clsDynamicArray
    Create.Arr = params
    Create.Init
End Function

Property Get Arr() as variant
  Arr = pArr
End Property

Property set Arr(val as variant)
  if not pInitialised then
    pArr = val
  else
    Err.Raise 0, "STD.Types.Array::setArr()", "Cannot set array of already initialised array."
  end if
End Property

Friend Sub Init()
  if not pInitialised then
    pInitialised = true
  End If
End Sub

Public Sub push(ByVal el As Variant)
  If pInitialised then
    If IsArrayAllocated(pArr) then
      ReDim Preserve pArr(UBound(pArr) + 1)
      pArr(UBound(pArr)) = el
    Else
      ReDim Preserve pArr(0)
      pArr(0) = e
    End if
  Else
    'Error
  End If
End Sub

Public Function pop() as variant
  If pInitialised then
    If IsArrayAllocated(pArr) then
      pop = pArr(UBound(pArr))
      ReDim Preserve pArr(UBound(pArr) - 1)
    End If
  Else
    'Error
  End If  
End Function

Public Function shift() as variant
  If pInitialised then
    If IsArrayAllocated(pArr) then
      shift = pArr(LBound(pArr))
      ReDim Preserve pArr(LBound(pArr) + 1, UBound(pArr)) 'This may not work?
    End If
  Else
    'Error
  End If
End Function
Public Function unshift(val as variant) as STD_Types_Array
  If pInitialised then
    If IsArrayAllocated(pArr) then
      ReDim Preserve pArr(LBound(pArr) - 1, UBound(pArr)) 'This may not work?
      pArr(LBound(pArr)) = val
    End If
  Else
    'Error
  End If
  set unshift = Me
End Function

Public Function Slice() as STD_Types_Array

End Function

Public Function Splice() as STD_Types_Array

End Function

Public Function Join(Optional ByVal delimeter As String = ",") As String
  If pInitialised then
    If IsArrayAllocated(pArr) then
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

Public Function Clone() as STD_Types_Array
  If pInitialised then
    set Clone = new STD_Types_Array
    Clone.Arr = pArr
    Clone.Init()
  Else
    'Error
  End If
End Function


'Iterative Functions (All require STD_Types_Callback):
Public Function every(ByVal cb As STD_Types_Callback) as Boolean
  If pInitialised then
    If IsArrayAllocated(pArr) then
      Dim i As Integer
      For i = LBound(pArr) To UBound(pArr)
        If Not cb.call(pArr(i)) then
          every = false
          Exit Function
        End If
      Next
      evert = true
    Else
      every = false
    End if
  Else
    'Error
  End If
End Function

Public Function some(ByVal cb As STD_Types_Callback) as Boolean
  If pInitialised then
    If IsArrayAllocated(pArr) then
      Dim i As Integer, test as boolean
      For i = LBound(pArr) To UBound(pArr)
        If cb.call(pArr(i)) then
          some = true
          Exit Function
        End If
      Next
      some = false
    Else
      some = false
    End if
  Else
    'Error
  End If
End Function

Public Sub forEach(ByVal cb As STD_Types_Callback)
  If pInitialised then
    If IsArrayAllocated(pArr) then
      Dim i As Integer
      For i = LBound(pArr) To UBound(pArr)
        Call cb.call(pArr(i))
      Next
    End If
  Else
    'Error
  End If
End Sub

Public Function Map(ByVal cb as STD_Types_Callback) as STD_Types_Array
  If pInitialised then
    Dim pMap as STD_Types_Array
    set pMap = Me.Clone()
    If IsArrayAllocated(pArr) then
      Dim i As Integer
      For i = LBound(pArr) To UBound(pArr)
        pMap(i) = cb.call(pArr(i))
      Next
    End If
    set Map = pMap
  Else
    'Error
  End If
End Function

Public Function Reduce(ByVal cb as STD_Types_Callback, Optional ByVal initialValue as Variant) as variant
  If pInitialised then
    If IsArrayAllocated(pArr) then
      Reduce = initialValue
      Dim i As Integer
      For i = LBound(pArr) To UBound(pArr)
        Reduce = cb.call(Reduce,pArr(i))
      Next
    Else
      Reduce = initialValue
    End If
  Else
    'Error
  End If
End Function

Public Function Sort(ByVal cb as STD_Types_Callback) as STD_Types_Array
  Msgbox "NOT IMPLEMENTED"
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

Public Function indexOf(ByVal el As Variant,Optional ByVal start as integer)
  If not start then start = UBound(pArr)
  if start < 0 then start = UBound(pArr) - start + 1
  
  'find the index of an item in the array. If array doesn't contain item, return -1
  Dim i As Integer
  For i = start To UBound(pArr)
    If pArr(i) = el Then
      indexOf = i
      Exit Function
    End If
  Next
  indexOf = -1
End Function

Public Function lastIndexOf(ByVal el as Variant,Optional ByVal start as integer)
  If not start then start = UBound(pArr)
  if start < 0 then start = UBound(pArr) - start + 1
  
  'find the index of an item in the array. If array doesn't contain item, return -1
  Dim i As Integer
  For i = start to LBound(pArr) step -1
    If pArr(i) = el Then
      indexOf = i
      Exit Function
    End If
  Next
  indexOf = -1
End Function

Public Function includes(ByVal el as variant, Optional ByVal startFrom as Integer = -1) as boolean
  includes = indexOf(el) >= startFrom
End Function

Public Function Length()
    'get the total number of elements in the array
    Length = UBound(pArr) - LBound(pArr) + 1
End Function

Public Function item(ByVal i As Integer)
    'item(1) = 1st element
    'item(2) = 2nd element
    'etc.
    item = pArr(LBound(pArr) + i - 1)
End Function














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
    Dim Res As Integer
    On Error Resume Next
    ' Loop, increasing the dimension index Ndx, until an error occurs.
    ' An error will occur when Ndx exceeds the number of dimension
    ' in the array. Return Ndx - 1.
    Do
        Ndx = Ndx + 1
        Res = UBound(arr, Ndx)
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

'##########################################################################################################
'##                                  END   SORT IN PLACE - CHIP PEARSON                                  ##
'##########################################################################################################
