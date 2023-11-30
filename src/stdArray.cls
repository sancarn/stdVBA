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

#If Mac Then
  Private Declare PtrSafe Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As LongPtr) As LongPtr
#ElseIf VBA6 Then
  Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As Long)
#Else
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
#End If

Private Enum SortDirection
    Ascending = 1
    Descending = 2
End Enum
Private Type SortStruct
  Value as variant
  SortValue as variant
End Type

'@TODO: Convert to TThis
Private pArr() As Variant
Private pProxyLength As Long
Private pLength As Long

Private pChunking As Long
Private pInitialised As Boolean


'Event executed before the internal array is overwritten 
'@param arr - A reference to this array
'@param arr2 - The array which is being assigned to this array
Public Event BeforeArrLet(ByRef arr as stdArray, ByRef arr2 as variant)

'Event executed after the internal array is overwritten
'@param arr - A reference to this array
'@param arr2 - The array which is being assigned to this array
Public Event AfterArrLet(ByRef arr as stdArray, ByRef arr2 as variant)

'Event executed before an item is added to the array
'@param arr - A the array to which the item is being added
'@param iIndex - The index at which the item will be added
'@param item - The item which will be added
'@param cancel - Set to true to cancel the addition 
Public Event BeforeAdd(ByRef arr As stdArray, ByVal iIndex As Long, ByRef item As Variant, ByRef cancel As Boolean)

'Event executed after an item is added to the array
'@param arr - A the array to which the item has been added to.
'@param iIndex - The index at which the item was added
'@param item - The item which was be added
Public Event AfterAdd(ByRef arr as stdArray, ByVal iIndex as long, ByRef item as variant)

'Event executed before an item is removed from the array
'@param arr - A the array from which the item is being removed
'@param iIndex - The index at which the item will be removed
'@param item - The item which will be removed
'@param cancel - Set to true to cancel the removal
Public Event BeforeRemove(ByRef arr as stdArray, ByVal iIndex as long, ByRef item as variant, ByRef cancel as Boolean)

'Event executed after an item is removed from the array
'@param arr - A the array from which the item has been removed
'@param iIndex - The index at which the item was removed
Public Event AfterRemove(ByRef arr as stdArray, ByVal iIndex as long)

'Event executed after an array is cloned
'@param clone - A reference to the clone
Public Event AfterClone(ByRef clone as stdArray)

'Event executed after an array is created
'@param arr - A reference to the array
Public Event AfterCreate(ByRef arr as stdArray)

'Create a `stdArray` object from params
'@param params - The items of the array
'@returns stdArray<variant> - A `stdArray` from the parameters.
Public Function Create(ParamArray params() As Variant) As stdArray
    Set Create = New stdArray
    
    Dim i As Long
    Dim lb As Long: lb = LBound(params)
    Dim ub As Long: ub = UBound(params)
    
    Call Create.protInit(ub - lb + 1, 10)
    
    For i = lb To ub
        Call Create.Push(params(i))
    Next

    'Raise AfterCreate event
    RaiseEvent AfterCreate(Create)
End Function

'Create a `stdArray` object from params
'@param iInitialLength - The length of the initial private array created
'@param iChunking - The number of items the private array is increased by when required.
'@param params - The items of the array
'@returns stdArray<variant> - A `stdArray` from the parameters.
Public Function CreateWithOptions(ByVal iInitialLength As Long, ByVal iChunking As Long, ParamArray params() As Variant) As stdArray
    Set CreateWithOptions = New stdArray
    
    Dim i As Long
    Dim lb As Long: lb = LBound(params)
    Dim ub As Long: ub = UBound(params)
    
    Call CreateWithOptions.protInit(iInitialLength, iChunking)
    For i = lb To ub
        Call CreateWithOptions.Push(params(i))
    Next

    'Raise AfterCreate event
    RaiseEvent AfterCreate(Create)
End Function

'Create a `stdArray` object from a VBA array
'@param arr - Variant array to create a `stdArray` object from.
'@returns stdArray<variant> - Returns `stdArray` of variants. 
Public Function CreateFromArray(ByVal arr As Variant) As stdArray
    Set CreateFromArray = New stdArray
    
    Dim i As Long
    Dim lb As Long: lb = LBound(arr)
    Dim ub As Long: ub = UBound(arr)
    Call CreateFromArray.protInit(ub - lb + 1, 10)
    
    For i = lb To ub
        Call CreateFromArray.Push(arr(i))
    Next

    'Raise AfterCreate event
    RaiseEvent AfterCreate(Create)
End Function

'Create an array by splitting a string
'@param sHaystack  - Haystack to split
'@param sDelimiter - Delimiter
'@returns stdArray<String> - A list of strings
Public Function CreateFromString(ByVal sHaystack as string, Optional byval sDelimiter as string = ",") as stdArray
  set CreateFromString = CreateFromArray(split(sHaystack, sDelimiter))
End Function

'Initialise array
'@protected
'@param iInitialLength - The length of the initial private array created
'@param iChunking      - The number of items the private array is increased by when required.
Public Sub protInit(ByVal iInitialLength As Long, ByVal iChunking As Long)
  If iChunking > iInitialLength Then iInitialLength = iChunking
  If Not pInitialised Then
    pProxyLength = iInitialLength
    ReDim pArr(1 To iInitialLength) As Variant
    pChunking = iChunking
    pInitialised = True
  End If
End Sub

'Obtain a collection from the data contained within the array. Primarily used for NewEnum() method.
'@returns Collection - Collection from Array
Public Function AsCollection() as Collection
  set AsCollection = new Collection
  Dim i as long
  For i = 1 To Length()
    AsCollection.add pArr(i)
  next
End function

'For-each compatibility
'@protected
'@returns IEnumVARIANT - An enumerator with methods enumNext, enumRefresh etc.
'@example `For each obj in myEnum: ... : next`
'@TODO: Use custom `IEnumVARIANT` instead of casting to `Collection`
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Static oEnumCol As Collection: If oEnumCol Is Nothing Then Set oEnumCol = AsCollection()
    Set NewEnum = oEnumCol.[_NewEnum]
End Property

'Obtain the length of the array
'@returns Long - Length of the array
Public Property Get Length() As Long
    Length = pLength
End Property

'Obtain the length of the private array which stores the data of this array class
'@protected
'@returns Long - Length of the private array
Public Property Get zProxyLength() As Long
    zProxyLength = pProxyLength
End Property

'Resize the array to a length
'@param iLength - The length of the desired array
Public Sub Resize(ByVal iLength As Long)
  pLength = iLength
End Sub

'Rechunk the private array to the length / number of items.
'@protected
Public Sub Rechunk()
  Dim fNumChunks As Double, iNumChunks As Long
  fNumChunks = pLength / pChunking
  iNumChunks = CLng(fNumChunks)
  If fNumChunks > iNumChunks Then iNumChunks = iNumChunks + 1
  
  ReDim Preserve pArr(1 To iNumChunks * pChunking) As Variant
End Sub


'Sort the array
'@param cbSortBy as stdICallable<(variant)=>variant>              - A mapping function which should map whatever the input is to whatever variant the array should be sorted on.  
'@param cbComparrason as stdICallable<(variant,variant)=>boolean> - Comparrison function which consumes 2 variants and generates a boolean. See implementation of `Sort_QuickSort` for details. 
'@param iAlgorithm                                                - Currently only 1 algorithm: 0 - Quicksort
'@param bSortInPlace                                              - Sort the array in place. Sorting in-place is prefferred if possible as it is much more performant.
'@returns stdArray                                                - A sorted array
Public Function Sort(Optional ByVal cbSortBy As stdICallable = Nothing, Optional ByVal cbComparrason As stdICallable = Nothing, Optional ByVal iAlgorithm As Long = 0, Optional ByVal bSortInPlace As Boolean = False) As stdArray
    If Not bSortInPlace Then
        Set Sort = Clone().Sort(cbSortBy, cbComparrason, iAlgorithm, True)
    Else
        If Length() = 0 then
          set Sort = Me
          Exit Function
        End if

        Dim arr() As SortStruct
        ReDim arr(1 To Length()) As SortStruct
        
        Dim i As Long
        
        'Copy array to sort structures
        For i = 1 To Length()
            Call CopyVariant(arr(i).Value, pArr(i))
            If cbSortBy Is Nothing Then
                Call CopyVariant(arr(i).SortValue, pArr(i))
            Else
                Call CopyVariant(arr(i).SortValue, cbSortBy.Run(pArr(i)))
            End If
        Next
        
        'Call sort algorithm
        Select Case iAlgorithm
            Case 0 'QuickSort
                Call Sort_QuickSort(arr, cbComparrason)
            Case Else
                stdError.Raise "Invalid sorting algorithm specified"
        End Select
        
        'Copy sort structures to array
        For i = 1 To Length()
            Call CopyVariant(pArr(i), arr(i).Value)
        Next
        
        'Return array
        Set Sort = Me
    End If
End Function

'QuickSort3
'@private
'@param pvarArray                                                 - Array to sort
'@param cbComparrison as stdICallable<(variant,variant)=>boolean> - Comparrison function which consumes 2 variants and generates a boolean. See implementation of `Sort_QuickSort` for details.
'@param plngLeft                                                  - Left index of array to sort
'@param plngRight                                                 - Right index of array to sort
'@remark Omit plngLeft & plngRight; they are used internally during recursion
'@dev https://www.vbforums.com/showthread.php?473677-VB6-Sorting-algorithms-%28sort-array-sorting-arrays%29
Private Sub Sort_QuickSort(ByRef pvarArray() As SortStruct, Optional cbComparrison As stdICallable = nothing, Optional ByVal plngLeft As Long, Optional ByVal plngRight As Long)
    Dim lngFirst As Long
    Dim lngLast As Long
    Dim varMid As SortStruct
    Dim varSwap As SortStruct
    
    If plngRight = 0 Then
        plngLeft = 1
        plngRight = Length()
    End If
    lngFirst = plngLeft
    lngLast = plngRight
    varMid = pvarArray((plngLeft + plngRight) \ 2)
    Do
        If cbComparrison Is Nothing Then
            Do While pvarArray(lngFirst).SortValue < varMid.SortValue And lngFirst < plngRight
                lngFirst = lngFirst + 1
            Loop
            Do While varMid.SortValue < pvarArray(lngLast).SortValue And lngLast > plngLeft
                lngLast = lngLast - 1
            Loop
        Else
            Do While cbComparrison.Run(pvarArray(lngFirst).SortValue, varMid.SortValue) And lngFirst < plngRight
                lngFirst = lngFirst + 1
            Loop
            Do While cbComparrison.Run(varMid.SortValue, pvarArray(lngLast).SortValue) And lngLast > plngLeft
                lngLast = lngLast - 1
            Loop
        End If
        
        If lngFirst <= lngLast Then
            varSwap = pvarArray(lngFirst)
            pvarArray(lngFirst) = pvarArray(lngLast)
            pvarArray(lngLast) = varSwap
            lngFirst = lngFirst + 1
            lngLast = lngLast - 1
        End If
    Loop Until lngFirst > lngLast
    If plngLeft < lngLast Then Sort_QuickSort pvarArray, cbComparrison, plngLeft, lngLast
    If lngFirst < plngRight Then Sort_QuickSort pvarArray, cbComparrison, lngFirst, plngRight
End Sub

'Obtain the array as a regular VBA array
'@returns Variant - A VBA array
Public Property Get arr() As Variant
    if pLength = 0 then
      arr = Array()
    else
      Dim vRet() As Variant
      ReDim vRet(1 To pLength) As Variant
      For i = 1 To pLength
        Call CopyVariant(vRet(i), pArr(i))
      Next
      arr = vRet
    end if
End Property
Public Property Let arr(v As Variant)
    RaiseEvent BeforeArrLet(me,v)
    Dim lb As Long: lb = LBound(v)
    Dim ub As Long: ub = UBound(v)
    Dim cnt As Long: cnt = ub - lb + 1
    ReDim pArr(1 To (Int(cnt / pChunking) + 1) * pChunking) As Variant
    For i = lb To ub
        Call Push(pArr(i))
    Next
    RaiseEvent AfterArrLet(me,v)
End Property

'Add an element to the end of the array
'@param el - The element to add to the end of the array.
'@returns - A reference to the array to allow chaining.
'@TODO: Add multiple elements with push
Public Function Push(ByVal el As Variant) as stdArray
  If pInitialised Then
    'Before Add event
    Dim bCancel as Boolean
    RaiseEvent BeforeAdd(me, pLength + 1, el, bCancel)
    if bCancel then exit function

    If pLength = pProxyLength Then
        pProxyLength = pProxyLength + pChunking
        ReDim Preserve pArr(1 To pProxyLength) As Variant
    End If
    
    pLength = pLength + 1
    CopyVariant pArr(pLength), el

    'After add event
    RaiseEvent AfterAdd(me, pLength, pArr(pLength))

    set Push = me
  Else
    'Error
  End If
End Function

'Remove an element from the end of the array
'@returns - The element removed from the array
Public Function Pop() As Variant
    If pInitialised Then
        If pLength > 0 Then
            'Raise BeforeRemove event and optionally cancel
            Dim bCancel as Boolean
            RaiseEvent BeforeRemove(me, pLength, pArr(pLength), bCancel)
            If bCancel then exit function

            CopyVariant Pop, pArr(pLength)
            pLength = pLength - 1

            'Raise AfterRemove event
            RaiseEvent AfterRemove(me, pLength)
        Else
            Pop = Empty
        End If
    Else
        'Error
    End If
End Function

'Remove the ith element from the array
'@param index - Index of the element to remove
'@returns - The element removed
Public Function Remove(ByVal index As Long) As Variant
  'Ensure initialised
  If pInitialised Then
    'Ensure length > 0
    If pLength > 0 Then
      'Ensure index < length
      If index <= pLength Then
        'Raise BeforeRemove event and optionally cancel
        Dim bCancel as Boolean
        RaiseEvent BeforeRemove(me, index, pArr(index), bCancel)
        If bCancel then exit function

        'Copy party we are removing to return variable
        CopyVariant Remove, pArr(index)
        
        'Loop through array from removal, set i-1th element to ith element
        Dim i As Long
        For i = index + 1 To pLength
            pArr(i - 1) = pArr(i)
        Next

        'Set last element length and subtract total length by 1
        pArr(pLength) = Empty
        pLength = pLength - 1

        'Raise after remove event
        RaiseEvent AfterRemove(me, index)
      Else
        'Error
      End If
    Else
      'Error
    End If
  Else
      'Error
  End If
End Function

'Remove and return the first element from the array
'@returns - Element removed
Public Function Shift() As Variant
  'Would be good to use CopyMemory here
  
  CopyVariant Shift, pArr(1)
  Dim i As Long
  For i = 1 To pLength - 1
    Call CopyVariant(pArr(i), pArr(i+1))
  Next
  pLength = pLength -1
End Function

'Add an element to the start of the array
'@param val - Value to append to the start of the array
'@returns - A reference to the array to allow chaining.
Public Function Unshift(val As Variant) As stdArray
  'Would be good to use CopyMemory here
  
  'Before Add event
  Dim bCancel as Boolean
  RaiseEvent BeforeAdd(me, 1, val, bCancel)
  if bCancel then exit Function

  'Ensure array is big enough and increase pLength
  If pLength = pProxyLength Then
    pProxyLength = pProxyLength + pChunking
    ReDim Preserve pArr(1 To pProxyLength) As Variant
  End If
  pLength = pLength + 1
  
  'Unshift
  For i = pLength - 1 To 1 Step -1
    pArr(i + 1) = pArr(i)
  Next
  pArr(1) = val
  
  'After Add event
  RaiseEvent AfterAdd(me, 1, val)

  Set Unshift = Me
End Function

'Obtain a slice of the array
'@param iStart - Start index of the slice
'@param iEnd   - End index of the slice
'@returns - A slice of the array
Public Function Slice(ByVal iStart As Long, ByVal iEnd As Long) As stdArray
  If pInitialised Then
    If iStart <= pLength And iEnd <= pLength Then
      If iStart <= iEnd Then
        Dim ret As stdArray
        Set ret = stdArray.CreateWithOptions(iEnd - iStart + 1, pChunking)
        
        Dim i As Long
        For i = iStart To iEnd
          Call ret.Push(pArr(i))
        Next
        
        Set Slice = ret
      Else
        'Error
      End If
    Else
      'Error
    End If
  Else
    'Error
  End If
End Function

'Splice the array
'@param iStart - Start index of the splice
'@param iDeleteCount - Number of elements to delete
'@param newElements - Elements to add to the array
'@returns - An array containing the deleted elements
Public Function Splice(ByVal iStart As Long, ByVal iDeleteCount As Long, ByVal newElements As stdArray) As stdArray
  If pInitialised Then
    If iStart <= pLength Then
      If iStart + iDeleteCount <= pLength Then
        Dim ret As stdArray
        Set ret = stdArray.CreateWithOptions(iDeleteCount, pChunking)
        
        Dim i As Long
        For i = 1 To iDeleteCount
          Call ret.Push(pArr(iStart + i - 1))
        Next
        
        Dim iNewLength As Long
        iNewLength = pLength - iDeleteCount + newElements.Length
        If iNewLength > pProxyLength Then
          pProxyLength = iNewLength + pChunking
          ReDim Preserve pArr(1 To pProxyLength) As Variant
        End If
        
        For i = pLength To iStart + iDeleteCount Step -1
          pArr(i + newElements.Length - iDeleteCount) = pArr(i)
        Next
        
        For i = LBound(newElements) To UBound(newElements)
          pArr(iStart + i - 1) = newElements(i)
        Next
        
        pLength = iNewLength
        
        Set Splice = ret
      Else
        'Error
      End If
    Else
      'Error
    End If
  Else
    'Error
  End If
End Function

'Creates a new instance of the same array
'@returns - A new instance of the same array
Public Function Clone() As stdArray
  If pInitialised Then
      If pInitialised Then
        'Similar to CreateFromArray() but passing length through also:
        Set Clone = New stdArray
        
        Call Clone.protInit(pLength, 10)
        
        Dim i As Long
        For i = 1 To pLength
            Call Clone.Push(pArr(i))
        Next
      Else
        'Error
      End If

      RaiseEvent AfterClone(Clone)
  Else
    'Error
  End If
End Function

'Returns a new array with all elements in reverse order
'@returns - A new array with all elements in reverse order
Public Function Reverse() As stdArray
    'TODO: Need to find a better more low level approach to creating arrays from existing arrays/preventing redim for methods like this
    Dim ret As stdArray
    Set ret = stdArray.Create()
    For i = pLength To 1 Step -1
      Call ret.Push(pArr(i))
    Next
    Set Reverse = ret
End Function

'Concatenate an existing array of elements onto the end of this array
'@param arr - Array whose elements we wish to append to the end of this array
'@returns - New composite array.
Public Function Concat(ByVal arr As stdArray) As stdArray
  Dim x As stdArray
  Set x = Clone()
  
  If Not arr Is Nothing Then
    Dim i As Long
    For i = 1 To arr.Length
      Call x.Push(arr.item(i))
    Next
  End If
  
  Set Concat = x
End Function

'Join each of the elements of this array together as a string
'@param delimiter - Delimiter to insert between strings
'@returns - A string containing all elements of the array joined together with the delimiter
Public Function Join(Optional ByVal delimiter As String = ",") As String
  If pInitialised Then
      If pLength > 0 Then
        Dim sOutput As String
        sOutput = pArr(1)
        
        Dim i As Long
        For i = 2 To pLength
          sOutput = sOutput & delimiter & pArr(i)
        Next
        Join = sOutput
      Else
        Join = ""
      End If
  Else
    'Error
  End If
End Function

'Get/Let/Set item
'@defaultMember
'@param i - The location to get/set the item
Public Property Get item(ByVal i As long) As Variant
Attribute item.VB_UserMemId = 0
    CopyVariant item, pArr(i)
End Property
Public Property Set item(ByVal i As long, ByVal item As Object)
  set pArr(i) = item
End Property
Public Property Let item(ByVal i As long, ByVal item As Variant)
  pArr(i) = item
End Property

'Copy a variant into the array's ith element. This saves from having to test the item and call the correct `set` keyword
'@param i    - The index at which the item's data should be set
'@param item - Item to set at the index
Public Sub PutItem(ByVal i As long, ByRef item As Variant)
  CopyVariant pArr(i), item
End Sub

'Obtain the index of an element
'@param el - Element to find
'@param start - Location to start search for element.
'@returns - Index of element
Public Function indexOf(ByVal el As Variant, Optional ByVal start As long = 1) as long
  Dim elIsObj As Boolean, i As Long, item As Variant, itemIsObj As Boolean
  
  'Is element an object?
  elIsObj = IsObject(el)
  
  'Loop over contents starting from start
  For i = start To pLength
    'Get item data
    CopyVariant item, pArr(i)
    
    'Is item an object?
    itemIsObj = IsObject(item)
    
    'If both item and el are objects     (must be the same type in order to be the same data)
    If itemIsObj And elIsObj Then
      If item Is el Then 'check items equal
        indexOf = i 'return item index
        Exit Function
      End If
    'If both item and el are not objects (must be the same type in order to be the same data)
    ElseIf Not itemIsObj And Not elIsObj Then
      If item = el Then 'check items equal
        indexOf = i 'return item index
        Exit Function
      End If
    End If
  Next

  'Return -1 i.e. no match found
  indexOf = -1
End Function

'Obtain the last index of an element
'@param el - Element to find
'@returns - Last index of element
Public Function lastIndexOf(ByVal el As Variant) as Long
  Dim elIsObj As Boolean, i As Long, item As Variant, itemIsObj As Boolean
  
  'Is element an object?
  elIsObj = IsObject(el)
  
  'Loop over contents starting from start
  For i = pLength To 1 Step -1
    'Get item data
    CopyVariant item, pArr(i)
    
    'Is item an object?
    itemIsObj = IsObject(item)
    
    'If both item and el are objects     (must be the same type in order to be the same data)
    If itemIsObj And elIsObj Then
      If item Is el Then 'check items equal
        lastIndexOf = i 'return item index
        Exit Function
      End If
    'If both item and el are not objects (must be the same type in order to be the same data)
    ElseIf Not itemIsObj And Not elIsObj Then
      If item = el Then 'check items equal
        lastIndexOf = i 'return item index
        Exit Function
      End If
    End If
  Next

  'Return -1 i.e. no match found
  lastIndexOf = -1
End Function

'Returns true if the array contains an item
'@param el        - Item to find
'@param startFrom - Index to start search for item at. (Internally uses indexOf())
'@returns - True if the array contains the item
Public Function includes(ByVal el As Variant, Optional ByVal startFrom As long = 1) As Boolean
  includes = indexOf(el, startFrom) >= startFrom
End Function

'**************************************************
'* Iterative Functions (All require stdICallable) *
'**************************************************

'Check if all elements in the array pass a test
'@param cb as stdICallable<(element: Variant)=>Boolean> - Callback to run on each element
'@returns - True if all elements pass the test
'@example `if incidents.IsEvery(cbValid) then ...`
Public Function IsEvery(ByVal cb As stdICallable) As Boolean
  If pInitialised Then
    Dim i As Long
    For i = 1 To pLength
      Dim bFlag as Boolean
      bFlag = cb.run(pArr(i))
      
      If Not bFlag Then
        IsEvery = False
        Exit Function
      End If
    Next
    
    IsEvery = True
  Else
    'Error
  End If
End Function

'Check if any elements in the array pass a test
'@param cb as stdICallable<(element: Variant)=>Boolean> - Callback to run on each element
'@returns - True if any elements pass the test
'@example `if incidents.IsSome(cbValid) then ...`
Public Function IsSome(ByVal cb As stdICallable) As Boolean
  If pInitialised Then
    Dim i As long
    For i = 1 To pLength
      Dim bFlag as Boolean
      bFlag = cb.Run(pArr(i))
      
      if bFlag then
        IsSome = true
        Exit Function
      end if
    Next
    IsSome = False
  Else
    'Error
  End If
End Function

'Call a callback on each element of the array
'@param cb as stdICallable<(element: Variant)=>Void> - Callback to run on each element
'@example `incidents.ForEach(cbPrint)`
Public Sub ForEach(ByVal cb As stdICallable)
  If pInitialised Then
    Dim i As long
    For i = 1 To pLength
      Call cb.Run(pArr(i))
    Next
  Else
    'Error
  End If
End Sub

'Call a callback on each element of the array and return a new array with the results
'@param cb as stdICallable<(element: Variant)=>Variant> - Callback to run on each element
'@returns - A new array with the results of the callback
'@example `incidents.Map(cbGetId)`
Public Function Map(ByVal cb As stdICallable) As stdArray
  If pInitialised Then
    Dim pMap As stdArray
    Set pMap = Clone()
    
    Dim i As long
    For i = 1 To pLength
      'BUGFIX: Sometimes required, not sure when
      Dim v As Variant
      CopyVariant v, item(i)
      
      'Call callback
      Call pMap.PutItem(i, cb.Run(v))
    Next
    
    Set Map = pMap
  Else
    'Error
  End If
End Function

'Remove all non-unqiue elements from the array. Given some callback to generate a key for each element, this method will remove all elements which have the same key.
'@param cb as stdICallable<(element: Variant)=>Variant> - Callback to generate a key for each element
'@returns - A new array with all non-unique elements removed
'@example `incidents.Unique(cbGetId)`
'@TODO: Needs optimisation. Currently very sub-optimal
Public Function Unique(Optional ByVal cb As stdICallable = Nothing) As stdArray
  Dim ret As stdArray: Set ret = stdArray.CreateWithOptions(pLength, pChunking)
  Dim retL As stdArray: Set retL = CreateWithOptions(pLength, pChunking)
  
  'Collect keys
  Dim vKeys As stdArray
  If cb Is Nothing Then
    Set vKeys = Clone()
  Else
    Set vKeys = Map(cb)
  End If
  
  'Unique by key
  For i = 1 To pLength
    If Not retL.includes(vKeys.item(i)) Then
      Call retL.Push(vKeys.item(i))
      Call ret.Push(pArr(i))
    End If
  Next

  'Return data
  Set Unique = ret
End Function

'Executes a user-supplied "reducer" callback function on each element of the array, in order, passing in the return
'value from the calculation on the preceding element. The final result of running the reducer across all elements
'of the array is a single value.
'@param cb as stdICallable<(accumulator: Variant, element: Variant) => Variant> - Reducer callback to run on each element
'@param initialValue - Initial value to pass to the reducer callback
'@returns - The final value returned by the reducer callback
'@example `values.Reduce(cbSum, 0)`
Public Function Reduce(ByVal cb As stdICallable, Optional ByVal initialValue As Variant) As Variant
  Dim iStart as long
  If pInitialised Then
    if pLength > 0 then
      if isMissing(initialValue) then 
        Call CopyVariant(Reduce, pArr(1))
        iStart = 2
      else
        Call CopyVariant(Reduce, initialValue)
        iStart = 1
      end if
    else
      if isMissing(initialValue) then
        Reduce = Empty
      else
        Call CopyVariant(Reduce, initialValue)
      end if
      Exit Function
    end if 
    
    Dim i As long
    For i = iStart To pLength
      'BUGFIX: Sometimes required, not sure when
      Dim el As Variant
      CopyVariant el, pArr(i)

      'Reduce
      CopyVariant Reduce, cb.Run(Reduce, el)
    Next
  Else
    'Error
  End If
End Function

'Filter the array based on a condition
'@param cb as stdICallable<(element: Variant)=>Boolean> - Callback to run on each element. If the callback returns true, the element is included in the returned array.
'@returns - A new array containing only the elements which passed the filter
Public Function Filter(ByVal cb As stdICallable) As stdArray
    Dim ret As stdArray
    Set ret = stdArray.CreateWithOptions(pLength, pChunking)
    Set Filter = ret

    'If initialised...
    If pInitialised Then
        Dim i As Long, v As Variant
        'Loop over array
        For i = 1 To pLength
            'If callback succeeds, push retvar
            If cb.Run(pArr(i)) Then
                Call ret.Push(pArr(i))
            End If
        Next i
    Else
        'error
    End If
End Function

'Count the number of elements in the array. If a callback is provided, the number of elements which pass the callback is returned.
'@param cb as stdICallable<(element: Variant)=>Boolean> - Callback to run on each element. If the callback returns true, the element is included in the count.
'@returns - The number of elements in the array. If a callback is provided, the number of elements which pass the callback is returned.
Public Function Count(Optional ByVal cb As stdICallable = nothing) As Long
  if cb is nothing then
    Count = Length
  else
    Dim i As Long, lCount As Long
    lCount = 0
    For i = 1 To pLength
      If cb.Run(pArr(i)) Then
        lCount = lCount + 1
      End If
    Next i
    Count = lCount
  end if
End Function

'Group the elements of the array by some key generated by a callback
'@param cb as stdICallable<(element: Variant)=>Variant> - Callback to run on each element. The return value of this callback is used as the key to group the elements by.
'@returns Object<Dictionary<Variant, stdArray<Variant>>> - The keys of the dictionary are the keys generated by the callback. The values of the dictionary are arrays containing the elements which were grouped by the key.
Public Function GroupBy(ByVal cb As stdICallable) As Object
    'Array to store result in
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    'Loop over items
    Dim i As Long
    For i = 1 To pLength
        'Get grouping key
        Dim key As Variant
        key = cb.Run(pArr(i))

        'If key is not set then set it
        If Not result.exists(key) Then Set result(key) = stdArray.Create()

        'Push item to key
        result(key).Push pArr(i)
    Next

    'Return result
    Set GroupBy = result
End Function

'Obtain the maximum value in the array
'@param cb as stdICallable<(element: Variant)=>Variant> - Callback to run on each element. The return value of this callback is used to determine the maximum value.
'@param startingValue - The starting value to compare against. If not provided, the first element of the array is used.
'@returns - The maximum value in the array
Public Function Max(Optional ByVal cb As stdICallable = nothing, Optional ByVal startingValue As Variant = Empty) as variant
  Dim vRet, vMaxValue, v
  vMaxValue = startingValue: vRet = startingValue
  For i = 1 to pLength
    Call CopyVariant(v,pArr(i))
    
    'Get value to test
    Dim vtValue as variant
    if cb is nothing then
      Call CopyVariant(vtValue,v)
    else
      Call CopyVariant(vtValue,cb.Run(v))
    end if

    'Compare values and return 
    if isEmpty(vRet) then
      Call CopyVariant(vRet,v)
      Call CopyVariant(vMaxValue, vtValue)
    elseif vMaxValue < vtValue then
      Call CopyVariant(vRet,v)
      Call CopyVariant(vMaxValue, vtValue)
    end if
  next

  Call CopyVariant(Max,vRet)
End Function

'Obtain the minimum value in the array
'@param cb as stdICallable<(element: Variant)=>Variant> - Callback to run on each element. The return value of this callback is used to determine the minimum value.
'@param startingValue - The starting value to compare against. If not provided, the first element of the array is used.
'@returns - The minimum value in the array
Public Function Min(Optional ByVal cb As stdICallable = nothing, Optional ByVal startingValue As Variant = Empty) as variant
  Dim vRet, vMinValue, v
  vMinValue = startingValue: vRet = startingValue
  For i = 1 to pLength
    Call CopyVariant(v,pArr(i))
    
    'Get value to test
    Dim vtValue as variant
    if cb is nothing then
      Call CopyVariant(vtValue,v)
    else
      Call CopyVariant(vtValue,cb.Run(v))
    end if

    'Compare values and return 
    if isEmpty(vRet) then
      Call CopyVariant(vRet,v)
      Call CopyVariant(vMinValue, vtValue)
    elseif vMinValue > vtValue then
      Call CopyVariant(vRet,v)
      Call CopyVariant(vMinValue, vtValue)
    end if
  next

  Call CopyVariant(Min,vRet)
End Function

'Copies one variant to a destination
'@param {ByRef Variant} dest Destination to copy variant to
'@param {Variant} value Source to copy variant from.
'@perf This appears to be a faster variant of "oleaut32.dll\VariantCopy" + it's multi-platform
Private Sub CopyVariant(ByRef dest As Variant, ByVal value As Variant)
  If IsObject(value) Then
    Set dest = value
  Else
    dest = value
  End If
End Sub
