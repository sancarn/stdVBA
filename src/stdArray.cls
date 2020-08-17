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

#If VBA6 Then
  Private Declare PtrSafe Sub VariantCopy Lib "oleaut32.dll" (ByRef pvargDest As Variant, ByRef pvargSrc As Variant)
  Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As Long)
#Else
  Private Declare Sub VariantCopy Lib "oleaut32.dll" (ByRef pvargDest As Variant, ByRef pvargSrc As Variant)
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
#End If
Private Enum SortDirection
    Ascending = 1
    Descending = 2
End Enum


Private pArr() As Variant
Private pProxyLength As Long
Private pLength As Long

Private pChunking As Long
Private pInitialised As Boolean




Public Function Create(ParamArray params() As Variant) As stdArray
    Set Create = New stdArray
    
    Dim i As Long
    Dim lb As Long: lb = LBound(params)
    Dim ub As Long: ub = UBound(params)
    
    Call Create.init(ub - lb + 1, 10)
    
    For i = lb To ub
        Call Create.Push(params(i))
    Next
End Function


Public Function CreateWithOptions(ByVal iInitialLength As Long, ByVal iChunking As Long, ParamArray params() As Variant) As stdArray
    Set CreateWithOptions = New stdArray
    
    Dim i As Long
    Dim lb As Long: lb = LBound(params)
    Dim ub As Long: ub = UBound(params)
    
    Call CreateWithOptions.init(iInitialLength, iChunking)
    For i = lb To ub
        Call CreateWithOptions.Push(params(i))
    Next
End Function


Public Function CreateFromArray(ByVal arr As Variant) As stdArray
    Set CreateFromArray = New stdArray
    
    Dim i As Long
    Dim lb As Long: lb = LBound(arr)
    Dim ub As Long: ub = UBound(arr)
    Call CreateFromArray.init(ub - lb + 1, 10)
    
    For i = lb To ub
        Call CreateFromArray.Push(arr(i))
    Next
End Function


Friend Sub init(ByVal iInitialLength As Long, ByVal iChunking As Long)
  If iChunking > iInitialLength Then iInitialLength = iChunking
  If Not pInitialised Then
    pProxyLength = iInitialLength
    ReDim pArr(1 To iInitialLength) As Variant
    pChunking = iChunking
    pInitialised = True
  End If
End Sub

Public Property Get Length() As Long
    Length = pLength
End Property
Public Property Get zProxyLength() As Long
    zProxyLength = pProxyLength
End Property

Public Sub Resize(ByVal iLength As Long)
  pLength = iLength
End Sub

Public Sub Rechunk()
  Dim fNumChunks As Double, iNumChunks As Long
  fNumChunks = pLength / pChunking
  iNumChunks = CLng(fNumChunks)
  If fNumChunks > iNumChunks Then iNumChunks = iNumChunks + 1
  
  ReDim Preserve pArr(1 To iNumChunks * pChunking) As Variant
End Sub






Public Property Get arr() As Variant
    Dim vRet() As Variant
    ReDim vRet(1 To pLength) As Variant
    CopyMemory VarPtr(vRet(1)), VarPtr(pArr(1)), pLength * 16 '16 is variant size
    arr = vRet
End Property
Public Property Let arr(v As Variant)
    Dim lb As Long: lb = LBound(v)
    Dim ub As Long: ub = UBound(v)
    Dim cnt As Long: cnt = ub - lb + 1
    ReDim pArr(1 To (Int(cnt / pChunking) + 1) * pChunking) As Variant
    For i = lb To ub
        Call Push(pArr(i))
    Next
End Property

Public Sub Push(ByVal el As Variant)
  If pInitialised Then
    If pLength = pProxyLength Then
        pProxyLength = pProxyLength + pChunking
        ReDim Preserve pArr(1 To pProxyLength) As Variant
    End If
    
    pLength = pLength + 1
    VariantCopy pArr(pLength), el
  Else
    'Error
  End If
End Sub

Public Function Pop() As Variant
    If pInitialised Then
        If pLength > 0 Then
            VariantCopy Pop, pArr(pLength)
            pLength = pLength - 1
        Else
            Pop = Empty
        End If
    Else
        'Error
    End If
End Function

Public Function Remove(ByVal index As Long) As Variant
  'Ensure initialised
  If pInitialised Then
    'Ensure length > 0
    If pLength > 0 Then
      'Ensure index < length
      If index <= pLength Then
        'Copy party we are removing to return variable
        VariantCopy Remove, pArr(index)
        
        'Loop through array from removal, set i-1th element to ith element
        Dim i As Long
        For i = index + 1 To pLength
            pArr(i - 1) = pArr(i)
        Next

        'Set last element length and subtract total length by 1
        pArr(pLength) = Empty
        pLength = pLength - 1
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

Public Function Shift() As Variant
  'Would be good to use CopyMemory here
  
  VariantCopy Shift, pArr(1)
  Dim i As Long
  For i = 1 To pLength - 1
    pArr(i) = pArr(i + 1)
  Next
End Function

Public Function Unshift(val As Variant) As stdArray
  'Would be good to use CopyMemory here
  
  'Ensure array is big enough and increase pLength
  Call Push(0)
  
  'Unshift
  For i = pLength - 1 To 1 Step -1
    pArr(i + 1) = pArr(i)
  Next
  pArr(0) = val
  
  Set Unshift = Me
End Function

Public Function Slice() As stdArray
  
End Function

Public Function Splice() As stdArray
  
End Function

Public Function Clone() As stdArray
  If pInitialised Then
      If pInitialised Then
        'Similar to CreateFromArray() but passing length through also:
        Set Clone = New stdArray
        
        Call Clone.init(pLength, 10)
        
        Dim i As Long
        For i = 1 To pLength
            Call Clone.Push(pArr(i))
        Next
      Else
        'Error
      End If
  Else
    'Error
  End If
End Function

Public Function Reverse() As stdArray
    Dim ret As stdArray
    Set ret = stdArray.CreateWithOptions(pLength, pChunking)
    For i = pLength To 1 Step -1
      Call ret.PutItem(i - pLength + 1, pArr(i))
    Next
    Set Reverse = ret
End Function

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

Public Function Join(Optional ByVal delimeter As String = ",") As String
  If pInitialised Then
      If pLength > 0 Then
        Dim sOutput As String
        sOutput = pArr(1)
        
        Dim i As Long
        For i = 2 To pLength
          sOutput = sOutput & delimeter & pArr(i)
        Next
        Join = sOutput
      Else
        Join = ""
      End If
  Else
    'Error
  End If
End Function


Public Property Get item(ByVal i As Integer) As Variant
    'item(1) = 1st element
    'item(2) = 2nd element
    'etc.
    VariantCopy item, pArr(i)
    
End Property
Public Property Set item(ByVal i As Integer, item As Object)
  VariantCopy pArr(i), item
End Property
Public Property Let item(ByVal i As Integer, item As Variant)
  VariantCopy pArr(i), item
End Property
Public Sub PutItem(ByVal index As Integer, ByRef item As Variant)
  VariantCopy pArr(index), item
End Sub


Public Function indexOf(ByVal el As Variant, Optional ByVal start As Integer = 1)
  Dim elIsObj As Boolean, i As Long, item As Variant, itemIsObj As Boolean
  
  'Is element an object?
  elIsObj = IsObject(el)
  
  'Loop over contents starting from start
  For i = start To pLength
    'Get item data
    VariantCopy item, pArr(i)
    
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

Public Function lastIndexOf(ByVal el As Variant, Optional ByVal start As Integer)
  Dim elIsObj As Boolean, i As Long, item As Variant, itemIsObj As Boolean
  
  'Is element an object?
  elIsObj = IsObject(el)
  
  'Loop over contents starting from start
  For i = pLength To start Step -1
    'Get item data
    VariantCopy item, pArr(i)
    
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

Public Function includes(ByVal el As Variant, Optional ByVal startFrom As Integer = 1) As Boolean
  includes = indexOf(el, startFrom) >= startFrom
End Function

Public Function Count(Optional ByVal v As Variant = Empty, Optional ByVal metadata As Object = Nothing) As Long
    Dim i As Long, lCount As Long
    lCount = 0
    
    If IsObject(v) Then
      For i = 1 To pLength
        If pArr(i) Is v Then
          lCount = lCount + 1
        End If
      Next
    Else
      For i = 1 To pLength
        If pArr(i) = v Then
          lCount = lCount + 1
        End If
      Next
    End If

    'Return count
    Count = lCount
End Function





'Iterative Functions (All require stdICallable):

'Example: if incidents.IsEvery(cbValid) then ...
Public Function IsEvery(ByVal cb As stdICallable, Optional ByVal metadata As Object = Nothing) As Boolean
  If pInitialised Then
    Dim i As Long
    For i = 1 To pLength
      If Not IIf(metadata Is Nothing, cb(pArr(i)), cb(metadata, pArr(i))) Then
        IsEvery = False
        Exit Function
      End If
    Next
    
    IsEvery = True
  Else
    'Error
  End If
End Function

Public Function IsSome(ByVal cb As stdICallable, Optional ByVal metadata As Object = Nothing) As Boolean
  If pInitialised Then
    Dim i As Integer
    For i = LBound(pArr) To UBound(pArr)
      If IIf(metadata Is Nothing, cb(pArr(i)), cb(metadata, pArr(i))) Then
        IsSome = True
        Exit Function
      End If
    Next
    IsSome = False
  Else
    'Error
  End If
End Function

Public Sub ForEach(ByVal cb As stdICallable, Optional ByVal metadata As Object = Nothing)
  If pInitialised Then
    Dim i As Integer
    For i = 1 To pLength
      If metadata Is Nothing Then
          Call cb.Run(pArr(i))
      Else
          Call cb.Run(metadata, pArr(i))
      End If
    Next
  Else
    'Error
  End If
End Sub

Public Function Map(ByVal cb As stdICallable, Optional ByVal metadata As Object = Nothing) As stdArray
  If pInitialised Then
    Dim pMap As stdArray
    Set pMap = Me.Clone()
    
    Dim i As Integer
    For i = 1 To pLength
      'BUGFIX: Sometimes required, not sure when
      Dim v As Variant
      VariantCopy v, item(i)
      
      'Call callback based on metadata
      If metadata Is Nothing Then
        Call pMap.PutItem(i, cb.Run(v))
      Else
        Call pMap.PutItem(i, cb.Run(metadata, v))
      End If
    Next
    
    Set Map = pMap
  Else
    'Error
  End If
End Function


'OPTIMISE: Needs optimisation. Currently very sub-optimal
Public Function Unique(Optional ByVal cb As stdICallable = Nothing, Optional ByVal metadata As Variant) As stdArray
  Dim ret As stdArray
  Set ret = stdArray.CreateWithOptions(pLength, pChunking)
  
  Dim vKeys As stdArray
  If cb Is Nothing Then
    Set vKeys = Me.Clone
  Else
    If metadata Is Nothing Then
      Set vKeys = Me.Map(cb)
    Else
      Set vKeys = Me.Map(cb, metadata)
    End If
  End If
  
  
  For i = 1 To pLength
    If Not ret.includes(vKeys(i)) Then
      Call ret.Push(vKeys(i))
    End If
  Next

  'Return data
  Set Unique = ret
End Function


Public Function Reduce(ByVal cb As stdICallable, Optional ByVal initialValue As Variant=0, Optional ByVal metadata As Variant) As Variant
  If pInitialised Then
    Reduce = initialValue
    
    Dim i As Integer
    For i = 1 To pLength
      'BUGFIX: Sometimes required, not sure when
      Dim el As Variant
      VariantCopy el, pArr(i)

      'Reduce
      If isMissing(metadata) Then
        Reduce = cb(Reduce, el)
      Else
        Reduce = cb(metadata, Reduce, el)
      End If
    Next
  Else
    'Error
  End If
End Function

Public Function Filter(ByVal cb As stdICallable, Optional ByVal metadata As Object = Nothing) As stdArray
    Dim ret As stdArray
    Set ret = stdArray.CreateWithOptions(pLength, pChunking)
    Set Filter = ret

    'If initialised...
    If pInitialised Then
        Dim i As Long, v As Variant
        'Loop over array
        For i = 1 To pLength
          
          
            'If callback succeeds (include metadata if given)
            If metadata Is Nothing Then
                If cb(pArr(i)) Then
                    Call ret.Push(pArr(i))
                End If
            Else
                If cb(metadata, pArr(i)) Then
                    Call ret.Push(pArr(i))
                End If
            End If
        Next i
    Else
        'error
    End If
End Function

Public Function Count_By(cb As stdICallable, Optional ByVal metadata As Object = Nothing) As Long
  Dim i As Long, lCount As Long
  lCount = 0
  For i = 1 To pLength
    If metadata Is Nothing Then
      If cb(pArr(i)) Then
        lCount = lCount + 1
      End If
    Else
      If cb(metadata, pArr(i)) Then
        lCount = lCount + 1
      End If
    End If
  Next i
  Count_By = lCount
End Function

Public Function groupBy(ByVal cb As stdICallable, Optional ByVal metadata As Object) As Object
    'Array to store result in
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    'Loop over items
    Dim i As Long
    For i = 1 To pLength
        'Get grouping key
        Dim key As Variant
        If metadata Is Nothing Then
          key = cb(pArr(i))
        Else
          key = cb(metadata, pArr(i))
        End If

        'If key is not set then set it
        If Not result.exists(key) Then Set result(key) = stdArray.Create()

        'Push item to key
        result(key).Push pArr(i)
    Next

    'Return result
    Set groupBy = result
End Function





Public Sub Test()
  'New stdArray tests
  Dim i As Long
  Dim x As stdArray
  Set x = stdArray.Create()
  For i = 1 To 10
      Call x.Push(i)
  Next
  Debug.Assert x.zProxyLength = 10
  Call x.Push(11)
  Debug.Assert x.zProxyLength = 20
  
  Dim y As stdArray
  Set y = stdArray.CreateWithOptions(5, 11)
  For i = 1 To 11
      Call y.Push(i)
  Next
  Debug.Assert y.zProxyLength = 11
  Call y.Push(12)
  Debug.Assert y.zProxyLength = 22
  
  
  y.item(1) = 10
  Debug.Assert False
  
  y.PutItem 1, 30
  
  
  Dim z As stdArray
  Set z = stdArray.CreateWithOptions(15, 10)
  
  For i = 1 To 16
    Call z.Push(i)
  Next
  Debug.Assert z.Length = 16
  Debug.Assert z.zProxyLength = 25
    
    
    
    
    'ORIGINAL stdArray tests
    On Error Resume Next
    Debug.Print "1-Generic tests:"
    Dim v As stdArray:
    Set v = stdArray.Create(1)
    Debug.Print "1-1:" & (v.item(1) = 1)
    Call v.Push(2)
    Debug.Print "1-2:" & (v.item(1) = 1)
    Debug.Print "1-3:" & (v.item(2) = 2)
    
    'Test object pushing and access
    Dim o As Object
    Set o = CreateObject("Scripting.Dictionary")
    o("hello") = "world"
    Call v.Push(o)
    Debug.Print "1-4:" & (v.item(3)("hello") = "world")
    
    Debug.Print "2-Length tests:"
    Debug.Print "2-1:" & (stdArray.Create().Length = 0)
    Debug.Print "2-2:" & (stdArray.Create(1).Length = 1)
    Debug.Print "2-3:" & (stdArray.Create(1, 2).Length = 2)
    
    
    'If stdICallable Then
    '  Debug.Print "9-stdICallable tests"
    'End If
    
    
End Sub
