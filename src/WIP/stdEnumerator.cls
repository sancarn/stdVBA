VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdEnumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Spec:
'For some enumerable set ["foo","bar","baz"] let stdEnumerator implement functions:
'* next
'* each
'* map
'* with_index
'...
'Can stack each call as follows:
'map.with_index

'CreateFromForEach
'CreateFromICallable
'Sort() as stdArray
'Reverse() as stdArray
'Concat()
'Join
'indexOf
'lastIndexOf
'includes
'count
'isEvery
'isSome
'ForEach
'Map
'Unique
'Reduce
'Filter
'Count_By
'groupBy

'RUBY  https://ruby-doc.org/core-2.7.1/Enumerable.html#method-i-cycle
'checkAll
'checkAny
'checkNone
'checkOnlyOne

'chain
'collect
'collect_concat
'count
'cycle
'detect
'drop(n)           [1,2,3,4,5].drop(2) ==> [3,4,5]
'drop_while(cb)
'each_cons  [1,2,3,4,5].each_cons(2,cb) ==> cb([1,2]) ==> cb([2,3]) ==> cb([3,4]) ==> cb([4,5])
'each_slice [1,2,3,4,5].each_slice(2,cb) ==> cb([1,2]) ==> cb([3,4]) ==> cb([5])
'each_with_index [1,3,5].each_with_index(cb)  ==> cb(1, 1) ==> cb(2,3) ==> cb(3,5)
'filter/select
'find_first
'flatten(cb?)
'max(cb)
'min(cb)
'sum(cb)
'partition   [1,2,3,4,5,6].partition(a=>a%2=0) ==> [[2,4,6],[1,3,5]]
'zip         [1,2,3].zip([4,5,6]) ==> [[1,4],[2,5],[3,6]]          |            [1,2,3].zip([1,2]) ==> [[1,1],[2,2],[3,null]]

'It's difficult to think of uses for zip, so here are some:
'a = stdEnumerator.CreateFromArray(split("a b c"," ")).zip([1,2,3]).to_dict() ==> {a:1, b:2, c:3}
'vector addition and multiplication:
'[1,2].zip([2,3]).map(e=>e[0]+e[1]) ==> [3,5]
'[1,2].zip([2,3]).map(e=>e[0]*e[1]) ==> [2,6]

'Things we can't do (yet)
'take <-- can't do this unless we implement IEnumVARIANT and call Next() method
'tally <-- Would like to do this but can't until we have stdDictionary     ["a","b","c","b","a","b"].tally ==> {a:2, b:3, c:1}
'to_dict <-- requires stdDictioanry                                        ["a",1,"b",2].to_dict ==> {a:1, b:2}


Private Enum EnumeratorType
    FromCallback
    FromIEnumVariant
End Enum

Private Type SortStruct
    value as variant
    sortValue as variant
    iIndex as long
    iNext as long 
    iPrev as long
End Type

Private mode As EnumeratorType
Private pEnumObject as object

Private Iteration As Long
Private lowerBound As Long

Public Function Sort(cb as callback, ctor as object) as stdArray 'stdArray? Or can i use stdEnum again? and have a `stdArray.Create(stdEnum)`?
    Dim arr() as SortStruct
    Redim arr(1 to me.count)
    Dim iIndex as Long: iIndex = 0
    Dim iFirstItem as long: iFirstItem = 1
    Dim val as variant
    For each val in pEnumObject
        'Increment index
        iIndex = iIndex + 1

        'Bind to SortStruct
        if iIndex > 1 Then
            'Initialise sorting struct
            Call CopyVariant(arr(iIndex).value, val)
            arr(iIndex).iIndex = iIndex
            arr(iIndex).sortValue = cb.Run(arr(iIndex).value)

            'Sort/Compare
            Dim iCompareIndex as long: iCompareIndex = iFirstItem

            Do While iCompareIndex <> 0
                'If sort value at current index is less than at compare index then but this index to compare index via next
                if arr(iIndex).sortValue < arr(iCompareIndex).sortValue then
                    'Bind this index to compare index via iNext property
                    arr(iIndex).iNext = arr(iCompareIndex).iIndex

                    'Rebind previous element if required
                    if arr(iCompareIndex).iPrev <> 0 then
                        'There is a previous element, rebind it!
                        arr(arr(iCompareIndex).iPrev).iNext = iIndex
                    else
                        'There is no previous element i.e. this is the first element, change iFirstItem
                        iFirstItem = iIndex
                    End if
                    
                    'No need to carry on searching for where item should go, exit do loop
                    Exit Do
                Else
                    'Ensure next element defined, if not then we have a new next element
                    if arr(iCompareIndex).iNext <> 0 then
                        'Schedule next sorting check and keep searching
                        iCompareIndex = arr(iCompareIndex).iNext
                    else
                        'Next element is not defined, therefore this is max
                        'in this case set next of arr(iCompareIndex) to this
                        'set prev of this to iCompareIndex
                        arr(iCompareIndex).iNext = iIndex
                        arr(iIndex).iPrev = iCompareIndex
                        
                        'No need to carry on searching for where item should go, exit do loop
                        Exit Do
                    end if
                end if
            Loop
        Else
            'Initialise sorting struct
            Call CopyVariant(arr(1).value, val)
            arr(1).sortValue = cb.Run(arr(1).value)
            arr(1).iIndex = 1
            arr(1).iNext = 0
            arr(1).iPrev = 0
            
        end if
    next

    'Collect sorted elements
    Dim ret as stdArray
    Dim i as long: i = iFirstItem
    While i <> 0
        Call ret.push(arr(i).value)
        i = arr(i).iNext
    Wend

    'Return sorted array
    set sort = ret
End Function





Public Function CreateFromIEnumVariant(ByRef o as object)
    set pEnumObject = o
End Function
'CreateFromICallable

Public Function FeedItem(v As Variant)
    Call oRetArray.push(v)
End Function

'Incompatible with oEnumVariant
Public Function NextItem() As Variant
    mode = iMode
    If mode = FromForEach Then
        
    ElseIf mode = FromCallback Then
        
    ElseIf mode = FromIEnumVariant Then
        Set NextItem = oEnumVariant.Next()
    End If
End Function
Public Function NextItems() As stdArray
    mode = iMode
    Set NextItems = stdArray.Create()
    If mode = FromForEach Then
        
    ElseIf mode = FromCallback Then
        
    ElseIf mode = FromIEnumVariant Then
        Dim x As IEnumVARIANT
        Set x = oEnumVariant.Clone
        
        Z = x.Next
        While Z
            Call NextItems.push(Z)
            Z = x.Next
        Wend
    End If
End Function
Public Function PeekItem() As Variant
    If mode = FromForEach Then
        
    ElseIf mode = FromCallback Then
        
    ElseIf mode = FromIEnumVariant Then
        Dim x As IEnumVARIANT
        Set x = oEnumVariant.Clone
        Set PeekItem = x.Next
    End If
End Function
Public Function PeekItems() As stdArray
    
End Function
Public Function Rewind() As stdEnumerator
    
End Function
Public Function Size() As Long

End Function




'Public Static stdEnumerator Create(IDispatch obj, string method, *args)
Public Function Create(ByVal cb As stdCallback) As stdEnumerator
    Set Create = New stdEnumerator
    Call Create.Init(FromForEach, cb)
End Function

'Implements each like:
'  For i = 1 to size: cb(size,metadata...): next
Public Function CreateFromCallback(ByVal Size As Long, ByVal cb As stdCallback) As stdEnumerator
    Set FromCallback = New stdEnumerator
    Call CreateFromCallback.Init(FromCallback, Size, cb)
End Function

Public Function CreateFromIEnumVariant(ByVal obj As Object) As stdEnumerator
    Set CreateFromIEnumVariant = New stdEnumerator
    Call CreateFromIEnumVariant.Init(FromIEnumVariant, obj)
    
End Function


Public Sub Init(iMode As BiState, ParamArray v() As Variant)
    mode = iMode
    If mode = FromForEach Then
        
    ElseIf mode = FromCallback Then
        
    ElseIf mode = FromIEnumVariant Then
        Set oEnumVariant = v(0)
    End If
End Sub



Public Function GetArray() As stdArray
    Set GetArray = oRetArray
End Function
Public Sub ForEach(Optional ByVal cb As stdCallback)
    Dim v As Variant
    If mode = FromForEach Then
        
    ElseIf mode = FromCallback Then
        
    ElseIf mode = FromIEnumVariant Then
        For Each v In oEnumVariant
            Call cb(v)
        Next
    End If
End Sub
Public Function Map(Optional ByVal cb As stdCallback) As stdArray
    Dim v As Variant
    If mode = FromForEach Then
        
    ElseIf mode = FromCallback Then
        
    ElseIf mode = FromIEnumVariant Then
        For Each v In oEnumVariant
            oRetArray.push (cb(v))
        Next v
    End If
End Function
Public Sub WithIndex(Optional ByVal cb As stdCallback)
    If mode = FromForEach Then
        
    ElseIf mode = FromCallback Then
        
    ElseIf mode = FromIEnumVariant Then
        
    End If
End Sub
Public Sub EachWithIndex(Optional ByVal cb As stdCallback)
    If mode = FromForEach Then
        
    ElseIf mode = FromCallback Then
        
    ElseIf mode = FromIEnumVariant Then
        
    End If
End Sub
Public Sub WithObject(Optional ByVal cb As stdCallback)
    Call EachWithObject(cb)
End Sub
Public Sub EachWithObject(Optional ByVal cb As stdCallback)
    If mode = FromForEach Then
        
    ElseIf mode = FromCallback Then
        
    ElseIf mode = FromIEnumVariant Then
        
    End If
End Sub



