Attribute VB_Name = "stdArrayTests"
'@lang VBA

Sub testAll()
    Test.Topic "stdArray"

    Dim tmp as Variant 'used by various


    Dim arr as stdArray
    set arr = stdArray.Create(1,2,3)
    Test.Assert "Check array exists", not arr is nothing
    Test.Assert "Check item 1", arr.item(1) = 1
    Test.Assert "Check item 2", arr.item(2) = 2
    Test.Assert "Check item 3", arr.item(3) = 3

    Dim vIter, iCount as long: iCount = 0
    For each vIter in arr
        iCount=iCount+1
        Test.Assert "Check item is number", isNumeric(vIter)
    next
    Test.Assert "Check loop triggered", iCount = 3

    set arr = stdArray.CreateFromArray(Array(1,2,3))
    Test.Assert "Check CreateFromArray 1", arr.item(1) = 1
    Test.Assert "Check CreateFromArray 2", arr.item(2) = 2
    Test.Assert "Check CreateFromArray 3", arr.item(3) = 3

    Test.Assert "Length", arr.Length = 3
    Call arr.Resize(4): Test.Assert "Resize 1 increase", arr.length = 4
    Call arr.resize(3): Test.Assert "Resize 2 decrease", arr.length = 3

    Call arr.push(4): Test.Assert "Push 1", arr.item(4) = 4
    Test.Assert "Ensure array length", arr.length = 4
    Test.Assert "Pop 1 returns value", arr.pop() = 4
    Test.Assert "Pop 1 changes array length", arr.length = 3

    Test.Assert "Shift 1", arr.shift() = 1
    Test.Assert "Shift 2", arr.length = 2
    'debug.assert false
    Test.Assert "Unshift 1", TypeOf arr.unshift(1) is stdArray
    Test.Assert "Unshift 2", arr.item(1) = 1

    'Test property arr
    Test.Assert "Property Arr 1 always returns array when empty", TypeName(stdArray.Create().arr) = "Variant()"
    Dim vArr as variant
    vArr = arr.arr
    Test.Assert "Property Arr 2 typename", TypeName(vArr) = "Variant()"
    if TypeName(vArr) = "Variant()" then
        Test.Assert "Property Arr 3 lbound", lbound(vArr) = 1
        Test.Assert "Property Arr 4 item 1 equal", vArr(1) = arr.item(1)
        Test.Assert "Property Arr 4 item 2 equal", vArr(2) = arr.item(2)
        Test.Assert "Property Arr 4 item 3 equal", vArr(3) = arr.item(3)
    end if

    Call testDefaultBase
    Call testArrLet
    
    'Remove
    set tmp = stdArray.Create(1,2,3)
    Test.Assert "Remove 1 Item returned", tmp.Remove(2) = 2
    Test.Assert "Remove 2 Item removed", tmp.join = "1,3"

    'Slice
    'TODO: UNIMPLEMENTED

    'Splice
    'TODO: UNIMPLEMENTED

    set tmp = arr.clone: tmp.Item(1) = "x"

    'Item [let]
    Test.Assert "item [let]", tmp.join() = "x,2,3"

    'Clone
    Test.Assert "Clone, clones data not instance", (arr.join = "1,2,3") AND (tmp.join() = "x,2,3")

    'Reverse
    Test.Assert "Reverse of array", arr.Reverse.Join() = "3,2,1"

    'Concat
    Test.Assert "Concat works as expected", arr.concat(stdArray.Create(4,5,6)).Join() = "1,2,3,4,5,6"
    Test.Assert "Concat doesn't alter array", arr.Join() = "1,2,3"

    'Testing join - ease of future tests
    Test.assert "Join 1 default seperator", arr.join() = "1,2,3"
    Test.Assert "Join 2 with seperator", arr.join(";") = "1;2;3"

    'item set
    set tmp2 = stdArray.Create()
    set tmp.item(1) = tmp2
    Test.Assert "item [set]", tmp.item(1) is tmp2

    'PutItem
    'TODO: Not sure what we need to test explicitely here. I.E. Why use PutItem in the first place?

    'indexOf
    Test.Assert "indexOf 1 If can find ret index", stdArray.Create(1,2,3,2,3,4).indexOf(3) = 3
    Test.Assert "indexOf 2 If cannot find ret -1", stdArray.Create(1,2,3,2,3,4).indexOf(5) = -1

    'lastIndexOf
    Test.Assert "lastIndexOf 1 If can find ret index", stdArray.Create(1,2,3,2,3,4).lastIndexOf(3) = 5
    Test.Assert "lastIndexOf 2 If cannot find ret -1", stdArray.Create(1,2,3,2,3,4).lastIndexOf(5) = -1

    'includes
    Test.Assert "Includes 1", arr.includes(1)
    Test.Assert "Includes 2", not arr.includes(99)

    'Used in Unique and Group
    Dim dict as object: set dict = CD(1,"A",2,"A",3,"B",4,"C",5,"C")
    Dim lookup as stdCallback: set lookup = stdCallback.CreateFromObjectProperty(dict,"Item", vbGet)

    set arr = stdArray.Create(1,4,5,2,3)
    'CALLBACKS, using stdLambda:
    'IsEvery(cb)
    Test.Assert "IsEvery 1 Correct", arr.IsEvery(stdLambda.Create("$1<=5"))
    Test.Assert "IsEvery 2 Incorrect", Not arr.IsEvery(stdLambda.Create("$1<=3"))

    'IsSome(cb)
    Test.Assert "IsSome 1", arr.IsSome(stdLambda.Create("$1<=3"))
    Test.Assert "IsSome 2", Not arr.IsSome(stdLambda.Create("$1<=0"))
    
    'ForEach(cb)
    set tmp = stdArray.Create(0)
    Call arr.ForEach(stdLambda.Create("$1.push($2)").Bind(tmp))
    Test.Assert "ForEach", tmp.join() = "0,1,4,5,2,3"

    'Map(cb)
    Test.Assert "Map", arr.Map(stdLambda.Create("$1*2")).Join() = "2,8,10,4,6"

    'Unique(cb?)
    Test.Assert "Unique no callback", stdArray.Create(1,1,1,2,3,2,4,5,5).Unique().join() = "1,2,3,4,5"
    Test.Assert "Unique w/ callback", stdArray.Create(1,2,3,4,5).Unique(lookup).Join() = "1,3,4"

    'Reduce(cb,initialValue?, metadata?)
    Test.Assert "Reduce", arr.Reduce(stdLambda.Create("$1+$2"),0)
    Test.Assert "1st Arg is accumulator", stdArray.Create(False).Reduce(stdLambda.Create("$1"),TRUE)
    Test.Assert "2nd Arg is not accumulator", Not stdArray.Create(False).Reduce(stdLambda.Create("$2"),TRUE)
    
    'Filter(cb)
    Test.Assert "Filter", arr.Filter(stdLambda.Create("$1>=3")).join() = "4,5,3"

    'Count(cb?)
    Test.Assert "Count no cb == Length", arr.Count = arr.Length
    Test.Assert "Count w/ cb", arr.Count(stdLambda.Create("$1<=3")) = 3
    
    'Group(cb)
    set tmp = arr.groupBy(lookup)
    Test.Assert "Group 1", typename(tmp("A")) = "stdArray"
    Test.Assert "Group 2", tmp("A").join() = "1,2"
    Test.Assert "Group 3", tmp("B").join() = "3"
    Test.Assert "Group 4", tmp("C").join() = "4,5"
    
    'Sort(cbSortBy?,cbComparrason?,iAlgorithm?,bSortInPlace?)
    Test.Assert "Sort no lambda", arr.sort().join() = "1,2,3,4,5"
    Test.Assert "Sort w/ lambda", arr.sort(stdLambda.Create("1/$1")).join = "5,4,3,2,1"

    'Min(cbMinBy?) and Max(cbMaxBy?) tests
    Test.Assert "Max no lambda", arr.max() = 5
    Test.Assert "Min no lambda", arr.min() = 1
    Test.Assert "Max w/ lambda", arr.max(stdLambda.Create("1/$1")) = 1
    Test.Assert "Min w/ lambda", arr.min(stdLambda.Create("1/$1")) = 5
End Sub

Public Sub testArrLet()
    Dim arr As stdArray
    Dim oneBased(1 To 2) As Variant
    oneBased(1) = 9
    oneBased(2) = 10
    Set arr = stdArray.Create(7, 8)
    arr.arr = oneBased
    Test.Assert "Arr Let replaces length", arr.Length = 2
    Test.Assert "Arr Let copies one-based values", arr.Join() = "9,10"

    Dim exported As Variant
    exported = arr.arr
    Test.Assert "Arr Let export lower bound", LBound(exported) = 1
    Test.Assert "Arr Let export upper bound", UBound(exported) = 2
    Test.Assert "Arr Let export values", exported(1) = 9 And exported(2) = 10

    Dim zeroBased(0 To 1) As Variant
    zeroBased(0) = 11
    zeroBased(1) = 12
    arr.arr = zeroBased
    Test.Assert "Arr Let copies zero-based values", arr.Join() = "11,12"
    Test.Assert "Arr Let repeated assignment resets length", arr.Length = 2

    Dim negativeBased(-3 To -2) As Long
    negativeBased(-3) = 13
    negativeBased(-2) = 14
    arr.arr = negativeBased
    Test.Assert "Arr Let copies typed arrays with negative bounds", arr.Join() = "13,14"

    'Replacing a large allocation with a smaller one must also update capacity.
    Set arr = stdArray.CreateWithOptions(40, 4, 1, 7, 8)
    arr.arr = oneBased
    Dim capacity As Long: capacity = arr.zProxyLength
    Dim i As Long
    For i = 3 To capacity + 1
        Call arr.Push(i)
    Next
    Test.Assert "Arr Let smaller replacement can grow", arr.Length = capacity + 1
    Test.Assert "Arr Let smaller replacement preserves values", arr.Item(1) = 9 And arr.Item(2) = 10
    Test.Assert "Arr Let smaller replacement appends values", arr.Item(capacity + 1) = capacity + 1

    'Replacing with more items must permit growth past the new allocation.
    Dim longer(1 To 12) As Variant
    For i = 1 To 12
        longer(i) = i
    Next
    Set arr = stdArray.Create(7, 8)
    arr.arr = longer
    Test.Assert "Arr Let longer replacement length", arr.Length = 12
    Test.Assert "Arr Let longer replacement values", arr.Join() = "1,2,3,4,5,6,7,8,9,10,11,12"
    capacity = arr.zProxyLength
    For i = 13 To capacity + 1
        Call arr.Push(i)
    Next
    Test.Assert "Arr Let longer replacement can grow", arr.Length = capacity + 1
    Test.Assert "Arr Let longer replacement appends values", arr.Item(capacity + 1) = capacity + 1

    arr.arr = Array()
    Test.Assert "Arr Let empty array clears length", arr.Length = 0
    exported = arr.arr
    Test.Assert "Arr Let empty array exports empty array", UBound(exported) < LBound(exported)
    Call arr.Push(15)
    Test.Assert "Arr Let empty replacement can append", arr.Length = 1 And arr.Item(1) = 15

    Dim obj As Collection: Set obj = New Collection
    Dim mixed(0 To 4) As Variant
    Set mixed(0) = obj
    mixed(2) = Null
    Set mixed(3) = Nothing
    mixed(4) = Array(3, 4)
    arr.arr = mixed
    Test.Assert "Arr Let preserves object identity", arr.Item(1) Is obj
    Test.Assert "Arr Let preserves Empty", IsEmpty(arr.Item(2))
    Test.Assert "Arr Let preserves Null", IsNull(arr.Item(3))
    Test.Assert "Arr Let preserves Nothing", arr.Item(4) Is Nothing
    exported = arr.Item(5)
    Test.Assert "Arr Let preserves nested arrays", exported(0) = 3 And exported(1) = 4
    Test.Assert "Arr Let leaves source unchanged", mixed(0) Is obj
End Sub


Private Function CD(Paramarray v() as variant) as object
    set o = CreateObject("Scripting.Dictionary")
    For i = lbound(v) to ubound(v) step 2
        if isObject(v(i+1)) then
            set o(v(i)) = v(i+1)
        else
            o(v(i)) = v(i+1)
        end if
    next
    Set CD = o
End Function

Private Sub testDefaultBase()
  Dim originalBase As Long: originalBase = stdArray.DefaultBase
  Dim arr As stdArray, derived As stdArray, values As Variant, chunks As stdArray
  Dim failure As Long, description As String
  On Error GoTo Failed
  Test.Assert "Default base starts at one", originalBase = 1
  stdArray.DefaultBase = 0
  Set arr = stdArray.Create(10, 20, 30)
  Test.Assert "Default base applies to Create", arr.Base = 0 And arr(0) = 10 And arr(2) = 30
  stdArray.DefaultBase = 5
  Test.Assert "Existing instance retains base", arr.Base = 0 And arr(0) = 10
  Set derived = stdArray.CreateFromArray(Array(1, 2))
  Test.Assert "CreateFromArray uses default", derived.Base = 5 And derived(5) = 1
  Set derived = stdArray.CreateFromCollection(arr.AsCollection)
  Test.Assert "Collection constructor uses default", derived.Base = 5 And derived(5) = 10
  Set derived = stdArray.CreateFromString("a,b")
  Test.Assert "String constructor uses default", derived.Base = 5 And derived(6) = "b"
  Set derived = stdArray.CreateWithOptions(10, 10, 1, 7, 8)
  Test.Assert "CreateWithOptions overrides default", derived.Base = 1 And derived(1) = 7 And derived(2) = 8
  Set derived = stdArray.CreateWithOptions(10, 10, Empty)
  derived.Push 42
  Test.Assert "CreateWithOptions omitted base", derived.Base = 5 And derived(5) = 42
  Set derived = New stdArray
  derived.protInit 10, 10, -2
  derived.Push 7
  Test.Assert "protInit accepts explicit base", derived.Base = -2 And derived(-2) = 7
  Set derived = New stdArray
  derived.protInit 10, 10
  derived.Push 8
  Test.Assert "protInit uses default", derived.Base = 5 And derived(5) = 8
  arr(1) = 22
  values = arr.arr
  Test.Assert "Export uses instance bounds", LBound(values) = 0 And UBound(values) = 2 And values(1) = 22
  arr.arr = Array(10, 20, 30)
  Test.Assert "Assignment retains base", arr.Base = 0 And arr(0) = 10
  Test.Assert "Search indices use base", arr.indexOf(10) = 0 And arr.lastIndexOf(30) = 2
  Test.Assert "Search start uses base", arr.indexOf(10, 1) = -1 And Not arr.includes(10, 1)
  Test.Assert "Membership works at zero", arr.includes(10) And Not arr.includes(99)
  Set derived = stdArray.CreateWithOptions(10, 10, -2, 10, 20, 30)
  Test.Assert "Membership works at negative one", derived.indexOf(20) = -1 And derived.includes(20) And Not derived.includes(99)
  Set derived = arr.Clone
  Test.Assert "Clone preserves base", derived.Base = 0 And derived.Join = "10,20,30"
  Set derived = arr.Reverse
  Test.Assert "Reverse preserves base", derived.Base = 0 And derived.Join = "30,20,10"
  Set derived = arr.Map(stdLambda.Create("$1*2"))
  Test.Assert "Map preserves base", derived.Base = 0 And derived(0) = 20 And derived(2) = 60
  Set derived = arr.Filter(stdLambda.Create("$1>10"))
  Test.Assert "Filter preserves base", derived.Base = 0 And derived.Join = "20,30"
  Set derived = arr.Unique()
  Test.Assert "Unique preserves base", derived.Base = 0 And derived.Join = "10,20,30"
  Set derived = arr.Sort()
  Test.Assert "Sort preserves base", derived.Base = 0 And derived.Join = "10,20,30"
  Set derived = arr.Concat(stdArray.CreateWithOptions(10, 10, 9, 40))
  Test.Assert "Concat accepts different bases", derived.Base = 0 And derived.Join = "10,20,30,40"
  Set derived = arr.Slice(0, 1)
  Test.Assert "Slice uses instance indices", derived.Base = 0 And derived.Join = "10,20"
  Set chunks = arr.SplitToChunks(2)
  Set derived = chunks(0)
  Test.Assert "Chunks preserve base", chunks.Base = 0 And derived.Base = 0 And derived.Join = "10,20"
  Dim groups As Object: Set groups = arr.GroupBy(stdLambda.Create("$1>10"))
  Set derived = groups(True)
  Test.Assert "Groups preserve base", derived.Base = 0 And derived.Join = "20,30"
  Set derived = arr.Splice(1, 1, stdArray.CreateWithOptions(10, 10, 5, 40, 50))
  Test.Assert "Splice uses mixed bases", derived.Base = 0 And derived.Join = "20" And arr.Join = "10,40,50,30"
  Test.Assert "Remove uses base", arr.Remove(0) = 10 And arr.Join = "40,50,30"
  Test.Assert "Shift and pop use order", arr.Shift = 40 And arr.Pop = 30
  arr.Unshift 70
  arr.Push 80
  Test.Assert "Push/unshift retain base", arr(0) = 70 And arr.Join = "70,50,80"
  Dim obj As Collection: Set obj = New Collection
  Set arr(0) = obj
  Test.Assert "Object assignment uses base", arr(0) Is obj
  arr.PutItem 1, obj
  Test.Assert "PutItem uses base", arr(1) Is obj
  Set arr = stdArray.CreateWithOptions(2, 2, 0, 1, 2, 3)
  Set derived = arr.Splice(1, 2, stdArray.Create())
  Test.Assert "Splice deletes through end", arr.Join = "1" And derived.Join = "2,3"
  Set derived = arr.Splice(1, 0, arr)
  Test.Assert "Splice snapshots self insertion", arr.Join = "1,1" And derived.Length = 0
  Set arr = stdArray.CreateWithOptions(2, 2, -3, 1, 2, 3, 4)
  Set derived = arr.Splice(-2, 2, stdArray.CreateWithOptions(1, 1, 7, 9))
  Test.Assert "Splice shrinks and shifts tail", arr.Join = "1,9,4" And derived.Join = "2,3"
  Dim events As stdArrayBaseEvents: Set events = New stdArrayBaseEvents
  Set events.Values = arr
  arr.Push 8
  Test.Assert "Push events use instance index", events.BeforeIndex = 0 And events.AfterIndex = 0
  arr.Pop
  Test.Assert "Pop events use removed index", events.BeforeIndex = 0 And events.AfterIndex = 0
  arr.Unshift 0
  Test.Assert "Unshift events use base", events.BeforeIndex = -3 And events.AfterIndex = -3
  arr.Remove -2
  Test.Assert "Remove events use instance index", events.BeforeIndex = -2 And events.AfterIndex = -2
  stdArray.DefaultBase = originalBase
  Exit Sub
Failed:
  failure = Err.Number: description = Err.Description
  stdArray.DefaultBase = originalBase
  Err.Raise failure, "stdArrayTests.testDefaultBase", description
End Sub
