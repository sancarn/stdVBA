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
    set tmp = arr.Group(lookup)
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