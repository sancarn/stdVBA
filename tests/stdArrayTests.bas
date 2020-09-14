Attribute VB_Name = "stdArrayTests"

Sub testAll()
    Test.Topic "stdArray"

    Dim arr as stdArray
    set arr = stdArray.Create(1,2,3)
    Test.Assert "Check array exists", not arr is nothing
    Test.Assert "Check item 1", arr.item(1) = 1
    Test.Assert "Check item 2", arr.item(2) = 2
    Test.Assert "Check item 3", arr.item(3) = 3

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

    'Remove
    'Slice
    'Splice
    'Clone
    'Reverse
    'Concat

    'Testing join - ease of future tests
    Test.assert "Join 1 default seperator", arr.join() = "1,2,3"
    Test.Assert "Join 2 with seperator", arr.join(";") = "1;2;3"

    'item set
    'item let
    'PutItem
    'indexOf
    'lastIndexOf
    'includes
    'count

    'CALLBACKS, using stdLambda:
    'IsEvery
    'IsSome
    'ForEach
    'Map
    'Unique
    'Reduce
    'Filter
    'Count_By
    'groupBy
    
End Sub