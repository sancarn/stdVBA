<<<<<<< HEAD
Attribute VB_Name = "stdEnumeratorTests"
Sub testAll()
    test.Topic "stdEnumerator"
    
    'Create and populate test collection
    Dim c1 as collection
    set c1 = new collection
    c1.add 1
    c1.add 2
    c1.add 3
    c1.add 4
    c1.add 5
    c1.add 6
    c1.add 7
    c1.add 8
    c1.add 9

    Dim c2 as collection
    set c2 = new collection
    c2.add "Lorem"
    c2.add "ipsum"
    c2.add "dolor"
    c2.add "sit"
    c2.add "amet"
    c2.add "consectetur"
    c2.add "adipiscing"
    c2.add "elit"
    c2.add "sed"
    c2.add "do"
    c2.add "eiusmod"
    c2.add "tempor"
    c2.add "incididunt"
    c2.add "ut"
    c2.add "labore"
    c2.add "et"
    c2.add "dolore"
    c2.add "magna"
    c2.add "aliqua"
    c2.add "Ut"
    c2.add "enim"
    c2.add "ad"
    c2.add "minim"
    c2.add "veniam"
    c2.add "quis"
    c2.add "nostrud"
    c2.add "exercitation"
    c2.add "ullamco"
    c2.add "laboris"
    c2.add "nisi"
    c2.add "ut"
    c2.add "aliquip"
    c2.add "ex"
    c2.add "ea"
    c2.add "commodo"
    c2.add "consequat"
    c2.add "Duis"
    c2.add "aute"
    c2.add "irure"
    c2.add "dolor"
    c2.add "in"
    c2.add "reprehenderit"
    c2.add "in"
    c2.add "voluptate"
    c2.add "velit"
    c2.add "esse"
    c2.add "cillum"
    c2.add "dolore"
    c2.add "eu"
    c2.add "fugiat"
    c2.add "nulla"
    c2.add "pariatur"
    c2.add "Excepteur"
    c2.add "sint"
    c2.add "occaecat"
    c2.add "cupidatat"
    c2.add "non"
    c2.add "proident"
    c2.add "sunt"
    c2.add "in"
    c2.add "culpa"
    c2.add "qui"
    c2.add "officia"
    c2.add "deserunt"
    c2.add "mollit"
    c2.add "anim"
    c2.add "id"
    c2.add "est"
    c2.add "laborum"

    On Error Resume Next
    Dim e1 as stdEnumerator: set e1 = stdEnumerator.CreateFromIEnumVariant(c1)
    Dim e2 as stdEnumerator: set e2 = stdEnumerator.CreateFromIEnumVariant(c2)
    
    'We'll be using join a lot for tests so test this first:
    Test.Assert "Join", e1.Join() = "1,2,3,4,5,6,7,8,9"
    Test.Assert "Join w/ Delim", e1.Join("|") = "1|2|3|4|5|6|7|8|9"
    Test.Assert "Map", e1.map(stdLambda.Create("$1*2")).join() = "2,4,6,8,10,12,14,16,18"
    Test.Assert "Map w/ Index", e1.map(stdLambda.Create("$1+$2"),true).join() = "2,4,6,8,10,12,14,16,18"
    Test.Assert "Reverse", e1.reverse().join() = "9,8,7,6,5,4,3,2,1"
    Test.Assert "Filter", e1.Filter(stdLambda.Create("$1<=4")).join() = "1,2,3,4"
    Test.Assert "Filter w/ Index", e1.Filter(stdLambda.Create("($1+$2)<=4"),true).join() = "1,2"
    Test.Assert "Concat", e1.concat(c1).join() = "1,2,3,4,5,6,7,8,9,1,2,3,4,5,6,7,8,9"

    With e1.concat(c1)
        Test.Assert "IndexOf 1", .indexOf(5)=5
        Test.Assert "IndexOf 2", .indexOf(42)=0
        Test.Assert "lastIndexOf 1", .lastIndexOf(5)=14
        Test.Assert "lastIndexOf 2", .lastIndexOf(42)=0
        Test.Assert "includes 1", .includes(4)
        Test.Assert "includes 2", not .includes(42)
        Test.Assert "Unique", .Unique().Join() = "1,2,3,4,5,6,7,8,9"
    End with

    Test.Assert "Reduce", e1.reduce(stdLambda.Create("$1+$2"))=45
    Test.Assert "Reduce w/ Initial value", e1.reduce(stdLambda.Create("$1+$2"),10)=55
    Test.Assert "CountBy", e2.countBy(stdLambda.Create("len($1)<=5"))=39
    Test.Assert "CheckAll", e1.checkAll(stdLambda.Create("$1<=42"))
    Test.Assert "CheckAny", e1.checkAny(stdLambda.Create("$1=5"))
    Test.Assert "CheckNone 1", Not e1.checkNone(stdLambda.Create("$1=5"))
    Test.Assert "CheckNone 2", e1.checkNone(stdLambda.Create("$1=42"))
    Test.Assert "CheckOnlyOne 1 matched thus true", e1.checkOnlyOne(stdLambda.Create("$1=1"))
    Test.Assert "CheckOnlyOne 2 matched thus false", not e1.checkOnlyOne(stdLambda.Create("$1<=2"))
    Test.Assert "CheckOnlyOne 0 matched thus false", not e1.checkOnlyOne(stdLambda.Create("$1=42"))
    Test.Assert "Max", e1.max()=9
    Test.Assert "Max w/ callback", e2.max(stdLambda.Create("len($1)"))="reprehenderit"
    Test.Assert "Min", e1.min()=1
    Test.Assert "Min w/ callback", e2.min(stdLambda.Create("len($1)"))="do"
    Test.Assert "Sum", e1.sum()=45
    Test.Assert "Sum", e1.sum(stdLambda.Create("$1*2"))=90
    Test.Assert "FindFirst found", e2.FindFirst(stdLambda.Create("len($1)=6"))="tempor"
    Test.Assert "FindFirst not found", isNull(e2.FindFirst(stdLambda.Create("len($1)=42")))
    Test.Assert "Sort", e2.Sort(stdLambda.Create("len($1)")).Join = "do,ut,et,Ut,ad,ut,ex,ea,in,in,eu,in,id,sit,sed,non,qui,est,amet,elit,enim,quis,nisi,Duis,aute,esse,sint,sunt,anim,Lorem,ipsum,dolor,magna,minim,irure,dolor,velit,nulla,culpa,tempor,labore,dolore,aliqua,veniam,cillum,dolore,fugiat,mollit,eiusmod,nostrud,ullamco,laboris,aliquip,commodo,officia,laborum,pariatur,occaecat,proident,deserunt,consequat,voluptate,Excepteur,cupidatat,adipiscing,incididunt,consectetur,exercitation,reprehenderit"
    Test.Assert "Length", e1.Length()=9
    Test.Assert "Item 1 gets item", e1.item(5)=5
    Test.Assert "Item 2 returns null", isNull(e1.item(99))

    'ForEach style tests
    Dim tCol as collection
    set tCol = new collection
    Call e1.forEach(stdLambda.Create("$1#add($2)").bind(tCol))
    Test.Assert "ForEach", stdEnumerator.CreateFromIEnumVariant(tCol).join() = "1,2,3,4,5,6,7,8,9"
    set tCol = new collection
    Call e1.forEach(stdLambda.Create("$1#add($2+$3)").bind(tCol),true)
    Test.Assert "ForEach w\ Index", stdEnumerator.CreateFromIEnumVariant(tCol).join() = "2,4,6,8,10,12,14,16,18"
    set tCol = new collection
    Call e1.cycle(2, stdLambda.Create("$1#add($2)").bind(tCol))
    Test.Assert "Cycle", stdEnumerator.CreateFromIEnumVariant(tCol).join() = "1,2,3,4,5,6,7,8,9,1,2,3,4,5,6,7,8,9"

    'Big flatten example:
    set tCol = new collection
    tCol.add new collection '1
    tCol(1).add 1
    tCol(1).add 2
    tCol.add 3 '2
    tCol.add 4 '3
    tCol.add new collection '4
    tCol.add new collection '5
    tCol(5).add 5
    tCol(5).add 6
    Test.Assert "Flatten", stdEnumerator.CreateFromIEnumVariant(tCol).Flatten().join() = "1,2,3,4,5,6"

    Dim dict as object
    set dict = e1.groupBy(stdLambda.Create("if ($1 mod 2) = 0 then ""Even"" else ""Odd"""))
    Test.Assert "GroupBy - Even numbers", dict("Even").join() = "2,4,6,8"
    Test.Assert "GroupBy - Odd numbers" , dict("Odd").join() = "1,3,5,7,9"
End Sub



' {} [X] Sort() as stdArray
' {T} [X] Reverse() as stdArray
' {T} [X] ForEach                  +T w/Index
' {T} [X] Map                      +T w/Index
' {T} [X] Unique
' {T} [X] Filter
' {T} [X] Concat
' {T} [X] Join
' {T} [X] indexOf                  +T w/ not found value
' {T} [X] lastIndexOf              +T w/ not found value
' {T} [X] includes                 +
' {T} [X] reduce
' {T} [X] countBy
' {T} [X] max
' {T} [X] min
' {T} [X] sum
' {T} [X] Flatten
' {T} [X] cycle
' {T} [X] findFirst
' {T} [X] checkAll
' {T} [X] checkAny
' {T} [X] checkNone
' {T} [X] checkOnlyOne
=======
Attribute VB_Name = "stdEnumeratorTests"
Sub testAll()
    test.Topic "stdEnumerator"
    
    'Create and populate test collection
    Dim c1 as collection
    set c1 = new collection
    c1.add 1
    c1.add 2
    c1.add 3
    c1.add 4
    c1.add 5
    c1.add 6
    c1.add 7
    c1.add 8
    c1.add 9

    Dim c2 as collection
    set c2 = new collection
    c2.add "Lorem"
    c2.add "ipsum"
    c2.add "dolor"
    c2.add "sit"
    c2.add "amet"
    c2.add "consectetur"
    c2.add "adipiscing"
    c2.add "elit"
    c2.add "sed"
    c2.add "do"
    c2.add "eiusmod"
    c2.add "tempor"
    c2.add "incididunt"
    c2.add "ut"
    c2.add "labore"
    c2.add "et"
    c2.add "dolore"
    c2.add "magna"
    c2.add "aliqua"
    c2.add "Ut"
    c2.add "enim"
    c2.add "ad"
    c2.add "minim"
    c2.add "veniam"
    c2.add "quis"
    c2.add "nostrud"
    c2.add "exercitation"
    c2.add "ullamco"
    c2.add "laboris"
    c2.add "nisi"
    c2.add "ut"
    c2.add "aliquip"
    c2.add "ex"
    c2.add "ea"
    c2.add "commodo"
    c2.add "consequat"
    c2.add "Duis"
    c2.add "aute"
    c2.add "irure"
    c2.add "dolor"
    c2.add "in"
    c2.add "reprehenderit"
    c2.add "in"
    c2.add "voluptate"
    c2.add "velit"
    c2.add "esse"
    c2.add "cillum"
    c2.add "dolore"
    c2.add "eu"
    c2.add "fugiat"
    c2.add "nulla"
    c2.add "pariatur"
    c2.add "Excepteur"
    c2.add "sint"
    c2.add "occaecat"
    c2.add "cupidatat"
    c2.add "non"
    c2.add "proident"
    c2.add "sunt"
    c2.add "in"
    c2.add "culpa"
    c2.add "qui"
    c2.add "officia"
    c2.add "deserunt"
    c2.add "mollit"
    c2.add "anim"
    c2.add "id"
    c2.add "est"
    c2.add "laborum"

    On Error Resume Next
    Dim e1 as stdEnumerator: set e1 = stdEnumerator.CreateFromIEnumVariant(c1)
    Dim e2 as stdEnumerator: set e2 = stdEnumerator.CreateFromIEnumVariant(c2)
    
    'We'll be using join a lot for tests so test this first:
    Test.Assert "Join", e1.Join() = "1,2,3,4,5,6,7,8,9"
    Test.Assert "Join w/ Delim", e1.Join("|") = "1|2|3|4|5|6|7|8|9"
    Test.Assert "Map", e1.map(stdLambda.Create("$1*2")).join() = "2,4,6,8,10,12,14,16,18"
    Test.Assert "Map w/ Index", e1.map(stdLambda.Create("$1+$2"),true).join() = "2,4,6,8,10,12,14,16,18"
    Test.Assert "Reverse", e1.reverse().join() = "9,8,7,6,5,4,3,2,1"
    Test.Assert "Filter", e1.Filter(stdLambda.Create("$1<=4")).join() = "1,2,3,4"
    Test.Assert "Filter w/ Index", e1.Filter(stdLambda.Create("($1+$2)<=4"),true).join() = "1,2"
    Test.Assert "Concat", e1.concat(c1).join() = "1,2,3,4,5,6,7,8,9,1,2,3,4,5,6,7,8,9"

    With e1.concat(c1)
        Test.Assert "IndexOf 1", .indexOf(5)=5
        Test.Assert "IndexOf 2", .indexOf(42)=0
        Test.Assert "lastIndexOf 1", .lastIndexOf(5)=14
        Test.Assert "lastIndexOf 2", .lastIndexOf(42)=0
        Test.Assert "includes 1", .includes(4)
        Test.Assert "includes 2", not .includes(42)
        Test.Assert "Unique", .Unique().Join() = "1,2,3,4,5,6,7,8,9"
    End with

    Test.Assert "Reduce", e1.reduce(stdLambda.Create("$1+$2"))=45
    Test.Assert "Reduce w/ Initial value", e1.reduce(stdLambda.Create("$1+$2"),10)=55
    Test.Assert "CountBy", e2.countBy(stdLambda.Create("len($1)<=5"))=39
    Test.Assert "CheckAll", e1.checkAll(stdLambda.Create("$1<=42"))
    Test.Assert "CheckAny", e1.checkAny(stdLambda.Create("$1=5"))
    Test.Assert "CheckNone 1", Not e1.checkNone(stdLambda.Create("$1=5"))
    Test.Assert "CheckNone 2", e1.checkNone(stdLambda.Create("$1=42"))
    Test.Assert "CheckOnlyOne 1 matched thus true", e1.checkOnlyOne(stdLambda.Create("$1=1"))
    Test.Assert "CheckOnlyOne 2 matched thus false", not e1.checkOnlyOne(stdLambda.Create("$1<=2"))
    Test.Assert "CheckOnlyOne 0 matched thus false", not e1.checkOnlyOne(stdLambda.Create("$1=42"))
    Test.Assert "Max", e1.max()=9
    Test.Assert "Max w/ callback", e2.max(stdLambda.Create("len($1)"))="reprehenderit"
    Test.Assert "Min", e1.min()=1
    Test.Assert "Min w/ callback", e2.min(stdLambda.Create("len($1)"))="do"
    Test.Assert "Sum", e1.sum()=45
    Test.Assert "Sum", e1.sum(stdLambda.Create("$1*2"))=90
    Test.Assert "FindFirst found", e2.FindFirst(stdLambda.Create("len($1)=6"))="tempor"
    Test.Assert "FindFirst not found", isNull(e2.FindFirst(stdLambda.Create("len($1)=42")))
    Test.Assert "Sort", e2.Sort(stdLambda.Create("len($1)")).Join = "do,ut,et,Ut,ad,ut,ex,ea,in,in,eu,in,id,sit,sed,non,qui,est,amet,elit,enim,quis,nisi,Duis,aute,esse,sint,sunt,anim,Lorem,ipsum,dolor,magna,minim,irure,dolor,velit,nulla,culpa,tempor,labore,dolore,aliqua,veniam,cillum,dolore,fugiat,mollit,eiusmod,nostrud,ullamco,laboris,aliquip,commodo,officia,laborum,pariatur,occaecat,proident,deserunt,consequat,voluptate,Excepteur,cupidatat,adipiscing,incididunt,consectetur,exercitation,reprehenderit"
    Test.Assert "Length", e1.Length()=9
    Test.Assert "Item 1 gets item", e1.item(5)=5
    Test.Assert "Item 2 returns null", isNull(e1.item(99))

    'ForEach style tests
    Dim tCol as collection
    set tCol = new collection
    Call e1.forEach(stdLambda.Create("$1#add($2)").bind(tCol))
    Test.Assert "ForEach", stdEnumerator.CreateFromIEnumVariant(tCol).join() = "1,2,3,4,5,6,7,8,9"
    set tCol = new collection
    Call e1.forEach(stdLambda.Create("$1#add($2+$3)").bind(tCol),true)
    Test.Assert "ForEach w\ Index", stdEnumerator.CreateFromIEnumVariant(tCol).join() = "2,4,6,8,10,12,14,16,18"
    set tCol = new collection
    Call e1.cycle(2, stdLambda.Create("$1#add($2)").bind(tCol))
    Test.Assert "Cycle", stdEnumerator.CreateFromIEnumVariant(tCol).join() = "1,2,3,4,5,6,7,8,9,1,2,3,4,5,6,7,8,9"

    'Big flatten example:
    set tCol = new collection
    tCol.add new collection '1
    tCol(1).add 1
    tCol(1).add 2
    tCol.add 3 '2
    tCol.add 4 '3
    tCol.add new collection '4
    tCol.add new collection '5
    tCol(5).add 5
    tCol(5).add 6
    Test.Assert "Flatten", stdEnumerator.CreateFromIEnumVariant(tCol).Flatten().join() = "1,2,3,4,5,6"

    Dim dict as object
    set dict = e1.groupBy(stdLambda.Create("if ($1 mod 2) = 0 then ""Even"" else ""Odd"""))
    Test.Assert "GroupBy - Even numbers", dict("Even").join() = "2,4,6,8"
    Test.Assert "GroupBy - Odd numbers" , dict("Odd").join() = "1,3,5,7,9"
End Sub



' {} [X] Sort() as stdArray
' {T} [X] Reverse() as stdArray
' {T} [X] ForEach                  +T w/Index
' {T} [X] Map                      +T w/Index
' {T} [X] Unique
' {T} [X] Filter
' {T} [X] Concat
' {T} [X] Join
' {T} [X] indexOf                  +T w/ not found value
' {T} [X] lastIndexOf              +T w/ not found value
' {T} [X] includes                 +
' {T} [X] reduce
' {T} [X] countBy
' {T} [X] max
' {T} [X] min
' {T} [X] sum
' {T} [X] Flatten
' {T} [X] cycle
' {T} [X] findFirst
' {T} [X] checkAll
' {T} [X] checkAny
' {T} [X] checkNone
' {T} [X] checkOnlyOne
>>>>>>> master
' {T} [X] groupBy