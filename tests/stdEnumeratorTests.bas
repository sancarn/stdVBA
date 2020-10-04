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
    Dim e2 as stdEnumerator: set e2 = stdEnumerator.CreateFromIEnumVariant(c1)
    Debug.Print e1.join(",")

    'test.Assert "", 

End Sub

