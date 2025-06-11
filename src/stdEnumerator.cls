VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdEnumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@module
'@description A class which implements an enumerator object. Use this class to
'iterate over a collection of data.
'@example ```vb
''Find the longest worksheet name
'Dim ws as stdEnumerator: set ws = stdEnumerator.CreateFromIEnumVariant(ThisWorkbook.Worksheets)
'Dim sht as Worksheet: set sht = ws.Max(stdLambda.Create("len($1.Name)"))
'Debug.Print sht.Name
'```
'@example Assuming we have a table named `People` with the following data:
'
' | Name  | Age | Gender |
' |-------|-----|--------| 
' | John  | 25  | Male   |
' | Jane  | 30  | Female |
' | Bob   | 35  | Male   |
' | Alice | 40  | Female |
'
'The following code will print the names of all females:
'```vb
''Loop over a table and print names of females
'Dim loPeople as ListObject: set loPeople = ThisWorkbook.Worksheets("People").ListObjects("People")
'Dim tbl as stdEnumerator: set tbl = stdEnumerator.CreateFromListObject(loPeople)
'Dim females as stdEnumerator: set females = tbl.Filter(stdLambda.Create("$1.Gender = ""Female"""))
'Dim names as stdEnumerator: set names = females.Map(stdLambda.Create("$1.Name"))
'Debug.Print names.Join(vbCrLf)
'```
'@example Get all worksheets of all open workbooks 
'```vb
'Dim allWorkbooks as stdEnumerator: set allWorkbooks = stdEnumerator.CreateFromIEnumVariant(Application.Workbooks)
'Dim allWorksheets as stdEnumerator: set allWorksheets = allWorkbooks.Map(stdLambda.Create("$1.Worksheets"))
'set allWorksheets = allWorksheets.Flatten
'```
'@example Create an archimedean spiral using the `CreateFromCallable` feature. 
'```vb
''Generates poiknts on an Archimedean spiral.
''@param lastItem - The last point generated (or null for the first)
''@param index - The current index of the point being generated
''@returns Array<Long> - A point containing an X and Y coordinate on the spiral.
'Public Function GetArchimedesSpiral(ByVal lastItem as variant, ByVal index as long) as variant
'  'Stop at 100 points
'  if index > 100 then
'    getArchimedesSpiral = null
'    exit function
'  end if
' 
'  const a as double = 0.1 'Controls the tightness of the spiral
'  const b as double = 0.5 'Controls the spacing between the turns
'
'  'Calculate the angle (theta) for the current point.
'  'We can make the angle increase linearly with the index for a smooth spiral.
'  Dim theta as Double: theta = index * b
'
'  'Calculate the radius (r) using the Archimedean spiral equation: r = a * theta
'  Dim r as double: r = a * theta
'
'  'Convert polar coordinates (r, theta) to Cartesian coordinates (x, y)
'  Dim x as double: x = r * cos(theta)
'  Dim y as double: y = r * sin(theta)
'  
'  GetArchimedesSpiral = Array(x, y)
'End Function
'Public Sub testSpiralPoints()
'  Dim cbSpiral as stdCallback: Set cbSpiral = stdCallback.CreateFromModule("MyModule", "GetArchimedesSpiral")
'  Dim spiral as stdEnumerator: Set spiral = stdEnumerator.CreateFromCallable(cbSpiral)
'  Dim arr: arr = spiral.AsArray
'  stop
'End Function
'```
'@example Generate points on a 2D random walk, where each step is a random direction and distance from the last point.
'```vb
'Public Function GetRandomWalk(ByVal lastItem As Variant, ByVal index As Long) As Variant
'  'Stop at index 100
'  If index > 100 Then
'    GetRandomWalk = Null
'    Exit Function
'  End If
'  
'  Const Pi As Double = 3.14159
'  Dim x As Double, y As Double, angle As Double
'  If index = 1 Then
'    x = 0
'    y = 0
'    angle = Rnd() * 2 * Pi
'  Else
'    x = lastItem(0)
'    y = lastItem(1)
'    angle = lastItem(2)
'  End If
'
'  Const maxDist As Double = 5
'  Dim maxAngle As Double: maxAngle = Pi / 2
'  Dim diff1 As Double: diff1 = (Rnd() - 0.5) * 2 * maxDist
'  Dim diff2 As Double: diff2 = (Rnd() - 0.5) * 2 * maxAngle
'  Dim nextAngle As Double: nextAngle = angle + diff2
'  x = x + diff1 * Cos(nextAngle)
'  y = y + diff1 * Sin(nextAngle)
'  angle = nextAngle
'  GetRandomWalk = Array(x, y, angle)
'End Function
'
'Public Sub TestRandomWalk()
'  Dim cbRandomWalk As stdCallback: Set cbRandomWalk = stdCallback.CreateFromModule("Main", "GetRandomWalk")
'  Dim walk As stdEnumerator: Set walk = stdEnumerator.CreateFromCallable(cbRandomWalk, 100)
'  Dim arr: arr = walk.AsArray
'  Stop
'End Sub
'```


'Spec:
'For some enumerable set ["foo","bar","baz","bazzer"] let stdEnumerator implement functions:
'    * forEach
'    * map
'    ...
'In this way we can stack calls as follows:
'    Debug.print enumerableSet.map(s=>len(s)).filter(i=>i<=3).sum()

'Functions implemented on this class:
'CONSTRUCTORS
'    [X] CreateFromIEnumVariant
'    [X] CreateFromCallable
'    [X] CreateFromArray
'    [ ] CreateFromDictionary
'    [X] protInit
'
'INSTANCE METHODS
'Many methods were inspired by those in Ruby's Enumerable: https://ruby-doc.org/core-2.7.2/Enumerable.html
'    [X] asCollection()
'    [X] asArray(iType as vbVarType) 
'    [X] Sort()
'    [X] Reverse()
'    [X] ForEach
'    [X] Map
'    [X] Unique
'    [X] Filter
'    [X] Concat
'    [X] Join
'    [X] indexOf
'    [X] lastIndexOf
'    [X] includes
'    [X] reduce
'    [X] countBy
'    [X] groupBy
'    [X] max(cb)
'    [X] min(cb)
'    [X] sum(cb)
'    [X] Flatten
'    [X] cycle
'    [X] findFirst
'    [X] checkAll
'    [X] checkAny
'    [X] checkNone
'    [X] checkOnlyOne
'    [X] item
'    [X] length

'TODO: Implement the following methods:
'    [?] each_cons  [1,2,3,4,5].each_cons(2,cb) ==> cb([1,2]) ==> cb([2,3]) ==> cb([3,4]) ==> cb([4,5])
'    [?] each_slice [1,2,3,4,5].each_slice(2,cb) ==> cb([1,2]) ==> cb([3,4]) ==> cb([5])
'    [?] partition   [1,2,3,4,5,6].partition(a=>a%2=0) ==> [[2,4,6],[1,3,5]]
'    [?] zip         [1,2,3].zip([4,5,6]) ==> [[1,4],[2,5],[3,6]]          |            [1,2,3].zip([1,2]) ==> [[1,1],[2,2],[3,null]]

'WHAT WE WON'T DO:
'    with_index    'this can't be easily done, so instead all methods will call callbacks with index supplied.
'    with_object   'this can be done with cb.Bind()

'It may be difficult to think of uses for zip, so here are some:
'    a = stdEnumerator.CreateFromArray(split("a b c"," ")).zip([1,2,3]).to_dict() ==> {a:1, b:2, c:3}
'    vector addition and multiplication:
'    [1,2].zip([2,3]).map(e=>e[0]+e[1]) ==> [3,5]
'    [1,2].zip([2,3]).map(e=>e[0]*e[1]) ==> [2,6]

'TODO: Things we can't do (yet)
'    take <-- can't do this unless we implement IEnumVARIANT and call Next() method
'    tally <-- Would like to do this but can't until we have stdDictionary     ["a","b","c","b","a","b"].tally ==> {a:2, b:3, c:1}
'    to_dict <-- requires stdDictioanry                                        ["a",1,"b",2].to_dict ==> {a:1, b:2}
'    groupBy <-- requires stdDictionary

#If Mac Then
    'Unfortunately I can't find any examples of creating an empty array on Mac :(. If you want to create Mac - compatable code don't use AsArray() with anything other than
    'vbVariant data type.
#Else
    #If VBA7 Then
        Private Declare PtrSafe Function createEmptyArrayBoolean Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbBoolean, Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) As Boolean()
        Private Declare PtrSafe Function createEmptyArrayByte Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbByte, Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) As Byte()
        Private Declare PtrSafe Function createEmptyArrayCurrency Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbCurrency, Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) As Currency()
        Private Declare PtrSafe Function createEmptyArrayDate Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbDate, Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) As Date()
        Private Declare PtrSafe Function createEmptyArrayDouble Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbDouble, Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) As Double()
        Private Declare PtrSafe Function createEmptyArrayLong Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbLong, Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) As Long()
        Private Declare PtrSafe Function createEmptyArrayObject Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbObject, Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) As Object()
        Private Declare PtrSafe Function createEmptyArrayString Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbString, Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) As String()
        Private Declare PtrSafe Function createEmptyArrayVariant Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbVariant, Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) As Variant()
        #If Win64 Then
            Private Declare PtrSafe Function createEmptyArrayLongLong Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbLongLong, Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) As LongLong()
        #End If
    #Else
        Private Declare Function createEmptyArrayBoolean     Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbBoolean  , Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) as Boolean()
        Private Declare Function createEmptyArrayByte        Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbByte     , Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) as Byte()
        Private Declare Function createEmptyArrayCurrency    Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbCurrency , Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) as Currency()
        Private Declare Function createEmptyArrayDate        Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbDate     , Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) as Date()
        Private Declare Function createEmptyArrayDouble      Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbDouble   , Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) as Double()
        Private Declare Function createEmptyArrayLong        Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbLong     , Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) as Long()
        Private Declare Function createEmptyArrayObject      Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbObject   , Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) as Object()
        Private Declare Function createEmptyArrayString      Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbString   , Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) as String()
        Private Declare Function createEmptyArrayVariant     Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbVariant  , Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) as Variant()
    #End If
#End If
Private Enum EnumeratorType
    FromCallable
    FromCallableVerbose
    FromIEnumVariant
    FromArray
End Enum

Private Type SortStruct
    value As Variant
    sortValue As Variant
    iIndex As Long
    iNext As Long
    iPrev As Long
End Type

Private Type TCurrentIteration
    Value As Variant
    Key As Variant
    Index As Long
End Type
Private Type TThis
    mode As EnumeratorType
    EnumObject As Object
    BaseArray As Variant
    Callback As stdICallable
    MaxLength As Long

    'Variables to allow procedural enumeration.
    Current As TCurrentIteration
End Type
Private This As TThis

'Variables to allow procedural enumeration.


Private Const VT_UNKNOWN = &Hd

'Create an enumerator from a callable. Simply return either the next item, or null to indicate the end of the item set.
'@constructor
'@param cb as stdICallable<(lastItem: variant,index: long)=> variant | null> - Callback which returns either the next item.
'       If no next item, null should be returned.
'@param iMaxLength - Maximum length of enumerator. The enumerator will refuse to go beyond this number of items.
'       Used in order to prevent runaway processes.
'@returns - Enumerator created from callback function
Public Function CreateFromCallable(ByVal cb As stdICallable, Optional ByVal iMaxLength As Long = 1000000) As stdEnumerator
    Set CreateFromCallable = New stdEnumerator
    Call CreateFromCallable.protInit(EnumeratorType.FromCallable, iMaxLength, cb)
End Function

'Creates an enumerator from a callable. This is more complex than non-verbose callable, however offers full flexibility of the loop.
'@constructor
'@param cb as ByVal stdICallable<(lastItem: variant,index: long)=>Array<Boolean,Long,Variant,Variant>> - Callback which returns an array. The array should be of the form:
'[hasANextItem, NextIndex, ValueToReturn, KeyToReturn]. This allows `ValueToReturn` to be null unlike Non-verbose callable.
'@param iMaxLength - Maximum length of enumerator. The enumerator will refuse to go beyond this number of items. Used in order to prevent runaway processes.
'@returns - Enumerator created from callback function
Public Function CreateFromCallableVerbose(ByVal cb As stdICallable, Optional ByVal iMaxLength As Long = 1000000) As stdEnumerator
    Set CreateFromCallableVerbose = New stdEnumerator
    Call CreateFromCallableVerbose.protInit(EnumeratorType.FromCallableVerbose, iMaxLength, cb)
End Function

'Create an enumerator from an array
'@constructor
'@param v - the Array to create the enumerator from.
'@param iMaxLength - Maximum length of enumerator. The enumerator will refuse to go beyond this number of items. Used in order to prevent runaway processes.
'@returns - An enumerator representing data in array
Public Function CreateFromArray(ByVal v As Variant, Optional ByVal iMaxLength As Long = 1000000) As stdEnumerator
    Set CreateFromArray = New stdEnumerator
    Call CreateFromArray.protInit(EnumeratorType.FromArray, iMaxLength, v)
End Function

'Create an enumerator from an object which supports IEnumVARIANT.
'@constructor
'@param o - the object to create the enumerator from. This object should support IEnumVARIANT
'@param iMaxLength - Maximum length of enumerator. The enumerator will refuse to go beyond this number of items. Used in order to prevent runaway processes.
'@returns - An enumerator representing variants in the IEnumVARIANT
Public Function CreateFromIEnumVariant(ByVal o As Object, Optional ByVal iMaxLength As Long = 1000000) As stdEnumerator
    Set CreateFromIEnumVariant = New stdEnumerator
    Call CreateFromIEnumVariant.protInit(EnumeratorType.FromIEnumVariant, iMaxLength, o)
End Function


'Create an empty enumerator
'@constructor
'@param iMaxLength - Maximum length of enumerator. The enumerator will refuse to go beyond this number of items. Used in order to prevent runaway processes.
'@returns - An empty enumerator
Public Function CreateEmpty(Optional ByVal iMaxLength As Long = 1000000) As stdEnumerator
    Set CreateEmpty = New stdEnumerator
    Call CreateEmpty.protInit(EnumeratorType.FromIEnumVariant, iMaxLength, New Collection)
End Function

'Create an enumerator from an Excel ListObject. Dictionaries are created for each ListRow with Header values as keys.
'@constructor
'@param lo - ListObject to create enumerator from.
'@param iMaxLength - Maximum length of enumerator. The enumerator will refuse to go beyond this number of items. Used in order to prevent runaway processes.
'@returns - Enumerator representing ListObject data
Public Function CreateFromListObject(ByVal lo As ListObject, Optional ByVal iMaxLength As Long = 1000000, Optional ByVal dictCompare As Long = vbTextCompare) As stdEnumerator
    Dim c As Collection: Set c = New Collection
    If Not lo.DataBodyRange Is Nothing Then
        'Get header and body data
        Dim oRow As Object
        Dim vHeader: vHeader = lo.HeaderRowRange.Value
        Dim vData: vData = lo.DataBodyRange.Value

        If TypeName(vData) = "Variant()" Then
            'Create rows for all rows
            Dim i As Long
            For i = 1 To UBound(vData, 1)
                Set oRow = CreateObject("Scripting.Dictionary")
                oRow.CompareMode = dictCompare
                Dim j As Long
                For j = 1 To UBound(vData, 2)
                    oRow(vHeader(1, j)) = vData(i, j)
                Next
                Set oRow("=ListRow") = lo.ListRows(i)
                Set oRow("=ListObject") = lo
                Set oRow("=ListColumns") = lo.ListColumns
                Call c.Add(oRow)
            Next
        Else
            'Create row for single value
            Set oRow = CreateObject("Scripting.Dictionary")
            oRow(vHeader) = vData
            c.Add oRow
        End If
        
        'Create from collection
        Set CreateFromListObject = CreateFromIEnumVariant(c)
    Else
        'No rows, so return empty enumerator
        Set CreateFromListObject = CreateEmpty()
    End If
End Function

'Create a stdEnumerator from a tree of objects
'@constructor
'@param oRoot - The parent / root node
'@param getChildren - A callback which returns the children of each individual node.
'@returns - Flat enumerator of all objects in the tree.
Public Function CreateFromTree(ByVal oRoot As Object, ByVal getChildren As stdICallable) As stdEnumerator
    Dim cRet As Collection: Set cRet = New Collection
    cRet.Add oRoot
    
    'Obtain flat collection from tree
    Dim iObj As Long, oObj As Object
    While iObj < cRet.Count
        iObj = iObj + 1
        Set oObj = cRet(iObj)
        Dim child As Variant
        For Each child In getChildren.Run(oObj)
            cRet.Add child
        Next
    Wend

    Set CreateFromTree = stdEnumerator.CreateFromIEnumVariant(cRet)
End Function

'Create a stdEnumerator from a dictionary
'@constructor
'@param o - The dictionary to create the enumerator from.
'@param iMaxLength - Maximum length of enumerator. The enumerator will refuse to go beyond this number of items. Used in order to prevent runaway processes.
'@returns - An enumerator representing the dictionary data
'@TODO: Finish function - requires key handling. Might be a bit of a pain as it requires refactoring to work with collection keys among other things.
'may not be needed given we have `CreateFromDictionaryEx`, but it would be far more performant.
Private Function CreateFromDictionary(ByVal dictionary As Object, Optional ByVal iMaxLength As Long = 1000000) As stdEnumerator
    
End Function

'Translates a dictionary to a Enumerator of dictionaries with key and value properties.
'@constructor
'@param dictionary - The dictionary to create the enumerator from.
'@param iMaxLength - Maximum length of enumerator. The enumerator will refuse to go beyond this number of items. Used in order to prevent runaway processes.
'@returns stdEnumerator<Object<Scripting.Dictionary<key: string, value: variant>>> - An enumerator representing the dictionary data
'@example ```vb
'set dict = stdEnumerator.CreateFromArray(Array(1,2,3,2,2,1,3,2,3,1,2,3)).groupBy(stdLambda.Create("$1"))
'set dictEx = stdEnumerator.CreateFromDictionaryEx(dict).forEach(stdLambda.Create("let $1.value = $1.value.count")).sort(stdLambda.Create("$1.count")).last
'```
Public Function CreateFromDictionaryEx(ByVal dictionary As Object) As stdEnumerator
    Dim keys As Variant: keys = dictionary.keys()
    Dim items As Variant: items = dictionary.items()
    Dim rows As Collection: Set rows = New Collection
    Dim i As Long
    For i = 0 To UBound(keys)
        Dim row As Object: Set row = CreateObject("Scripting.Dictionary")
        row.CompareMode = vbTextCompare
        row.add "key", keys(i)
        row.add "value", items(i)
        rows.Add row
    Next
    Set CreateFromDictionaryEx = stdEnumerator.CreateFromIEnumVariant(rows)
End Function

'Initialse the stdEnumerator
'@protected
'@param iEnumeratorType - Enumerator type to create
'@param iMaxLength - Maximum length of enumerator. The enumerator will refuse to go beyond this number of items. Used in order to prevent runaway processes.
'@param v - The data to create the enumerator from. This is either an array, a callable, or an IEnumVARIANT object.
Public Sub protInit(ByVal iEnumeratorType As Long, ByVal iMaxLength As Long, ParamArray v() As Variant)
    This.mode = iEnumeratorType
    This.MaxLength = iMaxLength
    Select Case This.mode
        Case EnumeratorType.FromIEnumVariant
            'HACK: This is a real work around for the fact that we can't access IEnumVARIANT methods (getNewEnum,getNextVariant,skipNextVariant,cloneEnumVariant,resetEnumVariant)
            'directly. Realistically this may sometimes not even work. A more long term solution would be to call getNewEnum() and then getNextVariant() until we reach the end of the
            'object. This may be implementable after stdCOM is fully implemented.

            Set This.EnumObject = v(0)
            
            'Calculate total length
            Dim iLength As Long: iLength = 0
            Dim obj As Variant
            For Each obj In This.EnumObject
                iLength = iLength + 1
            Next
            
            'Define array
            If iLength > 0 Then
                Dim vArr()
                ReDim Preserve vArr(1 To iLength)
                
                'Copy IEnumVARIANT into array
                Dim i As Long: i = 0
                For Each obj In This.EnumObject
                    i = i + 1
                    Call CopyVariant(vArr(i), obj)
                Next

                This.BaseArray = vArr
            Else
                This.BaseArray = Array()
            End If
        Case EnumeratorType.FromArray
            This.BaseArray = v(0)
        Case EnumeratorType.FromCallable
            Set This.Callback = v(0)
        Case EnumeratorType.FromCallableVerbose
            Set This.Callback = v(0)
    End Select
End Sub

' IEnumVariant methods

'For-each compatibility
'@protected
'@returns IEnumVARIANT - An enumerator with methods enumNext, enumRefresh etc.
'@example ```vb
'Dim v as variant
'For each v in myEnum
'  Debug.Pring v
'next
'```
'@TODO: Use custom IEnumVARIANT instead of casting to Collection
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Static oEnumCol As Collection: If oEnumCol Is Nothing Then Set oEnumCol = AsCollection()
    Set NewEnum = oEnumCol.[_NewEnum]
End Property



'Get the next value in the dataset
'@param nextValue - The new value retrieved.
'@returns - `True` if a value was retrieved, `false` if the end of the enumerator has been reached.
'@example ```vb
'Dim v as Variant
'While enumNext(v)
'  Debug.Print v
'Wend
'```
Public Function enumNext(ByRef nextValue As Variant) As Boolean
    With This.Current
        If NextItem(.Key, .Value, .Index) Then
            If IsObject(.Value) Then
                Set nextValue = .Value
            Else
                Let nextValue = .Value
            End If
            enumNext = True
        Else
            enumNext = False
        End If
    End With
End Function

'Refresh the dataset to the beginning
Public Sub enumRefresh()
    With This.Current
        .Value = Null
        .Index = 0
        .Key = Null
    End With
End Sub

'Get this enumerator as a collection
'@returns - The enumerator's data as a collection
Public Function AsCollection() As Collection
    Dim i As Long: i = 0
    Dim v As Variant: v = Null
    Dim kv As String: kv = ""
    Dim retCol As Collection: Set retCol = New Collection
    While NextItem(kv, v, i)
        retCol.Add v, kv
    Wend

    Set AsCollection = retCol
End Function

'Get this enumerator as an array
'@param iType - The vartype of the array generated 
'@returns Array<T> - The enumerator's data as an array
Public Function AsArray(Optional ByVal iType As VbVarType = vbVarType.vbVariant) As Variant
    Dim arrLen As Long: arrLen = Length()
    If arrLen = 0 Then
        AsArray = CreateEmptyVBArray(iType)
        Exit Function
    End If
    Dim vRet As Variant
    Select Case iType
        Case VbVarType.vbBoolean
            Dim a1() As Boolean
            ReDim a1(1 To arrLen)
            vRet = a1
        Case VbVarType.vbByte
            Dim a2() As Byte
            ReDim a2(1 To arrLen)
            vRet = a2
        Case VbVarType.vbCurrency
            Dim a3() As Currency
            ReDim a3(1 To arrLen)
            vRet = a3
        Case VbVarType.vbDate
            Dim a4() As Date
            ReDim a4(1 To arrLen)
            vRet = a4
        Case VbVarType.vbDouble
            Dim a5() As Double
            ReDim a5(1 To arrLen)
            vRet = a5
        Case VbVarType.vbLong
            Dim a6() As Long
            ReDim a6(1 To arrLen)
            vRet = a6
        Case VbVarType.vbObject
            Dim a7() As Object
            ReDim a7(1 To arrLen)
            vRet = a7
        Case VbVarType.vbString
            Dim a8() As String
            ReDim a8(1 To arrLen)
            vRet = a8
        Case VbVarType.vbVariant
            Dim a9() As Variant
            ReDim a9(1 To arrLen)
            vRet = a9
    End Select

    Dim v As Variant: v = Null
    Dim kv As Variant: kv = Null
    Dim i As Long: i = 0
    Select Case iType
        Case VbVarType.vbObject
            While NextItem(kv, v, i)
                Set vRet(i) = v
            Wend
        Case VbVarType.vbVariant
            While NextItem(kv, v, i)
                If IsObject(v) Then
                    Set vRet(i) = v
                Else
                    vRet(i) = v
                End If
            Wend
        Case Else
            While NextItem(kv, v, i)
                vRet(i) = v
            Wend
    End Select
    
    AsArray = vRet
End Function

'Converts an enumerator of dictionaries to a 2D array
'@returns Array2d<Variant> - 2D Array of data including headers. Number of rows is equal to the length of the enumerator, number of columns is equal to the number of keys in the first item.
Public Function AsArray2D(Optional ByVal includeObjects As Boolean = True)
    Dim vItem
    Call CopyVariant(vItem, Item(1, True))

    Dim vRet() As Variant

    Dim vKeys As Variant
    If TypeName(vItem) = "Dictionary" Then
        vKeys = vItem.keys()
        Dim iKeysLb As Long: iKeysLb = LBound(vKeys)
        Dim iKeysUb As Long: iKeysUb = UBound(vKeys)
        ReDim vRet(0 To Length, iKeysLb To iKeysUb)

        'Create keys
        Dim i As Long
        For i = iKeysLb To iKeysUb
            vRet(0, i) = vKeys(i)
        Next

        'Populate values
        Dim v As Variant: v = Null
        Dim kv As Variant: kv = Null
        i = 0
        While NextItem(kv, v, i)
            Dim j As Long
            For j = iKeysLb To iKeysUb
                If IsObject(v(vKeys(j))) Then
                    If includeObjects Then
                        Set vRet(i, j) = v(vKeys(j))
                    Else
                        vRet(i, j) = "[Object]"
                    End If
                Else
                    vRet(i, j) = v(vKeys(j))
                End If
            Next
        Wend
    Else
        Err.Raise 1, "stdEnumerator#As2DArray", "Type conversion not defined"
    End If

    'Return data
    AsArray2D = vRet
End Function

'Returns the stored enumerator as a dictionary
'@returns Object<Dictionary> - Dictionary containing enumerator data
Public Function AsDictionary() As Object
  Dim oRet As Object
  Set oRet = CreateObject("Scripting.Dictionary")
  
  Dim v As Variant: v = Null
  Dim kv As Variant: kv = Null
  Dim i As Long: i = 0
  While NextItem(kv, v, i)
    If IsObject(v) Then
      Set oRet(kv) = v
    Else
      Let oRet(kv) = v
    End If
  Wend

  Set AsDictionary = oRet
End Function

'Calls the callback on each item of the enumerator
'@param cb as stdICallable<(element: Variant)=>Void> - Callback to call for each item
'@returns - This is the enumerator itself for chaining
Public Function ForEach(ByVal cb As stdICallable) As stdEnumerator
  Dim v As Variant: v = Null
  Dim i As Long: i = 0
  Dim kv As Variant: kv = Null
  While NextItem(kv, v, i)
    Call cb.Run(v, kv)
  Wend
  Set ForEach = Me
End Function

'Call callback `cb` on each item in the enumerator. Creates a new enumerator from each value returned from `cb`.
'@param cb as stdICallable<(element: Variant)=>Variant> - Callback to map values to.
'@returns - Enumerator containing mapped values
Public Function Map(ByVal cb As stdICallable) As stdEnumerator
  Dim oRet As Collection: Set oRet = New Collection
  Dim v As Variant: v = Null
  Dim i As Long: i = 0
  Dim kv As Variant: kv = Null
  While NextItem(kv, v, i)
    Call oRet.Add(cb.Run(v, kv))
  Wend

  Set Map = stdEnumerator.CreateFromIEnumVariant(oRet)
End Function

'Returns a new enumerator containing all elements of `Me` for which the given callback returns `true`.
'@param cb as stdICallable<(element: Variant)=>Boolean> - callback to test condition for
'@returns - Enumerator containing all elements of `Me` for which the given callback returns `true`.
Public Function Filter(ByVal cb As stdICallable) As stdEnumerator
  Dim oRet As Collection: Set oRet = New Collection
  Dim v As Variant: v = Null
  Dim i As Long: i = 0
  Dim kv As Variant: kv = Null

  While NextItem(kv, v, i)
    If cb.Run(v, kv) Then Call oRet.Add(v)
  Wend

  Set Filter = stdEnumerator.CreateFromIEnumVariant(oRet)
End Function

'Sort the enumerator contents by the value retrieved by either the values in the enumerator or the values returned by `cb`.
'@param cb as stdICallable<(element: Variant)=>Variant> - If a callback provided, sorts items based on values returned by the callback. Otherwise sorts items based on values in enumerator
'@returns - enumerator of sorted values
Public Function Sort(Optional ByVal cb As stdICallable = Nothing) As stdEnumerator
    If Length > 0 Then
        Dim arr() As SortStruct
        Dim iIndex As Long: iIndex = 0
        Dim iFirstItem As Long: iFirstItem = 1
        Dim val As Variant: val = Null
        Dim keyVal As Variant: keyVal = Null
        While NextItem(keyVal, val, iIndex)
            'Increment index
            ReDim Preserve arr(1 To iIndex)

            'Bind to SortStruct
            If iIndex > 1 Then
                'Initialise sorting struct
                Call CopyVariant(arr(iIndex).value, val)
                arr(iIndex).iIndex = iIndex
                If cb Is Nothing Then
                arr(iIndex).sortValue = arr(iIndex).value
                Else
                arr(iIndex).sortValue = cb.Run(arr(iIndex).value, keyVal)
                End If

                'Sort/Compare
                Dim iCompareIndex As Long: iCompareIndex = iFirstItem

                Do While iCompareIndex <> 0
                    'If sort value at current index is less than at compare index then but this index to compare index via next
                    If arr(iIndex).sortValue < arr(iCompareIndex).sortValue Then
                        'Bind this index to compare index via iNext property
                        arr(iIndex).iNext = arr(iCompareIndex).iIndex

                        'Rebind previous element if required
                        If arr(iCompareIndex).iPrev <> 0 Then
                            'My new previous index is the previous elements previous index
                            arr(iIndex).iPrev = arr(iCompareIndex).iPrev
                            
                            'The previous elements iNext should bind to me
                            arr(arr(iIndex).iPrev).iNext = iIndex
                            
                            'The compare indexes iPrev should bind to me
                            arr(iCompareIndex).iPrev = iIndex
                        Else
                            'There is no previous element i.e. this is the first element, change iFirstItem, and link current index to iPrev of comparee
                            arr(iCompareIndex).iPrev = iIndex
                            iFirstItem = iIndex
                        End If
                        
                        'No need to carry on searching for where item should go, exit do loop
                        Exit Do
                    Else
                        'Ensure next element defined, if not then we have a new next element
                        If arr(iCompareIndex).iNext <> 0 Then
                            'Schedule next sorting check and keep searching
                            iCompareIndex = arr(iCompareIndex).iNext
                        Else
                            'Next element is not defined, therefore this is max
                            'in this case set next of arr(iCompareIndex) to this
                            'set prev of this to iCompareIndex
                            arr(iCompareIndex).iNext = iIndex
                            arr(iIndex).iPrev = iCompareIndex
                            
                            'No need to carry on searching for where item should go, exit do loop
                            Exit Do
                        End If
                    End If
                Loop
            Else
                'Initialise sorting struct
                Call CopyVariant(arr(1).value, val)
                If cb Is Nothing Then
                  arr(1).sortValue = arr(1).value
                Else
                  arr(1).sortValue = cb.Run(arr(1).value, keyVal)
                End If
                arr(1).iIndex = 1
                arr(1).iNext = 0
                arr(1).iPrev = 0
            End If
        Wend

        'Collect sorted elements
        Dim ret As Collection
        Set ret = New Collection
        Dim i As Long: i = iFirstItem
        While i <> 0
            Call ret.Add(arr(i).value)
            i = arr(i).iNext
        Wend

        'Return sorted collection as stdEnumerator
        Set Sort = stdEnumerator.CreateFromIEnumVariant(ret)
    Else
        Set Sort = Me
    End If
End Function

'Returns a new enumerator by removing duplicate values from `Me`.
'@param cb as stdICallable<(element: Variant)=>Variant> - If a callback provided, filters out duplicate items based on values returned by the callback. Otherwise filters duplicate item based on their value.
'@returns - enumerator of unique values
Public Function Unique(Optional ByVal cb As stdICallable = Nothing) As stdEnumerator
  Dim oRet As Collection: Set oRet = New Collection
  Dim v As Variant: v = Null
  Dim i As Long: i = 0
  Dim kv As Variant: kv = Null
  Dim comp1 As Variant, comp2 As Variant

  While NextItem(kv, v, i)
    Dim tv As Variant: tv = Null
    Dim ktv As Variant: ktv = Null
    Dim ti As Long: ti = i
    Dim matchFound As Boolean: matchFound = False
    
    'Obtain comparison token
    If cb Is Nothing Then
      Call CopyVariant(comp1, v)
    Else
      Call CopyVariant(comp1, cb.Run(v, kv))
    End If

    Do While NextItem(ktv, tv, ti)
        'Obtain 2nd comparison token
        If cb Is Nothing Then
          Call CopyVariant(comp2, tv)
        Else
          Call CopyVariant(comp2, cb.Run(tv, ktv))
        End If
        
        matchFound = areVariantsEqual(comp1, comp2)
        If matchFound Then Exit Do
    Loop
    
    'If a match isn't found return v
    If Not matchFound Then Call oRet.Add(v)
  Wend

  Set Unique = stdEnumerator.CreateFromIEnumVariant(oRet)
End Function

'Returns a new enumerator containing the reverse of the existing items in stdEnumerator.
'@returns - enumerator of values in reverse order
Public Function Reverse() As stdEnumerator
  Dim oRet As Collection: Set oRet = New Collection
  Dim v: v = AsArray()
  Dim i As Long:
  For i = UBound(v) To LBound(v) Step -1
    oRet.Add v(i)
  Next

  Set Reverse = stdEnumerator.CreateFromIEnumVariant(oRet)
End Function

'Concatenates to enumerators of values.
'@param obj - Enumerator to concat with that of `me`.
'@returns - A concatenated array of values.
Public Function Concat(ByVal obj As stdEnumerator) As stdEnumerator
    Dim oRet As Collection: Set oRet = New Collection
    Dim v As Variant: v = Null
    Dim i As Long: i = 0
    Dim kv As Variant: v = Null
    While NextItem(kv, v, i)
        Call oRet.Add(v)
    Wend
    For Each v In obj.AsCollection
        Call oRet.Add(v)
    Next

    Set Concat = stdEnumerator.CreateFromIEnumVariant(oRet)
End Function

'Join the elements of the enumerator joint with a delimeter
'@param sDelimiter - Enumerator to concat with that of `me`.
'@returns - A concatenated array of values.
Public Function Join(Optional ByVal sDelimiter As String = ",") As String
    Dim sRet As String
    sRet = ""

    Dim oRet As Collection: Set oRet = New Collection
    Dim v As Variant: v = Null
    Dim i As Long: i = 0
    Dim kv As Variant: kv = Null
    
    While NextItem(kv, v, i)
        sRet = sRet & sDelimiter & v
    Wend

    Join = Mid(sRet, Len(sDelimiter) + 1)
End Function

'Find the index of an element
'@param tv - Element to find.
'@returns - Index of element.
Public Function indexOf(ByVal tv As Variant) As Long
    Dim oRet As Collection: Set oRet = New Collection
    Dim v As Variant: v = Null
    Dim i As Long: i = 0
    Dim kv As Variant: kv = Null
    
    While NextItem(kv, v, i)
        If areVariantsEqual(tv, v) Then
            indexOf = i
            Exit Function
        End If
    Wend

    indexOf = 0
End Function

'Find the last index of an element in the enumerator
'@param tv - Element to find
'@returns - Last index of element.
Public Function lastIndexOf(ByVal tv As Variant) As Long
    Dim oRet As Collection: Set oRet = New Collection
    Dim v As Variant: v = Null
    Dim i As Long: i = 0
    Dim kv As Variant: kv = Null
    
    While NextItem(kv, v, i)
        If areVariantsEqual(tv, v) Then
            lastIndexOf = i
        End If
    Wend
End Function

'Find whether an element is included in the enumerator
'@param tv - Element to find
'@returns - `True` if element is found, else `false`.
Public Function includes(ByVal tv As Variant) As Boolean
    includes = indexOf(tv) > 0
End Function

'Ensure that all elements in an enumerator satisfy a condition.
'@param cb as stdICallable<(element: Variant)=>Boolean> - condition to check.
'@returns - `True` if all elements satisfy the condition, else `false`.
Public Function checkAll(ByVal cb As stdICallable) As Boolean
    Dim oRet As Collection: Set oRet = New Collection
    Dim v As Variant: v = Null
    Dim i As Long: i = 0
    Dim kv As Variant: kv = Null

    checkAll = True
    While NextItem(kv, v, i)
        If Not cb.Run(v, kv) Then
            checkAll = False
            Exit Function
        End If
    Wend
End Function

'Ensure that any of the elements in an enumerator satisfy a condition.
'@param cb as stdICallable<(element: Variant)=>Boolean> - condition to check.
'@returns - `True` if any element satisfies the condition, else `false`.
Public Function checkAny(ByVal cb As stdICallable) As Boolean
    Dim oRet As Collection: Set oRet = New Collection
    Dim v As Variant: v = Null
    Dim i As Long: i = 0
    Dim kv As Variant: kv = Null
    
    While NextItem(kv, v, i)
        If cb.Run(v, kv) Then
            checkAny = True
            Exit Function
        End If
    Wend
    checkAny = False
End Function

'Ensure that none of the elements in an enumerator satisfy a condition.
'@param cb as stdICallable<(element: Variant)=>Boolean> - condition to check.
'@returns - `True` if none of the elements satisfy the condition, else `false`.
Public Function checkNone(ByVal cb As stdICallable) As Boolean
    Dim oRet As Collection: Set oRet = New Collection
    Dim v As Variant: v = Null
    Dim i As Long: i = 0
    Dim kv As Variant: kv = Null

    While NextItem(kv, v, i)
        If cb.Run(v, kv) Then
            checkNone = False
            Exit Function
        End If
    Wend
    checkNone = True
End Function

'Ensure that only one of the elements in an enumerator satisfy a condition.
'@param cb as stdICallable<(element: Variant)=>Boolean> - condition to check.
'@returns - `True` if only one of the elements satisfy the condition, else `false`.
Public Function checkOnlyOne(ByVal cb As stdICallable) As Boolean
    Dim oRet As Collection: Set oRet = New Collection
    Dim v As Variant: v = Null
    Dim i As Long: i = 0
    Dim kv As Variant: kv = Null
    Dim iMatched As Long: iMatched = 0
    While NextItem(kv, v, i)
        If cb.Run(v, kv) Then
            iMatched = iMatched + 1
            If iMatched > 1 Then
                checkOnlyOne = False
                Exit Function
            End If
        End If
    Wend
    If iMatched = 0 Then
        checkOnlyOne = False
        Exit Function
    End If
    checkOnlyOne = True
End Function

'The reduce() method executes a reducer function (that you provide) on each element of the array, resulting in single output value.
'The reducer function takes three arguments:
'
'1. Accumultor
'2. Current Value
'3. Current Index/Key
'
'Your reducer function's returned value is assigned to the accumulator, whose value is remembered across each iteration throughout the array,
'and ultimately becomes the final, single resulting value.
'@param cb as stdICallable<(accumulator: variant, element: Variant, key: variant)=>Variant> - reducer function to execute.
'@param vInitialValue - initial value to use in reducer.
'@returns - The reduced value.
Public Function reduce(ByVal cb As stdICallable, Optional ByVal initialValue As Variant = 0) As Variant
    Dim v As Variant: v = Null
    Dim i As Long: i = 0
    Dim kv As Variant: kv = Null
    
    'Obtain initial value
    Dim vRet As Variant
    If IsMissing(initialValue) Then
        Call NextItem(kv, v, i)
        Call CopyVariant(vRet, v)
    Else
        Call CopyVariant(vRet, initialValue)
    End If
    
    'Reduce
    While NextItem(kv, v, i)
        Call CopyVariant(vRet, cb.Run(vRet, v, kv))
    Wend
    Call CopyVariant(reduce, vRet)
End Function

'Count the number of times a condition is met
'@param cb as stdICallable<(element: Variant)=>Boolean> - The callable to test for
'@returns - The number of times a condition is met
Public Function countBy(ByVal cb As stdICallable) As Long
    Dim iRet As Long: iRet = 0
    Dim v As Variant: v = Null
    Dim i As Long: i = 0
    Dim kv As Variant: kv = Null

    While NextItem(kv, v, i)
        If cb.Run(v, kv) Then iRet = iRet + 1
    Wend
    countBy = iRet
End Function

'Group elements in stdEnumerator based on the value return by a callback
'@param cb as stdICallable<(element: Variant)=>Boolean> - The callable which when run on an element returns the key to group by
'@returns - Creates a dictionary, who's keys lead to `stdEnumerator`'s of elements matching that grouping.
Public Function groupBy(ByVal cb As stdICallable) As Object
    Dim oRet As Object: Set oRet = CreateObject("Scripting.Dictionary")
    Dim key As Variant
    
    Dim v As Variant: v = Null
    Dim i As Long: i = 0
    Dim kv As Variant: kv = Null
    
    While NextItem(kv, v, i)
        Call CopyVariant(key, cb.Run(v, kv))
        If Not oRet.exists(key) Then Set oRet(key) = New Collection
        Call oRet(key).add(v)
    Wend
            
    Dim keys As Variant: keys = oRet.keys()
    For i = 0 To UBound(keys)
        Set oRet(keys(i)) = stdEnumerator.CreateFromIEnumVariant(oRet(keys(i)))
    Next
    Set groupBy = oRet
End Function

'Returns a new enumerator containing groups from groupBy, with the key and item properties added
'@param cb as stdICallable<(element: Variant)=>Variant> - The callable which when run on an element returns the key to group by
'@returns stdEnumerator<Object<Dictionary<key:string, item:variant>>> - Enumerator containing grouped data
'@example ```vb
'Dim e as stdEnumerator: set e = stdEnumerator.CreateFromArray(Array(1,2,3,2,2,1,3,2,3,1,2,3))
'Debug.Print e.groupByEx(stdLambda.Create("$1")).Map(stdLambda.Create("$1.key & ""-"" & $1.value.length")).Join(";")
'```
'@example ```vb
''Get the largest count
'Dim e as stdEnumerator: set e = stdEnumerator.CreateFromArray(Array(1,2,3,2,2,1,3,2,3,1,2,3))
'Debug.Print e.groupByEx(stdLambda.Create("$1")).Sort(stdLambda.Create("-1 * $1.value.length")).Item(1)("key")
'```
Public Function groupByEx(ByVal cb As stdICallable) As stdEnumerator
    Dim groups As Object: Set groups = groupBy(cb)
    Set groupByEx = CreateFromDictionaryEx(groups)
End Function

'Obtains the maximum value from the enumerator. If a callback is given the item which returns the largest value from the callback is returned.
'@param cb as stdICallable<(element: Variant)=>Variant> - If callback given, callback to obtain max value of. Else max of the elements values are found.
'@returns - Value containing max data
Public Function max(Optional ByVal cb As stdICallable = Nothing) As Variant
    Dim vRet As Variant
    Dim vMaxValue As Variant
    Dim v As Variant: v = Null
    Dim i As Long: i = 0
    Dim kv As Variant: kv = Null

    While NextItem(kv, v, i)
        'Get value to test
        Dim vtValue As Variant
        If cb Is Nothing Then
            Call CopyVariant(vtValue, v)
        Else
            Call CopyVariant(vtValue, cb.Run(v, kv))
        End If

        'Compare values and return 
        If IsEmpty(vRet) Then
            Call CopyVariant(vRet, v)
            Call CopyVariant(vMaxValue, vtValue)
        ElseIf vMaxValue < vtValue Then
            Call CopyVariant(vRet, v)
            Call CopyVariant(vMaxValue, vtValue)
        End If
    Wend
            
    Call CopyVariant(max, vRet)
End Function

'Obtains the minimum value from the enumerator. If a callback is given the item which returns the largest value from the callback is returned.
'@param cb as stdICallable<(element: Variant)=>Variant> -  If callback given, callback to obtain min value of. Else min of the elements values are found.
'@returns - Value containing min data
Public Function min(Optional ByVal cb As stdICallable = Nothing) As Variant
    Dim vRet As Variant
    Dim vMaxValue As Variant
    Dim v As Variant: v = Null
    Dim i As Long: i = 0
    Dim kv As Variant: kv = Null

    While NextItem(kv, v, i)
        'Get value to test
        Dim vtValue As Variant
        If cb Is Nothing Then
            Call CopyVariant(vtValue, v)
        Else
            Call CopyVariant(vtValue, cb.Run(v, kv))
        End If

        'Compare values and return
        If IsEmpty(vRet) Then
            Call CopyVariant(vRet, v)
            Call CopyVariant(vMaxValue, vtValue)
        ElseIf vMaxValue > vtValue Then
            Call CopyVariant(vRet, v)
            Call CopyVariant(vMaxValue, vtValue)
        End If
    Wend
            
    Call CopyVariant(min, vRet)
End Function

'Obtains the sum of the elements in an enumerator, or of the values returned by the callback.
'@param cb as stdICallable<(element: Variant)=>Variant> -  A callback returning an item to sum. If no callback supplied then the elements themselves will be summed
'@returns - The sum value
Public Function sum(Optional ByVal cb As stdICallable = Nothing) As Variant
    Dim vRet As Variant
    Dim oRet As Collection: Set oRet = New Collection
    Dim v As Variant: v = Null
    Dim i As Long: i = 0
    Dim kv As Variant: kv = Null
    
    While NextItem(kv, v, i)
        'Get value to test
        Dim vtValue As Variant
        If cb Is Nothing Then
            Call CopyVariant(vtValue, v)
        Else
            Call CopyVariant(vtValue, cb.Run(v, kv))
        End If
        vRet = vRet + vtValue
    Wend
            
    Call CopyVariant(sum, vRet)
End Function

'Flatten an enumerator of enumerators
'@returns - A flattened enumerator (depth = 1)
Public Function Flatten() As stdEnumerator
    Set Flatten = stdEnumerator.CreateFromIEnumVariant(flattenChildrenRecurse(Me))
End Function

'Run the callback on an all of the items i times.
'@param iTimes - Times to execute the cb on the collection
'@param cb as stdICallable<(element: variant, key: variant)=>Void> - Callback to execute on each item
'@returns  - This object.
Public Function Cycle(ByVal iTimes As Long, ByVal cb As stdICallable) As stdEnumerator
    Dim i As Long
    For i = 1 To iTimes
        Dim oRet As Collection: Set oRet = New Collection
        Dim v As Variant: v = Null
        Dim ti As Long: ti = 0
        Dim kv As Variant: kv = Null
        
        While NextItem(kv, v, ti)
            Call cb.Run(v, kv)
        Wend
    Next
    Set Cycle = Me
End Function

'Finds the first item which matches a certain condition
'@param cb as stdICallable<(element: variant, key: variant)=>Boolean> - Callback to check whether item matches condition
'@param ifMissing - Value to return if missing
'@returns - Value found or value of isMissing if missing
Public Function FindFirst(ByVal cb As stdICallable, Optional ByVal ifMissing As Variant = Null) As Variant
    Dim oRet As Collection: Set oRet = New Collection
    Dim v As Variant: v = Null
    Dim i As Long: i = 0
    Dim kv As Variant: kv = Null

    While NextItem(kv, v, i)
        If cb.Run(v, kv) Then
            Call CopyVariant(FindFirst, v)
            Exit Function
        End If
    Wend

    Call CopyVariant(FindFirst, ifMissing)
End Function

'Get the item at an index / key
'@param index   - An index / key of an element to obtain.
'@param byIndex - If true, then index of item is used, instead of potentially key.
'@returns       - Item returned at index/key
Public Property Get Item(ByVal index As Variant, Optional ByVal byIndex As Boolean = False) As Variant
    Dim oRet As Collection: Set oRet = New Collection
    Dim v As Variant: v = Null
    Dim ti As Long: ti = 0
    Dim kv As Variant: kv = Null

    'Loop through items until 
    While NextItem(kv, v, ti)
        Dim bMeetsRequirements As Boolean
        If byIndex Then
          bMeetsRequirements = ti = index
        Else
          bMeetsRequirements = areVariantsEqual(kv, index)
        End If
        
        If bMeetsRequirements Then
            Call CopyVariant(Item, v)
            Exit Property
        End If
    Wend
    
    Err.Raise 1, "stdEnumerator", "Error: No such item " & index & " in stdEnumerator."
End Property

'Get the length of the enumerator
'@returns - Length of enumerator.
Public Property Get Length() As Long
    Dim oRet As Collection: Set oRet = New Collection
    Dim v As Variant: v = Null
    Dim i As Long: i = 0
    Dim iMaxLength As Long: iMaxLength = This.MaxLength
    If This.mode = FromArray Then
        Length = UBound(This.BaseArray) - LBound(This.BaseArray) + 1
    Else
        Dim kv As Variant: kv = Null
        While NextItem(kv, v, i)
            If i > iMaxLength Then
                Err.Raise 1, "stdEnumerator", "Error: In order to prevent freezes stdEnumerator's Length() has a maxLength parameter. If the length of your collection is greater than this figure the program will error. Either Potential freeze detected."
            End If
            If i Mod 1000 = 0 Then DoEvents
        Wend
        Length = i - 1
    End If
End Property

'Get the first `iCount` elements of the collection. Useful for testing
'@param iCount - The number of elements to retrieve
'@returns - Enumerator containing the first `iCount` elements.
'@example ```vb
'stdEnumerator.CreateFromArray(Array(1,2,3,4,5,6,7,8,9)).First(3).Join(",") '=>1,2,3
'```
Public Function First(ByVal iCount As Long) As stdEnumerator
    Dim oRet As Collection: Set oRet = New Collection
    Dim v As Variant: v = Null
    Dim ti As Long: ti = 0
    Dim kv As Variant: kv = Null
    
    While NextItem(kv, v, ti)
        oRet.Add v
        
        'Reduce and check counter
        iCount = iCount - 1
        If iCount = 0 Then GoTo Jump
    Wend

Jump:
    Set First = stdEnumerator.CreateFromIEnumVariant(oRet)
End Function


'===================================================================================================
'Compares 2 variables and returns true if they are equal
'@param v1 - Variable 1
'@param v2 - Variable 2
'@returns - True if the variables are equal, false otherwise.
Private Function areVariantsEqual(ByVal v1 As Variant, ByVal v2 As Variant) As Boolean
    If IsObject(v1) And IsObject(v2) Then
        areVariantsEqual = v1 Is v2
    ElseIf (Not IsObject(v1)) And (Not IsObject(v2)) Then
        areVariantsEqual = v1 = v2
    End If
End Function

'Copy the value from vSrc into vDest
'@param vDest - Destination value
'@param vSrc  - Source value
Private Sub CopyVariant(ByRef vDest As Variant, ByVal vSrc As Variant)
    If IsObject(vSrc) Or VarType(vSrc) = VT_UNKNOWN Then
        Set vDest = vSrc
    Else
        Let vDest = vSrc
    End If
End Sub

'Check whether an object implements `IEnumVARIANT`
'@param o - Value to check
'@returns - `True` if it implements `IEnumVARIANT`, `false` otherwise
Private Function implementsIEnumVariant(ByVal o As Variant) As Boolean
    On Error GoTo ErrorOccurred
        If IsObject(o) Or VarType(o) = VT_UNKNOWN Then
            Dim x As Variant
            For Each x In o
                Exit For
            Next
            implementsIEnumVariant = True
            Exit Function
        End If
    On Error GoTo 0

ErrorOccurred:
    implementsIEnumVariant = False
End Function

'Recursively identify children based on iEnumVariant
'@param parent - `Object`/`Array` or whatever to flatten
'@returns - Flattened collection of items
Private Function flattenChildrenRecurse(ByVal parent As Variant) As Collection
    Dim oRet As Collection: Set oRet = New Collection
    Dim child As Variant, t As Variant
    If TypeName(parent) = "stdEnumerator" Then
        'Loop through children
        For Each child In parent.AsCollection
            'Recursive add to oRet
            For Each t In flattenChildrenRecurse(child)
                oRet.Add t
            Next
        Next
    ElseIf implementsIEnumVariant(parent) Then
        'Loop through children
        For Each child In parent
            'Recursive add to oRet
            For Each t In flattenChildrenRecurse(child)
                oRet.Add t
            Next
        Next
    Else
        'Parent is flat already
        oRet.Add parent
    End If
    Set flattenChildrenRecurse = oRet
End Function

'Given previous item and index, obtain the next item and index
'@param vNextKey      - Previous/Next Key
'@param vNextItem     - Previous/Next Item
'@param iCurrentIndex - Previous/Next Index
'@returns             - Whether a new item was succesfully obtained or not.
'@TODO: Use DispCallFunc to call IEnumVariant::Next instead of converting to array
Private Function NextItem(ByRef vNextKey As Variant, ByRef vNextItem As Variant, ByRef iCurrentIndex As Long) As Boolean
  iCurrentIndex = iCurrentIndex + 1
  Select Case This.mode
    Case FromIEnumVariant, FromArray
      Dim lb As Long: lb = LBound(This.BaseArray, 1)
      Dim ub As Long: ub = UBound(This.BaseArray, 1)
      Dim l As Long: l = ub - lb + 1

      If l > 0 Then
        Dim iIndex As Long: iIndex = lb + iCurrentIndex - 1
        If iIndex <= ub Then
          Call CopyVariant(vNextItem, This.BaseArray(iIndex))
          vNextKey = IIf(This.mode = FromIEnumVariant, iCurrentIndex, iIndex)
          
          NextItem = True
        Else
          NextItem = False
        End If
      Else
        NextItem = False
      End If
    Case FromCallable
      Dim vTmpItem As Variant
      Call CopyVariant(vNextItem, This.Callback.Run(vNextItem, iCurrentIndex))
      vNextKey = iCurrentIndex
      NextItem = Not IsNull(vNextItem)
    Case FromCallableVerbose
      Dim vVerboseReturn As Variant
      Call CopyVariant(vVerboseReturn, This.Callback.Run(vNextItem, iCurrentIndex))
      If IsArray(vVerboseReturn) Then
        Dim iArrLength As Long: iArrLength = UBound(vVerboseReturn) - LBound(vVerboseReturn) + 1
        If iArrLength > 0 And iArrLength <= 4 Then
          vNextItem = Null
          vNextKey = Null
          NextItem = True
          
          Dim iArrLb As Long: iArrLb = LBound(vVerboseReturn)
          If iArrLength >= 1 Then NextItem = vVerboseReturn(iArrLb + 0)
          If iArrLength >= 2 Then iCurrentIndex = vVerboseReturn(iArrLb + 1)
          If iArrLength >= 3 Then Call CopyVariant(vNextItem, vVerboseReturn(iArrLb + 2))
          If iArrLength >= 4 Then Call CopyVariant(vNextKey, vVerboseReturn(iArrLb + 3))
        Else
          Err.Raise 1, "stdEnumerator", "Incorrect return value length"
        End If
      Else
        If IsNull(vVerboseReturn) Then
          NextItem = False
        Else
          Err.Raise 1, "stdEnumerator", "Verbose callable should return an array of options: Continue?, nextIndex, nextValue, nextKey"
        End If
      End If
  End Select
End Function

'Create an enumerator from an array
'@param varType - The type of array to create
'@returns - An enumerator containing the array
Private Function CreateEmptyVBArray(Optional ByVal varType As VbVarType = vbVariant) As Variant
    #If Mac Then
        select case varType
            case vbVariant:   CreateEmptyVBArray = Array()
            case vbString:    CreateEmptyVBArray = Split("",",")
            case vbByte
                Dim b() as Byte
                b = ""
                CreateEmptyVBArray = b
            case else
                Err.Raise 1, "stdEnumerator#CreateEmpty()", "Creating Empty arrays of this type is currently unsupported on Mac"
        end select
    #Else
        'TODO: LongPtr?
        Select Case varType
            Case vbVariant:   CreateEmptyVBArray = createEmptyArrayVariant()
            Case vbString:    CreateEmptyVBArray = createEmptyArrayString()
            Case vbByte:      CreateEmptyVBArray = createEmptyArrayByte()
            Case vbBoolean:   CreateEmptyVBArray = createEmptyArrayBoolean()
            Case vbCurrency:  CreateEmptyVBArray = createEmptyArrayCurrency()
            Case vbDate:      CreateEmptyVBArray = createEmptyArrayDate()
            Case vbDouble:    CreateEmptyVBArray = createEmptyArrayDouble()
            Case vbLong:      CreateEmptyVBArray = createEmptyArrayLong()
            Case vbObject:    CreateEmptyVBArray = createEmptyArrayObject()
            #If Win64 Then
            Case vbLongLong:  CreateEmptyVBArray = createEmptyArrayLongLong()
            #End If
            Case Else
                Err.Raise 1, "stdEnumerator#CreateEmpty()", "Invalid vb vartype"
        End Select
    #End If
End Function


'TODO: Implement raw API for working with IEnumVARIANT:
Private Function getNewEnum(ByRef o As Object) As IEnumVARIANT

End Function
Private Function getNextVariant(ByRef o As IEnumVARIANT) As Variant

End Function
Private Sub skipNextVariant(ByRef o As IEnumVARIANT)

End Sub
Private Function cloneEnumVariant(ByRef o As IEnumVARIANT) As IEnumVARIANT

End Function
Private Sub resetEnumVariant(ByRef o As IEnumVARIANT)
    
End Sub

