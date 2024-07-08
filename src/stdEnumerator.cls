VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdEnumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

#if Mac then
    'Unfortunately I can't find any examples of creating an empty array on Mac :(. If you want to create Mac - compatable code don't use AsArray() with anything other than
    'vbVariant data type.
#else
    #if VBA7 then
        Private Declare PtrSafe Function createEmptyArrayBoolean     Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbBoolean  , Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) as Boolean()
        Private Declare PtrSafe Function createEmptyArrayByte        Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbByte     , Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) as Byte()
        Private Declare PtrSafe Function createEmptyArrayCurrency    Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbCurrency , Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) as Currency()
        Private Declare PtrSafe Function createEmptyArrayDate        Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbDate     , Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) as Date()
        Private Declare PtrSafe Function createEmptyArrayDouble      Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbDouble   , Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) as Double()
        Private Declare PtrSafe Function createEmptyArrayLong        Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbLong     , Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) as Long()
        Private Declare PtrSafe Function createEmptyArrayObject      Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbObject   , Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) as Object()
        Private Declare PtrSafe Function createEmptyArrayString      Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbString   , Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) as String()
        Private Declare PtrSafe Function createEmptyArrayVariant     Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbVariant  , Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) as Variant()
        #if Win64 then
            Private Declare PtrSafe Function createEmptyArrayLongLong    Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbLongLong , Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) as LongLong()
        #end if
    #else
        Private Declare Function createEmptyArrayBoolean     Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbBoolean  , Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) as Boolean()
        Private Declare Function createEmptyArrayByte        Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbByte     , Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) as Byte()
        Private Declare Function createEmptyArrayCurrency    Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbCurrency , Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) as Currency()
        Private Declare Function createEmptyArrayDate        Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbDate     , Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) as Date()
        Private Declare Function createEmptyArrayDouble      Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbDouble   , Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) as Double()
        Private Declare Function createEmptyArrayLong        Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbLong     , Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) as Long()
        Private Declare Function createEmptyArrayObject      Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbObject   , Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) as Object()
        Private Declare Function createEmptyArrayString      Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbString   , Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) as String()
        Private Declare Function createEmptyArrayVariant     Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbVariant  , Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) as Variant()
    #end if
#end if
Private Enum EnumeratorType
    FromCallable
    FromCallableVerbose
    FromIEnumVariant
    FromArray
End Enum

Private Type SortStruct
    value as variant
    sortValue as variant
    iIndex as long
    iNext as long 
    iPrev as long
End Type

Private Type TCurrentIteration
    Value as Variant
    Key as Variant
    Index as Long
End Type
Private Type TThis
    mode As EnumeratorType
    EnumObject as Object
    BaseArray as Variant
    Callback as stdICallable
    MaxLength as Long

    'Variables to allow procedural enumeration.
    Current as TCurrentIteration
End Type
Private This As TThis

'Variables to allow procedural enumeration.


Private const VT_UNKNOWN = &Hd

'Create an enumerator from a callable. Simply return either the next item, or null to indicate the end of the item set.
'@constructor
'@param cb as stdICallable<(lastItem: variant,index: long)=> variant | null> - Callback which returns either the next item.
'       If no next item, null should be returned.
'@param iMaxLength - Maximum length of enumerator. The enumerator will refuse to go beyond this number of items.
'       Used in order to prevent runaway processes.
'@returns - Enumerator created from callback function
Public Function CreateFromCallable(ByVal cb as stdICallable, optional byval iMaxLength as long = 1000000) as stdEnumerator
    set CreateFromCallable = new stdEnumerator
    Call CreateFromCallable.protInit(EnumeratorType.FromCallable,iMaxLength,cb)
End Function

'Creates an enumerator from a callable. This is more complex than non-verbose callable, however offers full flexibility of the loop.
'@constructor
'@param cb as ByVal stdICallable<(lastItem: variant,index: long)=>Array<Boolean,Long,Variant,Variant>> - Callback which returns an array. The array should be of the form:
'[hasANextItem, NextIndex, ValueToReturn, KeyToReturn]. This allows `ValueToReturn` to be null unlike Non-verbose callable.
'@param iMaxLength - Maximum length of enumerator. The enumerator will refuse to go beyond this number of items. Used in order to prevent runaway processes.
'@returns - Enumerator created from callback function
Public Function CreateFromCallableVerbose(ByVal cb as stdICallable, optional byval iMaxLength as long = 1000000) as stdEnumerator
    set CreateFromCallableVerbose = new stdEnumerator
    Call CreateFromCallableVerbose.protInit(EnumeratorType.FromCallableVerbose,iMaxLength,cb)
End Function

'Create an enumerator from an array
'@constructor
'@param v - the Array to create the enumerator from.
'@param iMaxLength - Maximum length of enumerator. The enumerator will refuse to go beyond this number of items. Used in order to prevent runaway processes.
'@returns - An enumerator representing data in array
Public Function CreateFromArray(ByVal v as variant, optional byval iMaxLength as long = 1000000) as stdEnumerator
    set CreateFromArray = new stdEnumerator
    Call CreateFromArray.protInit(EnumeratorType.FromArray,iMaxLength,v)
End Function

'Create an enumerator from an object which supports IEnumVARIANT.
'@constructor
'@param o - the object to create the enumerator from. This object should support IEnumVARIANT
'@param iMaxLength - Maximum length of enumerator. The enumerator will refuse to go beyond this number of items. Used in order to prevent runaway processes.
'@returns - An enumerator representing variants in the IEnumVARIANT
Public Function CreateFromIEnumVariant(ByVal o as Object, optional byval iMaxLength as long = 1000000) as stdEnumerator
    set CreateFromIEnumVariant = new stdEnumerator
    Call CreateFromIEnumVariant.protInit(EnumeratorType.FromIEnumVariant,iMaxLength,o)
End Function


'Create an empty enumerator
'@constructor
'@param iMaxLength - Maximum length of enumerator. The enumerator will refuse to go beyond this number of items. Used in order to prevent runaway processes.
'@returns - An empty enumerator
Public Function CreateEmpty(optional byval iMaxLength as long = 1000000) as stdEnumerator
    set CreateEmpty = new stdEnumerator
    Call CreateEmpty.protInit(EnumeratorType.FromIEnumVariant,iMaxLength,new Collection)
End Function

'Create an enumerator from an Excel ListObject. Dictionaries are created for each ListRow with Header values as keys.
'@constructor
'@param lo - ListObject to create enumerator from.
'@param iMaxLength - Maximum length of enumerator. The enumerator will refuse to go beyond this number of items. Used in order to prevent runaway processes.
'@returns - Enumerator representing ListObject data
Public Function CreateFromListObject(ByVal lo As ListObject, Optional ByVal iMaxLength As Long = 1000000, Optional ByVal dictCompare as Long = vbTextCompare) As stdEnumerator
    Dim c As Collection: Set c = New Collection
    If Not lo.DataBodyRange Is Nothing Then
        'Get header and body data
        Dim oRow As Object
        Dim vHeader: vHeader = lo.HeaderRowRange.value
        Dim vData: vData = lo.DataBodyRange.value

        If TypeName(vData) = "Variant()" Then
            'Create rows for all rows
            Dim i As Long
            For i = 1 To UBound(vData, 1)
                Set oRow = CreateObject("Scripting.Dictionary")
                oRow.CompareMode = dictCompare
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
Private Function CreateFromDictionary(ByVal dictionary as object, optional byval iMaxLength as long = 1000000) as stdEnumerator
    
End Function

'Translates a dictionary to a Enumerator of dictionaries with key and value properties.
'@constructor
'@param dictionary - The dictionary to create the enumerator from.
'@param iMaxLength - Maximum length of enumerator. The enumerator will refuse to go beyond this number of items. Used in order to prevent runaway processes.
'@returns stdEnumerator<Object<Scripting.Dictionary<key: string, value: variant>>> - An enumerator representing the dictionary data
'@example ```
'set dict = stdEnumerator.CreateFromArray(Array(1,2,3,2,2,1,3,2,3,1,2,3)).groupBy(stdLambda.Create("$1"))
'set dictEx = stdEnumerator.CreateFromDictionaryEx(dict).forEach(stdLambda.Create("let $1.value = $1.value.count")).sort(stdLambda.Create("$1.count")).last
'```
Public Function CreateFromDictionaryEx(ByVal dictionary as object) as stdEnumerator
    Dim keys as variant: keys = dictionary.keys()
    Dim items as variant: items = dictionary.items()
    Dim rows as collection: set rows = new collection
    For i = 0 to ubound(keys)
        Dim row as object: set row = CreateObject("Scripting.Dictionary")
        row.CompareMode = vbTextCompare
        row.add "key", keys(i)
        row.add "value", items(i)
        rows.add row
    next
    set CreateFromDictionaryEx = stdEnumerator.CreateFromIEnumVariant(rows)
End Function

'Initialse the stdEnumerator
'@constructor
'@protected
'@param iEnumeratorType - Enumerator type to create
'@param iMaxLength - Maximum length of enumerator. The enumerator will refuse to go beyond this number of items. Used in order to prevent runaway processes.
'@param v - The data to create the enumerator from. This is either an array, a callable, or an IEnumVARIANT object.
Public Sub protInit(ByVal iEnumeratorType as long, ByVal iMaxLength as long, ParamArray v() As Variant)
    this.mode = iEnumeratorType
    this.MaxLength = iMaxLength
    select case this.mode
        case EnumeratorType.FromIEnumVariant
            'HACK: This is a real work around for the fact that we can't access IEnumVARIANT methods (getNewEnum,getNextVariant,skipNextVariant,cloneEnumVariant,resetEnumVariant)
            'directly. Realistically this may sometimes not even work. A more long term solution would be to call getNewEnum() and then getNextVariant() until we reach the end of the
            'object. This may be implementable after stdCOM is fully implemented.

            set this.EnumObject = v(0)
            
            'Calculate total length
            Dim iLength As Long: iLength = 0
            Dim obj As Variant
            For Each obj In this.EnumObject
                iLength = iLength + 1
            next
            
            'Define array
            if iLength > 0 then
                Dim vArr()
                ReDim Preserve vArr(1 To iLength)
                
                'Copy IEnumVARIANT into array
                Dim i As Long: i = 0
                For Each obj In this.EnumObject
                    i = i + 1
                    Call CopyVariant(vArr(i), obj)
                Next

                this.BaseArray = vArr
            else
                this.BaseArray = Array()
            end if
        case EnumeratorType.FromArray
            this.BaseArray = v(0)
        case EnumeratorType.FromCallable
            set this.Callback   = v(0)
        case EnumeratorType.FromCallableVerbose
            set this.Callback   = v(0)
    end select
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
'@example ```
'Dim v as Variant
'While enumNext(v)
'  Debug.Print v
'Wend
'```
Public Function enumNext(ByRef nextValue As Variant) As Boolean
    With this.Current
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
    With this.Current
        .Value = Null
        .Index = 0
        .Key = Null
    End With
End Sub

'Get this enumerator as a collection
'@returns - The enumerator's data as a collection
Public Function AsCollection() as Collection
    Dim i as long: i = 0
    Dim v as variant: v = null
    Dim kv as string: kv = ""
    Dim retCol as Collection: set retCol = new Collection
    While NextItem(kv, v, i)
        retCol.add v, kv
    Wend

    set AsCollection = retCol
End Function

'Get this enumerator as an array
'@param iType - The vartype of the array generated 
'@returns Array<T> - The enumerator's data as an array
Public Function AsArray(Optional ByVal iType as vbVarType = vbVarType.vbVariant) as Variant
    Dim arrLen as long: arrLen = Length()
    If arrLen = 0 Then
        AsArray = CreateEmptyVBArray(iType)
        Exit Function
    End If
    Dim vRet as variant
    select case iType
        case vbVarType.vbBoolean
            Dim a1() as Boolean
            Redim a1(1 to arrLen)
            vRet = a1
        case vbVarType.vbByte
            Dim a2() as Byte
            Redim a2(1 to arrLen)
            vRet = a2
        case vbVarType.vbCurrency
            Dim a3() as Currency
            Redim a3(1 to arrLen)
            vRet = a3
        case vbVarType.vbDate
            Dim a4() as Date
            Redim a4(1 to arrLen)
            vRet = a4
        case vbVarType.vbDouble
            Dim a5() as Double
            Redim a5(1 to arrLen)
            vRet = a5
        case vbVarType.vbLong
            Dim a6() as Long
            Redim a6(1 to arrLen)
            vRet = a6
        case vbVarType.vbObject
            Dim a7() as Object
            Redim a7(1 to arrLen)
            vRet = a7
        case vbVarType.vbString 
            Dim a8() as String
            Redim a8(1 to arrLen)
            vRet = a8
        case vbVarType.vbVariant
            Dim a9() as Variant
            Redim a9(1 to arrLen)
            vRet = a9
    end select

    Dim v as variant: v = null
    Dim kv as variant: kv = null 
    Dim i as long: i = 0
    select case iType
        case vbVarType.vbObject
            while NextItem(kv, v, i)
                set vRet(i) = v
            Wend
        case vbVarType.vbVariant
            while NextItem(kv, v, i)
                if isObject(v) then
                    set vRet(i) = v
                else
                    vRet(i) = v
                end if
            Wend
        case else
            while NextItem(kv, v, i)
                vRet(i) = v
            Wend
    end select
    
    AsArray = vRet
End Function

'Converts an enumerator of dictionaries to a 2D array
'@returns Array2d<Variant> - 2D Array of data including headers. Number of rows is equal to the length of the enumerator, number of columns is equal to the number of keys in the first item.
Public Function AsArray2D(Optional ByVal includeObjects as boolean = true)
    Dim vItem
    Call CopyVariant(vItem, item(1,true))

    Dim vRet() as variant

    Dim vKeys as variant
    if typename(vItem) = "Dictionary" then
        vKeys = vItem.keys()
        Dim iKeysLb as long: iKeysLb = lbound(vKeys)
        Dim iKeysUb as long: iKeysUb = ubound(vKeys)
        Redim vRet(0 to length,iKeysLb to iKeysUb)

        'Create keys
        Dim i as long
        for i = iKeysLb to iKeysUb
            vRet(0, i) = vKeys(i)
        next

        'Populate values
        Dim v as variant: v = null
        Dim kv as variant: kv = null 
        i = 0
        While NextItem(kv,v,i)
            Dim j as long
            for j = iKeysLb to iKeysUb
                if isObject(v(vKeys(j))) then
                    if includeObjects then
                        set vRet(i,j) = v(vKeys(j))
                    else
                        vRet(i,j) = "[Object]"
                    end if
                else
                    vRet(i,j) = v(vKeys(j))
                end if
            next
        Wend
    else
        Err.Raise 1, "stdEnumerator#As2DArray", "Type conversion not defined"
    end if

    'Return data
    AsArray2D = vRet
End Function

'Returns the stored enumerator as a dictionary
'@returns Object<Dictionary> - Dictionary containing enumerator data
Public Function AsDictionary() as Object
  Dim oRet as object
  set oRet = CreateObject("Scripting.Dictionary")
  
  Dim v as variant: v = null
  Dim kv as variant: kv = null 
  Dim i as long: i = 0
  While NextItem(kv,v,i)
    if isObject(v) then
      set oRet(kv) = v
    else
      let oRet(kv) = v
    end if
  Wend

  set AsDictionary = oRet
End Function

'Calls the callback on each item of the enumerator
'@param cb as stdICallable<(element: Variant)=>Void> - Callback to call for each item
'@returns - This is the enumerator itself for chaining
Public Function ForEach(ByVal cb As stdICallable) As stdEnumerator
  Dim v as variant: v = null
  Dim i as long: i = 0
  Dim kv as variant: kv = null 
  while NextItem(kv,v,i)
    Call cb.Run(v, kv)
  Wend
  set ForEach = me
End Function

'Call callback `cb` on each item in the enumerator. Creates a new enumerator from each value returned from `cb`.
'@param cb as stdICallable<(element: Variant)=>Variant> - Callback to map values to.
'@returns - Enumerator containing mapped values
Public Function Map(ByVal cb As stdICallable) As stdEnumerator
  Dim oRet as Collection: Set oRet = new Collection
  Dim v as variant: v = null
  Dim i as long: i = 0
  Dim kv as variant: kv = null 
  while NextItem(kv,v,i)
    Call oRet.add(cb.Run(v,kv))
  Wend

  set Map = stdEnumerator.CreateFromIEnumVariant(oRet)
End Function

'Returns a new enumerator containing all elements of `Me` for which the given callback returns `true`.
'@param cb as stdICallable<(element: Variant)=>Boolean> - callback to test condition for
'@returns - Enumerator containing all elements of `Me` for which the given callback returns `true`.
Public Function Filter(ByVal cb as stdICallable) as stdEnumerator
  Dim oRet as Collection: Set oRet = new Collection
  Dim v as variant: v = null
  Dim i as long: i = 0
  Dim kv as variant: kv = null 

  While NextItem(kv, v,i)
    if cb.Run(v,kv) then Call oRet.add(v)
  Wend

  set Filter = stdEnumerator.CreateFromIEnumVariant(oRet)
End Function

'Sort the enumerator contents by the value retrieved by either the values in the enumerator or the values returned by `cb`.
'@param cb as stdICallable<(element: Variant)=>Variant> - If a callback provided, sorts items based on values returned by the callback. Otherwise sorts items based on values in enumerator
'@returns - enumerator of sorted values
Public Function Sort(Optional ByVal cb as stdICallable = nothing) as stdEnumerator
    If Length > 0 then
        Dim arr() as SortStruct
        Dim iIndex as Long: iIndex = 0
        Dim iFirstItem as long: iFirstItem = 1
        Dim val as variant: val = null
        Dim keyVal as variant: keyVal = null
        While NextItem(keyVal, val,iIndex)
            'Increment index
            Redim Preserve arr(1 to iIndex)

            'Bind to SortStruct
            if iIndex > 1 Then
                'Initialise sorting struct
                Call CopyVariant(arr(iIndex).value, val)
                arr(iIndex).iIndex = iIndex
                if cb is nothing then
                arr(iIndex).sortValue = arr(iIndex).value
                else
                arr(iIndex).sortValue = cb.Run(arr(iIndex).value, keyVal)
                end if

                'Sort/Compare
                Dim iCompareIndex as long: iCompareIndex = iFirstItem

                Do While iCompareIndex <> 0
                    'If sort value at current index is less than at compare index then but this index to compare index via next
                    if arr(iIndex).sortValue < arr(iCompareIndex).sortValue then
                        'Bind this index to compare index via iNext property
                        arr(iIndex).iNext = arr(iCompareIndex).iIndex

                        'Rebind previous element if required
                        if arr(iCompareIndex).iPrev <> 0 then
                            'My new previous index is the previous elements previous index
                            arr(iIndex).iPrev = arr(iCompareIndex).iPrev
                            
                            'The previous elements iNext should bind to me
                            arr(arr(iIndex).iPrev).iNext = iIndex
                            
                            'The compare indexes iPrev should bind to me
                            arr(iCompareIndex).iPrev = iIndex
                        else
                            'There is no previous element i.e. this is the first element, change iFirstItem, and link current index to iPrev of comparee
                            arr(iCompareIndex).iPrev = iIndex
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
                if cb is nothing then
                  arr(1).sortValue = arr(1).value
                else
                  arr(1).sortValue = cb.Run(arr(1).value, keyVal)
                end if
                arr(1).iIndex = 1
                arr(1).iNext = 0
                arr(1).iPrev = 0
            end if
        Wend

        'Collect sorted elements
        Dim ret as Collection
        set ret = new Collection
        Dim i as long: i = iFirstItem
        While i <> 0
            Call ret.add(arr(i).value)
            i = arr(i).iNext
        Wend

        'Return sorted collection as stdEnumerator
        set sort = stdEnumerator.CreateFromIEnumVariant(ret)
    else
        set sort = Me
    end if
End Function

'Returns a new enumerator by removing duplicate values from `Me`.
'@param cb as stdICallable<(element: Variant)=>Variant> - If a callback provided, filters out duplicate items based on values returned by the callback. Otherwise filters duplicate item based on their value.
'@returns - enumerator of unique values
Public Function Unique(optional byval cb as stdICallable = nothing) as stdEnumerator
  Dim oRet as Collection: Set oRet = new Collection
  Dim v as variant: v = null
  Dim i as long: i = 0
  Dim kv as variant: kv = null
  Dim comp1 as variant, comp2 as variant

  While NextItem(kv, v,i)
    Dim tv as variant: tv = null
    Dim ktv as variant: ktv = null
    Dim ti as long: ti = i
    Dim matchFound as boolean: matchFound=false
    
    'Obtain comparison token
    if cb is nothing then 
      Call CopyVariant(Comp1, v)
    else
      Call CopyVariant(Comp1, cb.Run(v, kv))
    end if

    Do While NextItem(ktv, tv, ti)
        'Obtain 2nd comparison token
        if cb is nothing then
          Call CopyVariant(Comp2, tv)
        else
          Call CopyVariant(Comp2, cb.Run(tv, ktv))
        end if
        
        matchFound = areVariantsEqual(Comp1,Comp2)
        if matchFound then Exit Do
    Loop
    
    'If a match isn't found return v
    if not matchFound then Call oRet.add(v)
  Wend

  set Unique = stdEnumerator.CreateFromIEnumVariant(oRet)
End Function

'Returns a new enumerator containing the reverse of the existing items in stdEnumerator.
'@returns - enumerator of values in reverse order
Public Function Reverse() as stdEnumerator
  Dim oRet as collection: set oRet = new Collection
  Dim v: v = AsArray()
  Dim i as long:
  For i = ubound(v) to lbound(v) step -1
    oRet.add v(i)
  next

  set Reverse = stdEnumerator.CreateFromIEnumVariant(oRet)
End Function

'Concatenates to enumerators of values.
'@param obj - Enumerator to concat with that of `me`.
'@returns - A concatenated array of values.
Public Function Concat(ByVal obj as stdEnumerator) as stdEnumerator
    Dim oRet as Collection: Set oRet = new Collection
    Dim v as variant: v = null
    Dim i as long: i = 0
    Dim kv as variant: v = null
    While NextItem(kv,v,i)
        Call oRet.add(v)
    Wend
    for each v in obj.AsCollection
        Call oRet.add(v)
    next

    set Concat = stdEnumerator.CreateFromIEnumVariant(oRet)
End Function

'Join the elements of the enumerator joint with a delimeter
'@param sDelimiter - Enumerator to concat with that of `me`.
'@returns - A concatenated array of values.
Public Function Join(Optional ByVal sDelimiter as string = ",") as string
    Dim sRet as string
    sRet = ""

    Dim oRet as Collection: Set oRet = new Collection
    Dim v as variant: v = null
    Dim i as long: i = 0
    Dim kv as variant: kv = null
    
    While NextItem(kv,v,i)
        sRet = sRet & sDelimiter & v
    Wend

    Join = mid(sRet,len(sDelimiter)+1)
End Function

'Find the index of an element
'@param tv - Element to find.
'@returns - Index of element.
Public Function indexOf(ByVal tv as variant) as long
    Dim oRet as Collection: Set oRet = new Collection
    Dim v as variant: v = null
    Dim i as long: i = 0
    Dim kv as variant: kv = null
    
    While NextItem(kv,v,i)
        if areVariantsEqual(tv,v) then
            indexOf = i
            Exit Function
        end if
    Wend

    indexOf = 0
End Function

'Find the last index of an element in the enumerator
'@param tv - Element to find
'@returns - Last index of element.
Public Function lastIndexOf(ByVal tv as variant) as long
    Dim oRet as Collection: Set oRet = new Collection
    Dim v as variant: v = null
    Dim i as long: i = 0
    Dim kv as variant: kv = null
    
    While NextItem(kv,v,i)
        if areVariantsEqual(tv,v) then
            lastIndexOf = i
        end if
    Wend
End Function

'Find whether an element is included in the enumerator
'@param tv - Element to find
'@returns - `True` if element is found, else `false`.
Public Function includes(ByVal tv as variant) as boolean
    includes = indexOf(tv) > 0
End Function

'Ensure that all elements in an enumerator satisfy a condition.
'@param cb as stdICallable<(element: Variant)=>Boolean> - condition to check.
'@returns - `True` if all elements satisfy the condition, else `false`.
Public Function checkAll(ByVal cb as stdICallable) as boolean
    Dim oRet as Collection: Set oRet = new Collection
    Dim v as variant: v = null
    Dim i as long: i = 0
    Dim kv as variant: kv = null

    checkAll = true
    While NextItem(kv,v,i)
        if not cb.Run(v,kv) then
            checkAll = false
            Exit Function
        end if
    Wend
End Function

'Ensure that any of the elements in an enumerator satisfy a condition.
'@param cb as stdICallable<(element: Variant)=>Boolean> - condition to check.
'@returns - `True` if any element satisfies the condition, else `false`.
Public Function checkAny(ByVal cb as stdICallable) as boolean
    Dim oRet as Collection: Set oRet = new Collection
    Dim v as variant: v = null
    Dim i as long: i = 0
    Dim kv as variant: kv = null
    
    While NextItem(kv,v,i)
        if cb.Run(v,kv) then
            checkAny = true
            Exit Function
        end if
    Wend
    checkAny = false
End Function

'Ensure that none of the elements in an enumerator satisfy a condition.
'@param cb as stdICallable<(element: Variant)=>Boolean> - condition to check.
'@returns - `True` if none of the elements satisfy the condition, else `false`.
Public Function checkNone(ByVal cb as stdICallable) as boolean
    Dim oRet as Collection: Set oRet = new Collection
    Dim v as variant: v = null
    Dim i as long: i = 0
    Dim kv as variant: kv = null

    While NextItem(kv, v,i)
        if cb.Run(v,kv) then
            checkNone = false
            Exit Function
        end if
    Wend
    checkNone = true
End Function

'Ensure that only one of the elements in an enumerator satisfy a condition.
'@param cb as stdICallable<(element: Variant)=>Boolean> - condition to check.
'@returns - `True` if only one of the elements satisfy the condition, else `false`.
Public Function checkOnlyOne(ByVal cb as stdICallable) as boolean
    Dim oRet as Collection: Set oRet = new Collection
    Dim v as variant: v = null
    Dim i as long: i = 0
    Dim kv as variant: kv = null
    Dim iMatched as long: iMatched = 0
    While NextItem(kv, v, i)
        if cb.Run(v,kv) then
            iMatched = iMatched + 1
            if iMatched > 1 then
                checkOnlyOne = false
                Exit Function
            end if
        end if
    Wend
    if iMatched = 0 then
        checkOnlyOne = false
        Exit Function
    End if
    checkOnlyOne = true
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
Public Function reduce(ByVal cb as stdICallable, Optional ByVal initialValue as variant = 0) as variant
    Dim v as variant: v = null
    Dim i as long: i = 0
    Dim kv as variant: kv = null
    
    'Obtain initial value
    Dim vRet as variant
    if isMissing(initialValue) then 
        Call NextItem(kv, v,i)
        Call CopyVariant(vRet, v)
    else
        Call CopyVariant(vRet, initialValue)
    end if
    
    'Reduce
    While NextItem(kv, v,i)
        Call CopyVariant(vRet, cb.Run(vRet,v,kv))
    Wend
    Call CopyVariant(reduce,vRet)
End Function

'Count the number of times a condition is met
'@param cb as stdICallable<(element: Variant)=>Boolean> - The callable to test for
'@returns - The number of times a condition is met
Public Function countBy(ByVal cb as stdICallable) as long
    Dim iRet as Long: iRet = 0
    Dim v as variant: v = null
    Dim i as long: i = 0
    Dim kv as variant: kv = null

    While NextItem(kv, v, i)
        if cb.Run(v,kv) then iRet = iRet + 1
    Wend
    countBy = iRet
End Function

'Group elements in stdEnumerator based on the value return by a callback
'@param cb as stdICallable<(element: Variant)=>Boolean> - The callable which when run on an element returns the key to group by
'@returns - Creates a dictionary, who's keys lead to `stdEnumerator`'s of elements matching that grouping.
Public Function groupBy(ByVal cb as stdICallable) as object
    Dim oRet as Object: set oRet = CreateObject("Scripting.Dictionary")
    Dim key as variant
    
    Dim v as variant: v = null
    Dim i as long: i = 0
    Dim kv as variant: kv = null
    
    While NextItem(kv,v,i)
        Call CopyVariant(key,cb.Run(v,kv))
        if not oRet.exists(key) then set oRet(key) = new Collection
        Call oRet(key).add(v)
    Wend
            
    Dim keys as variant: keys = oRet.keys()
    For i = 0 to ubound(keys)
        set oRet(keys(i)) = stdEnumerator.CreateFromIEnumVariant(oRet(keys(i)))
    next
    set groupBy = oRet
End Function

'Returns a new enumerator containing groups from groupBy, with the key and item properties added
'@param cb as stdICallable<(element: Variant)=>Variant> - The callable which when run on an element returns the key to group by
'@returns stdEnumerator<Object<Dictionary<key:string, item:variant>>> - Enumerator containing grouped data
'@example ```
'Dim e as stdEnumerator: set e = stdEnumerator.CreateFromArray(Array(1,2,3,2,2,1,3,2,3,1,2,3))
'Debug.Print e.groupByEx(stdLambda.Create("$1")).Map(stdLambda.Create("$1.key & ""-"" & $1.value.length")).Join(";")
'```
'@example ```
''Get the largest count
'Dim e as stdEnumerator: set e = stdEnumerator.CreateFromArray(Array(1,2,3,2,2,1,3,2,3,1,2,3))
'Debug.Print e.groupByEx(stdLambda.Create("$1")).Sort(stdLambda.Create("-1 * $1.value.length")).Item(1)("key")
'```
Public Function groupByEx(ByVal cb as stdICallable) as stdEnumerator
    Dim groups as object: set groups = groupBy(cb)
    set groupByEx = CreateFromDictionaryEx(groups)
End Function

'Obtains the maximum value from the enumerator. If a callback is given the item which returns the largest value from the callback is returned.
'@param cb as stdICallable<(element: Variant)=>Variant> - If callback given, callback to obtain max value of. Else max of the elements values are found.
'@returns - Value containing max data
Public Function max(Optional ByVal cb as stdICallable = nothing) as variant
    Dim vRet as variant
    Dim vMaxValue as variant
    Dim v as variant: v = null
    Dim i as long: i = 0
    Dim kv as variant: kv = null

    While NextItem(kv,v,i)
        'Get value to test
        Dim vtValue as variant
        if cb is nothing then
            Call CopyVariant(vtValue,v)
        else
            Call CopyVariant(vtValue,cb.Run(v, kv))
        end if

        'Compare values and return 
        if isEmpty(vRet) then
            Call CopyVariant(vRet,v)
            Call CopyVariant(vMaxValue, vtValue)
        elseif vMaxValue < vtValue then
            Call CopyVariant(vRet,v)
            Call CopyVariant(vMaxValue, vtValue)
        end if
    Wend
            
    Call CopyVariant(max,vRet)
End Function

'Obtains the minimum value from the enumerator. If a callback is given the item which returns the largest value from the callback is returned.
'@param cb as stdICallable<(element: Variant)=>Variant> -  If callback given, callback to obtain min value of. Else min of the elements values are found.
'@returns - Value containing min data
Public Function min(Optional ByVal cb as stdICallable = nothing) as variant
    Dim vRet as variant
    Dim vMaxValue as variant
    Dim v as variant: v = null
    Dim i as long: i = 0
    Dim kv as variant: kv = null

    While NextItem(kv,v,i)
        'Get value to test
        Dim vtValue as variant
        if cb is nothing then
            Call CopyVariant(vtValue,v)
        else
            Call CopyVariant(vtValue,cb.Run(v,kv))
        end if

        'Compare values and return
        if isEmpty(vRet) then
            Call CopyVariant(vRet,v)
            Call CopyVariant(vMaxValue, vtValue)
        elseif vMaxValue > vtValue then
            Call CopyVariant(vRet,v)
            Call CopyVariant(vMaxValue, vtValue)
        end if
    Wend
            
    Call CopyVariant(min,vRet)
End Function

'Obtains the sum of the elements in an enumerator, or of the values returned by the callback.
'@param cb as stdICallable<(element: Variant)=>Variant> -  A callback returning an item to sum. If no callback supplied then the elements themselves will be summed
'@returns - The sum value
Public Function sum(Optional ByVal cb as stdICallable = nothing) as variant
    Dim vRet as variant
    Dim oRet as Collection: Set oRet = new Collection
    Dim v as variant: v = null
    Dim i as long: i = 0
    Dim kv as variant: kv = null
    
    While NextItem(kv,v,i)
        'Get value to test
        Dim vtValue as variant
        if cb is nothing then
            Call CopyVariant(vtValue,v)
        else
            Call CopyVariant(vtValue,cb.Run(v,kv))
        end if
        vRet = vRet + vtValue
    Wend
            
    Call CopyVariant(sum,vRet)
End Function

'Flatten an enumerator of enumerators
'@returns - A flattened enumerator (depth = 1)
Public Function Flatten() as stdEnumerator
    set Flatten = stdEnumerator.CreateFromIEnumVariant(flattenChildrenRecurse(me))
End Function

'Run the callback on an all of the items i times.
'@param iTimes - Times to execute the cb on the collection
'@param cb as stdICallable<(element: variant, key: variant)=>Void> - Callback to execute on each item
'@returns  - This object.
Public Function Cycle(ByVal iTimes as long, ByVal cb as stdICallable) as stdEnumerator
    Dim i as long
    for i = 1 to iTimes
        Dim oRet as Collection: Set oRet = new Collection
        Dim v as variant: v = null
        Dim ti as long: ti = 0
        Dim kv as variant: kv = null
        
        While NextItem(kv,v,ti)
            Call cb.run(v,kv)
        Wend
    next
    set Cycle = Me
End Function

'Finds the first item which matches a certain condition
'@param cb as stdICallable<(element: variant, key: variant)=>Boolean> - Callback to check whether item matches condition
'@param ifMissing - Value to return if missing
'@returns - Value found or value of isMissing if missing
Public Function FindFirst(ByVal cb as stdICallable, optional ByVal ifMissing as variant = null) as variant
    Dim oRet as Collection: Set oRet = new Collection
    Dim v as variant: v = null
    Dim i as long: i = 0
    Dim kv as variant: kv = null

    While NextItem(kv,v,i)
        if cb.run(v,kv) then
            Call CopyVariant(FindFirst,v)
            Exit Function
        end if
    Wend

    Call CopyVariant(FindFirst, ifMissing)
End Function

'Get the item at an index / key
'@param index   - An index / key of an element to obtain.
'@param byIndex - If true, then index of item is used, instead of potentially key.
'@returns       - Item returned at index/key
Public Property Get Item(ByVal index as Variant, optional byval byIndex as boolean = false) as variant
    Dim oRet as Collection: Set oRet = new Collection
    Dim v as variant: v = null
    Dim ti as long: ti = 0
    Dim kv as variant: kv = null

    'Loop through items until 
    While NextItem(kv,v,ti)
        Dim bMeetsRequirements as boolean
        if byIndex then
          bMeetsRequirements = ti = index
        else
          bMeetsRequirements = areVariantsEqual(kv, index)
        end if
        
        if bMeetsRequirements then
            Call CopyVariant(Item,v)
            Exit Function
        end if
    Wend
    
    Err.Raise 1, "stdEnumerator", "Error: No such item " & index & " in stdEnumerator."
End Property 

'Get the length of the enumerator
'@returns - Length of enumerator.
Public Property Get Length() as Long
    Dim oRet as Collection: Set oRet = new Collection
    Dim v as variant: v = null
    Dim i as long: i = 0
    Dim iMaxLength as Long: iMaxLength = this.MaxLength
    if this.mode = FromArray then
        Length = UBound(this.BaseArray) - LBound(this.BaseArray) + 1
    else
        Dim kv as variant: kv = null
        While NextItem(kv,v,i)
            if i > iMaxLength then
                Err.Raise 1, "stdEnumerator", "Error: In order to prevent freezes stdEnumerator's Length() has a maxLength parameter. If the length of your collection is greater than this figure the program will error. Either Potential freeze detected."
            end if
            if i mod 1000 = 0 then DoEvents
        Wend
        Length = i-1
    End if
End Property

'Get the first `iCount` elements of the collection. Useful for testing
'@param iCount - The number of elements to retrieve
'@returns - Enumerator containing the first `iCount` elements.
'@example `stdEnumerator.CreateFromArray(Array(1,2,3,4,5,6,7,8,9)).First(3).Join(",") '=>1,2,3`
Public Function First(ByVal iCount As Long) As stdEnumerator
    Dim oRet As collection: Set oRet = New collection
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
Private Function areVariantsEqual(ByVal v1 as variant, ByVal v2 as variant) as boolean
    if isObject(v1) and isObject(v2) then
        areVariantsEqual = v1 is v2
    elseif (not isobject(v1)) and (not isObject(v2)) then
        areVariantsEqual = v1 = v2
    end if
End Function

'Copy the value from vSrc into vDest
'@param vDest - Destination value
'@param vSrc  - Source value
Private Sub CopyVariant(ByRef vDest as variant, ByVal vSrc as Variant)
    if isObject(vSrc)  Or  VarType(vSrc) = VT_UNKNOWN Then
        Set vDest = vSrc
    else
        Let vDest = vSrc
    end if
End Sub

'Check whether an object implements `IEnumVARIANT`
'@param o - Value to check
'@returns - `True` if it implements `IEnumVARIANT`, `false` otherwise
Private Function implementsIEnumVariant(ByVal o as Variant) as boolean
    On Error GoTo ErrorOccurred
        if isObject(o) or VarType(o) = VT_UNKNOWN then
            Dim x as variant
            For each x in o
                Exit For
            Next
            implementsIEnumVariant = true
            Exit Function
        end if
    On Error GoTo 0

ErrorOccurred:
    implementsIEnumVariant=false
End Function

'Recursively identify children based on iEnumVariant
'@param parent - `Object`/`Array` or whatever to flatten
'@returns - Flattened collection of items
Private Function flattenChildrenRecurse(ByVal parent as variant) as Collection
    Dim oRet as Collection: set oRet = new Collection
    Dim child as variant, t as variant
    if typename(parent) = "stdEnumerator" then
        'Loop through children
        for each child in parent.AsCollection
            'Recursive add to oRet
            for each t in flattenChildrenRecurse(child)
                oRet.Add t
            next
        next 
    elseif implementsIEnumVariant(parent) then
        'Loop through children
        for each child in parent
            'Recursive add to oRet
            for each t in flattenChildrenRecurse(child)
                oRet.Add t
            next
        next
    else
        'Parent is flat already
        oRet.add parent
    end if
    set flattenChildrenRecurse = oRet
End Function

'Given previous item and index, obtain the next item and index
'@param vNextKey      - Previous/Next Key
'@param vNextItem     - Previous/Next Item
'@param iCurrentIndex - Previous/Next Index
'@returns             - Whether a new item was succesfully obtained or not.
'@TODO: Use DispCallFunc to call IEnumVariant::Next instead of converting to array
Private Function NextItem(ByRef vNextKey as variant, ByRef vNextItem as Variant, ByRef iCurrentIndex as long) as Boolean
  iCurrentIndex = iCurrentIndex+1
  select case this.mode
    case FromIEnumVariant, FromArray
      Dim lb as long: lb = LBound(this.BaseArray,1)
      Dim ub as long: ub = Ubound(this.BaseArray,1)
      Dim l as long: l = ub-lb+1

      if l > 0 then
        Dim iIndex as long: iIndex = lb + iCurrentIndex - 1
        if iIndex <= ub then
          Call CopyVariant(vNextItem,this.BaseArray(iIndex))
          vNextKey = iif(this.mode = FromIEnumVariant, iCurrentIndex, iIndex)
          
          NextItem = true
        else
          NextItem = false
        end if  
      else
        NextItem = false
      end if
    case FromCallable
      Dim vTmpItem as variant
      Call CopyVariant(vNextItem,this.Callback.Run(vNextItem, iCurrentIndex))
      vNextKey = iCurrentIndex
      NextItem = not isNull(vNextItem)
    case FromCallableVerbose
      Dim vVerboseReturn as Variant
      Call CopyVariant(vVerboseReturn,this.Callback.Run(vNextItem, iCurrentIndex))
      if isArray(vVerboseReturn) then
        Dim iArrLength as long: iArrLength = ubound(vVerboseReturn) - lbound(vVerboseReturn)+1
        if iArrLength > 0 and iArrLength <= 4 then
          vNextItem = null
          vNextKey = null
          NextItem = true
          
          Dim iArrLb as long: iArrLb = lbound(vVerboseReturn)
          if iArrLength >= 1 then NextItem      = vVerboseReturn(iArrLb+0)
          if iArrLength >= 2 then iCurrentIndex = vVerboseReturn(iArrLb+1)
          if iArrLength >= 3 then Call CopyVariant(vNextItem,vVerboseReturn(iArrLb+2)) 
          if iArrLength >= 4 then Call CopyVariant(vNextKey, vVerboseReturn(iArrLb+3))
        else
          Err.Raise 1, "stdEnumerator", "Incorrect return value length"
        end if
      else
        if isNull(vVerboseReturn) then
          NextItem = false
        else
          Err.Raise 1, "stdEnumerator", "Verbose callable should return an array of options: Continue?, nextIndex, nextValue, nextKey"
        end if
      end if
  end select
End Function

'Create an enumerator from an array
'@param varType - The type of array to create
'@returns - An enumerator containing the array
Private Function CreateEmptyVBArray(Optional ByVal varType as vbVarType = vbVariant) as Variant
    #if Mac then
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
    #else
        'TODO: LongPtr?
        select case varType
            case vbVariant:   CreateEmptyVBArray = createEmptyArrayVariant()
            case vbString:    CreateEmptyVBArray = createEmptyArrayString()
            case vbByte:      CreateEmptyVBArray = createEmptyArrayByte()
            case vbBoolean:   CreateEmptyVBArray = createEmptyArrayBoolean()
            case vbCurrency:  CreateEmptyVBArray = createEmptyArrayCurrency()
            case vbDate:      CreateEmptyVBArray = createEmptyArrayDate()
            case vbDouble:    CreateEmptyVBArray = createEmptyArrayDouble()
            case vbLong:      CreateEmptyVBArray = createEmptyArrayLong()
            case vbObject:    CreateEmptyVBArray = createEmptyArrayObject()
            #if Win64 then
            case vbLongLong:  CreateEmptyVBArray = createEmptyArrayLongLong()
            #end if
            case else
                Err.Raise 1, "stdEnumerator#CreateEmpty()", "Invalid vb vartype"
        end select
    #end if
End Function


'TODO: Implement raw API for working with IEnumVARIANT:
Private Function getNewEnum(ByRef o as object) as IEnumVARIANT

End Function
Private Function getNextVariant(ByRef o as IEnumVARIANT) as variant

End Function
Private Sub skipNextVariant(ByRef o as IEnumVARIANT)

End Sub
Private Function cloneEnumVariant(ByRef o as IEnumVARIANT) as IEnumVARIANT

End Function
Private Sub resetEnumVariant(ByRef o as IEnumVARIANT)
    
End Sub

