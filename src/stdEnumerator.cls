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

Private mode As EnumeratorType
Private pEnumObject as Object
Private pBaseArray as Variant
Private pCallback as stdICallable
Private pMaxLength as Long

Private const VT_UNKNOWN = &Hd

Public Function CreateFromCallable(ByVal cb as stdICallable, optional byval iMaxLength as long = 1000000) as stdEnumerator
    set CreateFromCallable = new stdEnumerator
    Call CreateFromCallable.protInit(EnumeratorType.FromCallable,iMaxLength,cb)
End Function

'@param {stdICallable<()=>Array<Boolean,Long,Variant,Variant>>}
Public Function CreateFromCallableVerbose(ByVal cb as stdICallable, optional byval iMaxLength as long = 1000000) as stdEnumerator
    set CreateFromCallableVerbose = new stdEnumerator
    Call CreateFromCallableVerbose.protInit(EnumeratorType.FromCallableVerbose,iMaxLength,cb)
End Function


Public Function CreateFromArray(ByVal v as variant, optional byval iMaxLength as long = 1000000) as stdEnumerator
    set CreateFromArray = new stdEnumerator
    Call CreateFromArray.protInit(EnumeratorType.FromArray,iMaxLength,v)
End Function

Public Function CreateFromIEnumVariant(ByVal o as Object, optional byval iMaxLength as long = 1000000) as stdEnumerator
    set CreateFromIEnumVariant = new stdEnumerator
    Call CreateFromIEnumVariant.protInit(EnumeratorType.FromIEnumVariant,iMaxLength,o)
End Function

Public Function CreateEmpty(optional byval iMaxLength as long = 1000000) as stdEnumerator
    set CreateEmpty = new stdEnumerator
    Call CreateEmpty.protInit(EnumeratorType.FromIEnumVariant,iMaxLength,new Collection)
End Function

'TODO: CreateFromDictionary
'Public Function CreateFromDictionary(ByVal o as object, optional byval iMaxLength as long = 1000000) as stdEnumerator
'    set CreateFromDictionary = new stdEnumerator
'    Call CreateFromDictionary.protInit(EnumeratorType.FromDictionary,iMaxLength,o)
'End Function

Public Sub protInit(ByVal iEnumeratorType as long, ByVal iMaxLength as long, ParamArray v() As Variant)
    mode = iEnumeratorType
    pMaxLength = iMaxLength
    select case mode
        case EnumeratorType.FromIEnumVariant
            'HACK: This is a real work around for the fact that we can't access IEnumVARIANT methods (getNewEnum,getNextVariant,skipNextVariant,cloneEnumVariant,resetEnumVariant)
            'directly. Realistically this may sometimes not even work. A more long term solution would be to call getNewEnum() and then getNextVariant() until we reach the end of the
            'object. This may be implementable after stdCOM is fully implemented.

            set pEnumObject = v(0)
            
            'Calculate total length
            Dim iLength As Long: iLength = 0
            Dim obj As Variant
            For Each obj In pEnumObject
                iLength = iLength + 1
            next
            
            'Define array
            if iLength > 0 then
                Dim vArr()
                ReDim Preserve vArr(1 To iLength)
                
                'Copy IEnumVARIANT into array
                Dim i As Long: i = 0
                For Each obj In pEnumObject
                    i = i + 1
                    Call CopyVariant(vArr(i), obj)
                Next

                pBaseArray = vArr
            else
                pBaseArray = Array()
            end if
        case EnumeratorType.FromArray
            pBaseArray = v(0)
        case EnumeratorType.FromCallable
            set pCallback   = v(0)
        case EnumeratorType.FromCallableVerbose
            set pCallback   = v(0)
    end select
End Sub

'Get this enumerator as a collection
'@returns {Collection} The enumerator's data as a collection
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
'@param {ByVal vbVarType?} The vartype of the array generated 
'@returns {Collection} The enumerator's data as an array
Public Function AsArray(Optional ByVal iType as vbVarType = vbVarType.vbVariant) as Variant
    Dim arrLen as long: arrLen = Length()
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

'Returns the stored enumerator as a dictionary
'@returns {Dictionary} Dictionary containing enumerator data
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
'@param {stdICallable} callback to call
'@returns {stdEnumerator} This is the enumerator itself
Public Function ForEach(ByVal cb As stdICallable) As stdEnumerator
  Dim v as variant: v = null
  Dim i as long: i = 0
  Dim kv as variant: kv = null 
  while NextItem(kv,v,i)
    Call cb.Run(v, kv)
  Wend
  set ForEach = me
End Function

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

Public Function Sort(Optional ByVal cb as stdICallable = nothing) as stdEnumerator
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
End Function

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

Public Function Reverse() as stdEnumerator
  Dim oRet as collection: set oRet = new Collection
  Dim v: v = AsArray()
  Dim i as long:
  For i = ubound(v) to lbound(v) step -1
    oRet.add v(i)
  next

  set Reverse = stdEnumerator.CreateFromIEnumVariant(oRet)
End Function

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

Public Function includes(ByVal tv as variant) as boolean
    includes = indexOf(tv) > 0
End Function

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

Public Function reduce(ByVal cb as stdICallable, Optional ByVal vInitialValue as variant = 0) as variant
    Dim v as variant: v = null
    Dim i as long: i = 0
    Dim kv as variant: kv = null
    Dim vRet as variant
    Call CopyVariant(vRet, vInitialValue)
    
    While NextItem(kv, v,i)
        vRet = cb.Run(vRet,v,kv)
    Wend
    Call CopyVariant(reduce,vRet)
End Function

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

Public Function Flatten() as stdEnumerator
    set Flatten = stdEnumerator.CreateFromIEnumVariant(flattenChildrenRecurse(me))
End Function

'Run the callback on an all of the items i times.
'@param {ByVal Long} - Times to execute the cb on the collection
'@param {ByVal stdICallable} - Callback to execute on each item
'@returns {stdEnumerator} - This object.
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

Public Function FindFirst(ByVal cb as stdICallable) as variant
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

    FindFirst = Null
End Function

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

Public Property Get Length() as Long
    Dim oRet as Collection: Set oRet = new Collection
    Dim v as variant: v = null
    Dim i as long: i = 0
    Dim iMaxLength as Long: iMaxLength = pMaxLength
    if mode = FromArray then
        Length = UBound(pBaseArray) - LBound(pBaseArray) + 1
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




'===================================================================================================
'Compares 2 variables and returns true if they are equal
'@param {ByVal Variant} Variable 1
'@param {ByVal Variant} Variable 2
'@returns {Boolean} True if the variables are equal, false otherwise.
Private Function areVariantsEqual(ByVal v1 as variant, ByVal v2 as variant) as boolean
    if isObject(v1) and isObject(v2) then
        areVariantsEqual = v1 is v2
    elseif (not isobject(v1)) and (not isObject(v2)) then
        areVariantsEqual = v1 = v2
    end if
End Function

'Copy the value from vSrc into vDest
'@param {ByRef Variant} Destination value
'@param {ByVal Variant} Source value
'@returns {VOID}
Private Sub CopyVariant(ByRef vDest as variant, ByVal vSrc as Variant)
    if isObject(vSrc)  Or  VarType(vSrc) = VT_UNKNOWN Then
        Set vDest = vSrc
    else
        Let vDest = vSrc
    end if
End Sub

'Check whether an object implements IEnumVARIANT
'@param {ByVal Variant} Value to check
'@returns {Boolean} True if it implements IEnumVARIANT, false otherwise
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
'@param {ByVal Variant} Object/Array/Whatever to flatten
'@returns {Collection} Flattened collection of items
Private Function flattenChildrenRecurse(ByVal parent as variant) as Collection
    Dim oRet as Collection: set oRet = new Collection
    Dim child as variant
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
'@param {ByRef Variant} Previous/Next Key
'@param {ByRef Variant} Previous/Next Item
'@param {ByRef Long}    Previous/Next Index
'@returns {Boolean}     Whether a new item was succesfully obtained or not.
Private Function NextItem(ByRef vNextKey as variant, ByRef vNextItem as Variant, ByRef iCurrentIndex as long) as Boolean
  iCurrentIndex = iCurrentIndex+1
  select case mode
    case FromIEnumVariant, FromArray
      Dim lb as long: lb = LBound(pBaseArray,1)
      Dim ub as long: ub = Ubound(pBaseArray,1)
      Dim l as long: l = ub-lb+1

      if l > 0 then
        Dim iIndex as long: iIndex = lb + iCurrentIndex - 1
        if iIndex <= ub then
          Call CopyVariant(vNextItem,pBaseArray(iIndex))
          vNextKey = iif(mode = FromIEnumVariant, iCurrentIndex, iIndex)
          
          NextItem = true
        else
          NextItem = false
        end if  
      else
        NextItem = false
      end if
    case FromCallable
      Dim vTmpItem as variant
      Call CopyVariant(vNextItem,pCallback.Run(vNextItem, iCurrentIndex))
      vNextKey = iCurrentIndex
      NextItem = not isNull(vNextItem)
    case FromCallableVerbose
      Dim vVerboseReturn as Variant
      Call CopyVariant(vVerboseReturn,pCallback.Run(vNextItem, iCurrentIndex))
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

''Feel this will be impossible currently, perhaps possible with DispCallFunc to call raw Next, Skip, Clone, Reset methods:
''    Dim x as IEnumVARIANT
''    set x = pEnumObject.[_NewEnum]
''    DispCallFunc(x,1,...) 'Call QueryInterface
''    DispCallFunc(x,2,...) 'Call AddRef
''    DispCallFunc(x,3,...) 'Call Release
''    DispCallFunc(x,4,...) 'Call Next
''    DispCallFunc(x,5,...) 'Call Skip
''    DispCallFunc(x,6,...) 'Call Reset
''    DispCallFunc(x,7,...) 'Call Clone
'
'
'
'   Public Function FeedItem(v As Variant)
'       Call oRetArray.push(v)
'   End Function
'   
'   'Incompatible with oEnumVariant
'   Public Function NextItem() As Variant
'       mode = iMode
'       If mode = FromForEach Then
'           
'       ElseIf mode = FromCallback Then
'           
'       ElseIf mode = FromIEnumVariant Then
'           Set NextItem = oEnumVariant.Next()
'       End If
'   End Function
'   Public Function NextItems() As stdArray
'       mode = iMode
'       Set NextItems = stdArray.Create()
'       If mode = FromForEach Then
'           
'       ElseIf mode = FromCallback Then
'           
'       ElseIf mode = FromIEnumVariant Then
'           Dim x As IEnumVARIANT
'           Set x = oEnumVariant.Clone
'           
'           Z = x.Next
'           While Z
'               Call NextItems.push(Z)
'               Z = x.Next
'           Wend
'       End If
'   End Function
'   Public Function PeekItem() As Variant
'       If mode = FromForEach Then
'           
'       ElseIf mode = FromCallback Then
'           
'       ElseIf mode = FromIEnumVariant Then
'           Dim x As IEnumVARIANT
'           Set x = oEnumVariant.Clone
'           Set PeekItem = x.Next
'       End If
'   End Function
'   Public Function PeekItems() As stdArray
'       
'   End Function
'   Public Function Rewind() As stdEnumerator
'       
'   End Function
'   Public Function Size() As Long
'   
'   End Function