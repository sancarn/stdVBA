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
'    * mapWithIndex
'    ...
'In this way we can stack calls as follows:
'    Debug.print enumerableSet.map(s=>len(s)).filter(i=>i<=3).sum()

'Functions implemented on this class:
'CONSTRUCTORS
'    [X] CreateFromIEnumVariant
'    [ ] CreateFromICallable   'Only half implemented
'    [ ] CreateFromArray
'    [X] init  #PROTECTED
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
'    with_index    'this can't be easily done, so instead implement methods like `forEach(cb,withIndex?=false)
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

const VT_UNKNOWN = &Hd

Public Function CreateFromCallable(ByVal cb as stdICallable, optional byval iMaxLength as long = 1000000) as stdEnumerator
    set CreateFromCallable = new stdEnumerator
    Call CreateFromCallable.init(EnumeratorType.FromCallable,iMaxLength,cb)
End Function

Public Function CreateFromArray(ByVal v as variant, optional byval iMaxLength as long = 1000000) as stdEnumerator
    set CreateFromArray = new stdEnumerator
    Call CreateFromCallable.init(EnumeratorType.FromArray,iMaxLength,v)
End Function

Public Function CreateFromIEnumVariant(ByVal o as Object, optional byval iMaxLength as long = 1000000) as stdEnumerator
    set CreateFromIEnumVariant = new stdEnumerator
    Call CreateFromIEnumVariant.init(EnumeratorType.FromIEnumVariant,iMaxLength,o)
End Function

Public Sub Init(ByVal iEnumeratorType as long, ByVal iMaxLength as long, ParamArray v() As Variant)
    mode = iEnumeratorType
    pMaxLength = iMaxLength
    select case mode
        case EnumeratorType.FromIEnumVariant
            set pEnumObject = v(0)
            
            'Convert collection to array as IEnumVARIANT work around
            Dim i as long: i = 0
            Dim vArr()
            Redim vArr(1 to pEnumObject.count)
            Dim obj as object
            For each obj in pEnumObject
                i=i+1
                Call CopyVariant(vArr(i), obj)
            next

            pBaseArray = vArr
        case EnumeratorType.FromArray
            pBaseArray = v(0)
        case EnumeratorType.FromCallable
            set pCallback   = v(0)
    end select
End Sub

'Get this enumerator as a collection
'@returns {Collection} The enumerator's data as a collection
Public Function AsCollection() as Collection
    Dim i as long: i = 0
    Dim v as variant: v = null
    Dim retCol as Collection: set retCol = new Collection
    While NextItem(v,i)
        retCol.add v
    End Select

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
    Dim i as long: i = 0
    select case iType
        case vbVarType.vbObject
            while NextItem(v,i)
                set vRet(i) = v
            Wend
        case vbVarType.vbVariant
            while NextItem(v,i)
                if isObject(v) then
                    set vRet(i) = v
                else
                    vRet(i) = v
                end if
            Wend
        case else
            while NextItem(v,i)
                vRet(i) = v
            Wend
    end select
    
    AsArray = vRet
End Function


Public Function ForEach(Optional ByVal cb As stdICallable, Optional ByVal WithIndex as boolean = false) As stdEnumerator
    Dim v as variant: v = null
    Dim i as long: i = 0
    while NextItem(v,i)
        if withIndex then Call BindArg(v,i)
        Call cb.RunEx(v)
    Wend
    set ForEach = me
End Function

Public Function Map(Optional ByVal cb As stdICallable, Optional ByVal WithIndex as boolean = false) As stdEnumerator
    Dim oRet as Collection: Set oRet = new Collection
    Dim v as variant: v = null
    Dim i as long: i = 0
    while NextItem(v,i)
        if withIndex then Call BindArg(v,i)
        Call oRet.add(cb.RunEx(v))
    Wend

    set Map = stdEnumerator.CreateFromIEnumVariant(oRet)
End Function

Public Function Filter(ByVal cb as stdICallable, Optional ByVal WithIndex as boolean = false) as stdEnumerator
    Dim oRet as Collection: Set oRet = new Collection
    Dim v as variant: v = null
    Dim i as long: i = 0

    While NextItem(v,i)
        if withIndex then Call BindArg(v,i)
        if cb.RunEx(v) then Call oRet.add(v)
    Wend

    set Filter = stdEnumerator.CreateFromIEnumVariant(oRet)
End Function

Public Function Sort(ByVal cb as stdICallable) as stdEnumerator
    Dim arr() as SortStruct
    Dim iIndex as Long: iIndex = 0
    Dim iFirstItem as long: iFirstItem = 1
    Dim val as variant: val = null
    While NextItem(val,iIndex)
        'Increment index
        Redim Preserve arr(1 to iIndex)

        'Bind to SortStruct
        if iIndex > 1 Then
            'Initialise sorting struct
            Call CopyVariant(arr(iIndex).value, val)
            arr(iIndex).iIndex = iIndex
            arr(iIndex).sortValue = cb.RunEx(arr(iIndex).value)

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
            arr(1).sortValue = cb.RunEx(arr(1).value)
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

Public Function Unique() as stdEnumerator
    Dim oRet as Collection: Set oRet = new Collection
    Dim v as variant: v = null
    Dim i as long: i = 0

    While NextItem(v,i)
        Dim tv as variant: tv = null
        Dim ti as long: ti = i
        Dim matchFound as boolean: matchFound=false
        Do While NextItem(tv,ti)
            matchFound = areVariantsEqual(tv,v)
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
    
    While NextItem(v,i)
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
    
    While NextItem(v,i)
        sRet = sRet & sDelimiter & v
    Wend

    Join = mid(sRet,len(sDelimiter)+1)
End Function

Public Function indexOf(ByVal tv as variant) as long
    Dim oRet as Collection: Set oRet = new Collection
    Dim v as variant: v = null
    Dim i as long: i = 0
    
    While NextItem(v,i)
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
    
    While NextItem(v,i)
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
    checkAll = true
    While NextItem(v,i)
        if not cb.RunEx(v) then
            checkAll = false
            Exit Function
        end if
    Wend
End Function

Public Function checkAny(ByVal cb as stdICallable) as boolean
    Dim oRet as Collection: Set oRet = new Collection
    Dim v as variant: v = null
    Dim i as long: i = 0
    
    While NextItem(v,i)
        if cb.RunEx(v) then
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
    
    While NextItem(v,i)
        if cb.RunEx(v) then
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
    Dim iMatched as long: iMatched = 0
    While NextItem(v,i)
        if cb.RunEx(v) then
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
    end select
End Function

Public Function reduce(ByVal cb as stdICallable, Optional ByVal vInitialValue as variant = 0) as variant
    Dim v as variant: v = null
    Dim i as long: i = 0
    Dim v as variant, vRet as variant
    Call CopyVariant(vRet, vInitialValue)
    
    While NextItem(v,i)
        vRet = cb.RunEx(vRet,v)
    Wend
    Call CopyVariant(reduce,vRet)
End Function

Public Function countBy(ByVal cb as stdICallable) as long
    Dim iRet as Long: iRet = 0
    Dim v as variant: v = null
    Dim i as long: i = 0
    
    While NextItem(v,i)
        if cb.RunEx(v) then iRet = iRet + 1
    Wend
    countBy = iRet
End Function

Public Function groupBy(ByVal cb as stdICallable) as object
    Dim oRet as Object: set oRet = CreateObject("Scripting.Dictionary")
    Dim key as variant
    
    Dim v as variant: v = null
    Dim i as long: i = 0
    
    While NextItem(v,i)
        Call CopyVariant(key,cb.RunEx(v))
        if not oRet.exists(key) then set oRet(key) = new Collection
        Call oRet(key).add(v)
    Wend
            
    Dim keys as variant: keys = oRet.keys()
    Dim i as long
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
    
    While NextItem(v,i)
        'Get value to test
        Dim vtValue as variant
        if cb is nothing then
            Call CopyVariant(vtValue,v)
        else
            Call CopyVariant(vtValue,cb.RunEx(v))
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
    
    While NextItem(v,i)
        'Get value to test
        Dim vtValue as variant
        if cb is nothing then
            Call CopyVariant(vtValue,v)
        else
            Call CopyVariant(vtValue,cb.RunEx(v))
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
    
    While NextItem(v,i)
        'Get value to test
        Dim vtValue as variant
        if cb is nothing then
            Call CopyVariant(vtValue,v)
        else
            Call CopyVariant(vtValue,cb.RunEx(v))
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
        
        While NextItem(v,ti)
            Call cb.runEx(v)
        Wend
    next
    set Cycle = Me
End Function

Public Function FindFirst(ByVal cb as stdICallable) as variant
    Dim oRet as Collection: Set oRet = new Collection
    Dim v as variant: v = null
    Dim i as long: i = 0
    
    While NextItem(v,i)
        if cb.runEx(v) then
            Call CopyVariant(FindFirst,v)
            Exit Function
        end if
    Wend

    FindFirst = Null
End Function

Public Property Item(ByVal i as Long) as variant
    Dim oRet as Collection: Set oRet = new Collection
    Dim v as variant: v = null
    Dim ti as long: ti = 0
    
    While NextItem(v,ti)
        if ti = i then
            Call CopyVariant(Item,v)
            Exit Function
        end if
    Wend
    
    Err.Raise 1, "stdEnumerator", "Error: No such item " & i & " in stdEnumerator."
End Property 

Public Property Length() as Long
    Dim oRet as Collection: Set oRet = new Collection
    Dim v as variant: v = null
    Dim i as long: i = 0
    Dim iMaxLength as Long: iMaxLength = pMaxLength
    if mode = FromArray then
        Length = UBound(pBaseArray) - LBound(pBaseArray) + 1
    else
        While NextItem(v,i)
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
            for each t in flattenChildrenRecurse(sv)
                oRet.Add t
            next
        next 
    elseif implementsIEnumVariant(parent) then
        'Loop through children
        for each child in parent
            'Recursive add to oRet
            for each t in flattenChildrenRecurse(sv)
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
'@param {ByRef Variant} Previous/Next Item
'@param {ByRef Long}    Previous/Next Index
'@returns {Boolean}     Whether a new item was succesfully obtained or not.
Private Function NextItem(ByRef vNextItem as Variant, ByRef iCurrentIndex as long) as Boolean
  iCurrentIndex = iCurrentIndex+1
  select case mode
    case FromIEnumVariant, FromArray
      Dim lb as long: lb = LBound(pBaseArray,1)
      Dim ub as long: ub = Ubound(pBaseArray,1)
      Dim l as long: l = ub-lb+1
      if l > 0 then
        Dim iIndex as long: iIndex = lb + iCurrentIndex - 1
        if iIndex <= ub then
          Call CopyVariant(vNextItem,Array(pBaseArray(iIndex)))
          NextItem = true
        else
          NextItem = false
        end if  
      else
        NextItem = false
      end if
    case FromCallable
      Call CopyVariant(vNextItem,arrayFromCollection(pCallback.Bind(new Collection).Run(vNextItem, i)))
      if isNull(vNextItem(1)) then
        NextItem = false
      else
        NextItem = true
      end if
  end select
End Function

Private Sub BindArg(ByRef vArgs as Variant, ByVal vArg as variant)
    Redim Preserve vArgs(lbound(vArgs) to ubound(vArgs)+1)
    Call CopyVariant(vArgs(ubound(vArgs)), vArg)
End Sub

Private Function arrayFromCollection(ByVal col as Collection) as Variant
    Dim i as long: i = 0
    Dim vRet() as variant
    Redim vRet(1 to col.count)
    For each v in col
        i = i+1
        Call CopyVariant(vRet(i),v)
    next
    arrayFromCollection = vRet
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