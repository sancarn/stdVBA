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
'    [X] CreateFromICallable
'    [X] init  #PROTECTED
'
'INSTANCE METHODS
'Many methods were inspired by those in Ruby's Enumerable: https://ruby-doc.org/core-2.7.2/Enumerable.html
'    [X] Sort() as stdArray
'    [X] Reverse() as stdArray
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
Private pCallback as stdICallable

'TODO: Implement ICallable methods in some way
Public Function CreateFromCallable(ByVal cb as stdICallable) as stdEnumerator
    set CreateFromCallable = new stdEnumerator
    Call CreateFromCallable.init(EnumeratorType.FromCallable,cb)
End Function

Public Function CreateFromIEnumVariant(ByVal o as IUnknown) as stdEnumerator
    set CreateFromIEnumVariant = new stdEnumerator
    Call CreateFromIEnumVariant.init(EnumeratorType.FromIEnumVariant,o)
End Function

Public Sub Init(ByVal iEnumeratorType as long, ParamArray v() As Variant)
    mode = iEnumeratorType
    select case mode
        case EnumeratorType.FromIEnumVariant
            set pEnumObject = v(0)
        case EnumeratorType.FromCallable
            set pCallback   = v(0)
    end select
End Sub


Public Function ForEach(Optional ByVal cb As stdICallable, Optional ByVal WithIndex as boolean = false) As stdEnumerator
    select case mode
        case EnumeratorType.FromCallable
        
        case EnumeratorType.FromIEnumVariant
            Dim v as variant, i as long: i=0
            For Each v In pEnumObject
                i=i+1
                if withIndex then
                    Call cb.Run(i,v)
                else
                    Call cb.Run(v)
                end if
            Next
    End Select

    set ForEach = me
End Function

Public Function Map(Optional ByVal cb As stdICallable, Optional ByVal WithIndex as boolean = false) As stdEnumerator
    Dim oRet as Collection
    Set oRet = new Collection

    select case mode
        case EnumeratorType.FromCallable
        
        case EnumeratorType.FromIEnumVariant
            Dim v as variant, i as long: i=0
            For Each v In pEnumObject
                i=i+1
                if withIndex then
                    Call oRet.add(cb.Run(i,v))
                else
                    Call oRet.add(cb.Run(v))
                end if
            Next
    End Select

    set Map = stdEnumerator.CreateFromIEnumVariant(oRet)
End Function

Public Function Filter(ByVal cb as stdICallable, Optional ByVal WithIndex as boolean = false) as stdEnumerator
    Dim oRet as Collection
    Set oRet = new Collection

    select case mode
        case EnumeratorType.FromCallable
        
        case EnumeratorType.FromIEnumVariant
            Dim v as variant, i as long: i = 0
            For Each v In pEnumObject
                if cb.Run(v) then
                    if withIndex then
                        i=i+1
                        if cb.RunEx(i,v) then Call oRet.add(v)
                    else
                        if cb.RunEx(v) then Call oRet.add(v)
                    end if
                end if
            Next
    End Select

    set Filter = stdEnumerator.CreateFromIEnumVariant(oRet)
End Function

Public Function Sort(cb as callback, ctor as object) as stdEnumerator
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
    Dim oRet as Collection
    Set oRet = new Collection

    select case mode
        case EnumeratorType.FromCallable
            'TODO:
        case EnumeratorType.FromIEnumVariant
            Dim v as variant
            For Each v In pEnumObject
                Dim matchFound as boolean: matchFound=false
                Dim tv as variant
                for each tv in oRet
                    matchFound = areVariantsEqual(tv,v)
                next

                'If a match isn't found return v
                if not matchFound then Call oRet.add(v)
            Next
    End select

    set Unique = stdEnumerator.CreateFromIEnumVariant(oRet)
End Function

Public Function Reverse() as stdEnumerator
    'TODO: Might be hard to do optimally
    Dim oRet as Collection
    Set oRet = new Collection

    select case mode
        case EnumeratorType.FromCallable
        
        case EnumeratorType.FromIEnumVariant
            Dim x as collection
            set x = new collection
            For Each v In pEnumObject
                Call x.add(v)
            Next
            Dim i as long
            For i = x.count to 1 step -1
                Call oRet.add(x.item(i))
            Next
    End select

    set Reverse = stdEnumerator.CreateFromIEnumVariant(oRet)
End Function

Public Function Concat(ByVal obj as IUnknown) as stdEnumerator
    Dim oRet as Collection
    Set oRet = new Collection

    select case mode
        case EnumeratorType.FromCallable
        
        case EnumeratorType.FromIEnumVariant
            for each v in pEnumObject
                Call oRet.add(v)
            next
            for each v in obj
                Call oRet.add(v)
            next
    end select

    set Concat = stdEnumerator.CreateFromIEnumVariant(oRet)
End Function

Public Function Join(Optional ByVal sDelimiter as string = ",") as string
    Dim sRet as string
    sRet = ""

    select case mode
        case EnumeratorType.FromCallable
        
        case EnumeratorType.FromIEnumVariant
            for each v in pEnumObject
                sRet = sRet & "," & v
            next
    end select

    set Join = mid(sRet,2)
End Function

Public Function indexOf(ByVal tv as variant) as long
    select case mode
        case EnumeratorType.FromCallable
        
        case EnumeratorType.FromIEnumVariant
            Dim iRet as long: iRet = 0
            for each v in pEnumObject
                iRet = iRet + 1
                if areVariantsEqual(tv,v) then
                    indexOf = iRet
                    Exit Function
                end if
            next
    end select

    indexOf = 0
End Function

Public Function lastIndexOf(ByVal tv as variant) as long
    indexOf = 0
    select case mode
        case EnumeratorType.FromCallable
        
        case EnumeratorType.FromIEnumVariant
            Dim iRet as long: iRet = 0
            for each v in pEnumObject
                iRet = iRet + 1
                if areVariantsEqual(tv,v) then
                    indexOf = iRet
                end if
            next
    end select
End Function

Public Function includes(ByVal tv as variant) as boolean
    includes = indexOf(tv) > 0
End Function

Public Function checkAll(ByVal tv as variant) as boolean
    select case mode
        case EnumeratorType.FromCallable
        
        case EnumeratorType.FromIEnumVariant
            Dim v as variant
            for each v in pEnumObject
                if not areVariantsEqual(tv,v) then
                    checkAll = false
                    Exit Function
                end if
            next
            checkAll = true
    end select
End Function

Public Function checkAny(ByVal tv as variant) as boolean
    select case mode
        case EnumeratorType.FromCallable
        
        case EnumeratorType.FromIEnumVariant
            Dim v as variant
            for each v in pEnumObject
                if areVariantsEqual(tv,v) then
                    checkAll = true
                    Exit Function
                end if
            next
            checkAll = false
    end select
End Function

Public Function checkNone(ByVal tv as variant) as boolean
    select case mode
        case EnumeratorType.FromCallable
        
        case EnumeratorType.FromIEnumVariant
            Dim v as variant
            for each v in pEnumObject
                if areVariantsEqual(tv,v) then
                    checkNone = false
                    Exit Function
                end if
            next
            checkNone = true
    end select
End Function

Public Function checkOnlyOne(ByVal tv as variant) as boolean
    select case mode
        case EnumeratorType.FromCallable
        
        case EnumeratorType.FromIEnumVariant
            Dim v as variant, iMatched as long: iMatched = 0
            for each v in pEnumObject
                if areVariantsEqual(tv,v) then
                    iMatched = iMatched + 1
                    if iMatched > 1 then
                        checkOnlyOne = false
                        Exit Function
                    end if
                end if
            next
            if iMatched = 0 then
                checkOnlyOne = false
                Exit Function
            End if
            checkOnlyOne = true
    end select
End Function

Public Function reduce(ByVal cb as stdICallable) as variant
    select case mode
        case EnumeratorType.FromCallable
        
        case EnumeratorType.FromIEnumVariant
            Dim v as variant, oRet as variant
            for each v in pEnumObject
                oRet = cb.Run(oRet,v)
            next
            Call CopyVariant(reduce,oRet)
    end select
End Function

Public Function countBy(ByVal cb as stdICallable) as long
    Dim iRet as Long: iRet = 0
    select case mode
        case EnumeratorType.FromCallable
        
        case EnumeratorType.FromIEnumVariant
            Dim v as variant
            for each v in pEnumObject
                if cb.Run(v) then iRet = iRet + 1
            next
            countBy = iRet
    end select
End Function

Public Function groupBy(ByVal cb as stdICallable) as object
    Dim oRet as Object: set oRet = CreateObject("Scripting.Dictionary")
    Dim key as variant
    select case mode
        case EnumeratorType.FromCallable
        
        case EnumeratorType.FromIEnumVariant
            Dim v as variant
            for each v in pEnumObject
                CopyVariant(key,cb.Run(v))
                if not oRet.exists(key) then set oRet(key) = new Collection
                oRet(key).add(v)
            next
            
            Dim keys as variant: keys = oRet.keys()
            Dim i as long
            For i = 0 to ubound(keys)
                set oRet(keys(i)) = stdEnumerator.CreateFromIEnumVariant(oRet(keys(i)))
            next
            set groupBy = oRet
    end select
End Function

Public Function max(Optional ByVal cb as stdICallable = nothing) as long
    Dim vRet as variant
    Dim vMaxValue as variant
    select case mode
        case EnumeratorType.FromCallable
        
        case EnumeratorType.FromIEnumVariant
            Dim v as variant
            for each v in pEnumObject
                'Get value to test
                Dim vtValue as variant
                if cb is nothing then
                    CopyVariant(vtValue,v)
                else
                    CopyVariant(vtValue,cb.Run(v))
                end if

                'Compare values and return
                if isEmpty(vRet) then
                    CopyVariant(vRet,v)
                    CopyVariant(vMaxValue, vtValue)
                elseif vMaxValue < vtValue
                    CopyVariant(vRet,v)
                    CopyVariant(vMaxValue, vtValue)
                end if
            next
            
            CopyVariant(max,vRet)
    end select
End Function

Public Function min(Optional ByVal cb as stdICallable = nothing) as long
    Dim vRet as variant
    Dim vMaxValue as variant
    select case mode
        case EnumeratorType.FromCallable
        
        case EnumeratorType.FromIEnumVariant
            Dim v as variant
            for each v in pEnumObject
                'Get value to test
                Dim vtValue as variant
                if cb is nothing then
                    CopyVariant(vtValue,v)
                else
                    CopyVariant(vtValue,cb.Run(v))
                end if

                'Compare values and return
                if isEmpty(vRet) then
                    CopyVariant(vRet,v)
                    CopyVariant(vMaxValue, vtValue)
                elseif vMaxValue > vtValue
                    CopyVariant(vRet,v)
                    CopyVariant(vMaxValue, vtValue)
                end if
            next
            
            CopyVariant(min,vRet)
    end select
End Function

Public Function sum(Optional ByVal cb as stdICallable = nothing) as long
    Dim vRet as variant
    select case mode
        case EnumeratorType.FromCallable
        
        case EnumeratorType.FromIEnumVariant
            Dim v as variant
            for each v in pEnumObject
                'Get value to test
                Dim vtValue as variant
                if cb is nothing then
                    CopyVariant(vtValue,v)
                else
                    CopyVariant(vtValue,cb.Run(v))
                end if
                vRet = vRet + vtValue
            next
            
            CopyVariant(sum,vRet)
    end select
End Function

Public Function Flatten() as stdEnumerator
    Dim oRet as Collection
    Set oRet = new Collection

    select case mode
        case EnumeratorType.FromCallable
        
        case EnumeratorType.FromIEnumVariant
            for each v in pEnumObject
                Dim sv as variant
                if implementsIEnumVariant(v) then
                    for each sv in v
                        call oRet.add(sv)
                    next
                else
                    call oRet.add(v)
                end if
            next
    end select

    set Flatten = stdEnumerator.CreateFromIEnumVariant(oRet)
End Function

Public Function Cycle(ByVal iTimes as long, ByVal cb as stdICallable)
    select case mode
        case EnumeratorType.FromCallable
        
        case EnumeratorType.FromIEnumVariant
            Dim i as long
            for i = 1 to iTimes
                Dim v as variant
                For each v in pEnumObject
                    Call cb.run(v)
                next
            next
    end select
    set Cycle = Me
End Function

Public Function FindFirst(ByVal cb as stdICallable) as variant
    select case mode
        case EnumeratorType.FromCallable
        
        case EnumeratorType.FromIEnumVariant
            Dim v as variant
            For each v in pEnumObject
                if cb.run(v) then
                    CopyVariant(FindFirst,v)
                    Exit Function
                end if
            next
    end select
    set Cycle = Me
End Function








'===================================================================================================

Private Function areVariantsEqual(ByVal v1 as variant, ByVal v2 as variant) as boolean
    if isObject(tv) and isObject(v) then
        areVariantsEqual = tv is v
    elseif (not isobject(tv)) and (not isObject(v)) then
        areVariantsEqual = tv = v
    end if
End Function
Private Sub CopyVariant(ByRef vDest as variant, ByVal vSrc as Variant)
    if isObject(vSrc) then
        set vDest = vSrc
    else
        vDest = vSrc
    end if
End Sub
Private Function implementsIEnumVariant(ByVal o as IUnknown) as boolean
    On Error GoTo ErrorOccurred
        Dim x as IEnumVARIANT
        set x = o
    On Error GoTo 0
    
    implementsIEnumVariant = true
    Exit Function

ErrorOccurred:
    implementsIEnumVariant=false
End Function


'TODO: Implement raw API for working with IEnumVARIANT:
Private Function getNextVariant(o as IEnumVARIANT) as variant

End Function
Private Sub skipNextVariant(o as IEnumVARIANT)

End Sub
Private Function cloneEnumVariant(o as IEnumVARIANT) as IEnumVARIANT

End Function
Private Sub resetEnumVariant(o as IEnumVARIANT)

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









