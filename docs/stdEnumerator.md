# `stdEnumerator`

## Spec

### Constructors

#### `CreateFromIEnumVariant(ByVal oEnumerable as Object, optional byval iMaxLength as long = 1000000) as stdEnumerator`

This method is used to create a `stdEnumerator` from an existing object which implements `IEnumVARIANT`. The `IEnumVARIANT` interface is a hidden interface used by VBA to allow usage of `for each item in objectImplementingIEnumVariant` syntax. Common usage of this method will be to create enumerators from `Collections`, `Application.Workbooks`, `MyWorkbook.Sheets` or typically any other enumerable collection-style object.

Parameter `oEnumerable` is the enumerable object to build the `stdEnumerator` from.

Parameter `iMaxLength` will set a hard limit to the number of elements in the array to prevent freezing from infinitely large enumerator generators.

```vb
Dim col as collection: set col = new collection
col.add 1
col.add 2
Debug.Print stdEnumerator.CreateFromIEnumVariant(col).join() '=> 1,2

'Create stdEnumerator on Workbooks collection-style object
stdEnumerator.CreateFromIEnumVariant(Application.Workbooks).forEach(stdLambda.Create("$1.save"))
```

#### `CreateFromArray(ByVal vArray as variant, optional byval iMaxLength as long = 1000000) as stdEnumerator`

This method is used to create a `stdEnumerator` from an existing 1 dimensional array.

Parameter `vArray` is the array to build the `stdEnumerator` from.

Parameter `iMaxLength` will set a hard limit to the number of elements in the array to prevent freezing from infinitely large enumerator generators.

```vb
'Join an array
Debug.Print stdEnumerator.CreateFromArray(Array(1,2,3)).join() '=> 1,2,3

'Deduplicate array:
Dim arr: arr = Array(1,2,1,3,4,2,5,5,9,1)
arr = stdEnumerator.CreateFromArray(arr).Unique().AsArray(vbLong)
```

#### `CreateFromCallable(ByVal cb as stdICallable<(item: Variant, key: Variant)=>Variant|Null>, optional byval iMaxLength as long = 1000000) as stdEnumerator`

This is an advanced method which allows you to create custom `stdEnumerator` instances with custom behaviour.

On initial call, `vPreviousItem` will be `null` and `iCurrentIndex` will be `1`.

> Note: If `stdLambda` is used as callback, you don't have to bother about the above signature but `$1` will consume the previous item, and `$2` will consume the current index.

To mark the end of the enumerator, return `null`.

Parameter `iMaxLength` will set a hard limit to the number of elements in the array to prevent freezing from infinitely large enumerator generators.

```vb
'1,2,3,4,5,6,7,8,9
stdEnumerator.CreateFromCallable(stdLambda.Create("if $2 <= 9 then $2 else null"))

'"a","aa","aaa","aaaa","aaaaa"
stdEnumerator.CreateFromCallable(stdLambda.Create("if $2 = 1 then ""a"" else if len($1) <= 5 then $1 & ""a"" else null"))

'"a","aa","aaa","aaaa","aaaaa"
stdEnumerator.CreateFromCallable(stdLambda.Create("$1.getNext()").bind(customEnumObj))
```

> Note: Currently `stdEnumerator` DOES NOT implement `IEnumVARIANT` and for this reason `for each item in myEnumerator` syntax **will not work**.

#### `CreateFromCallableVerbose(ByVal cb as stdICallable<(item: Variant, key: Variant)=>[continue: boolean, nextIndex: Long, nextItem: Variant, nextKey: Variant]>, optional byval iMaxLength as long = 1000000) as stdEnumerator`

This is an advanced method which allows you to create custom `stdEnumerator` instances with custom behaviour. This is a significantly more verbose function but offers significantly greater flexibility.

The return value from myCallback is to return an array of up to 4 elements:

```vb
Array( _
  continue, _    'Whether a next element exists or not
  iNextIndex, _  'The index of the next element. This can be incremented beyond that which was passed into it giving ability to skip elements.
  vNextItem, _   'The next item to return
  vNextKey _     'The next key to return
)
```

Parameter `iMaxLength` will set a hard limit to the number of elements in the array to prevent freezing from infinitely large enumerator generators.

##### Examples

Here are a few examples:

```vb
'Return array from callable representing:
'1,2,3,4,5,6,7,8,9
set enumerator = stdEnumerator.CreateFromCallable(stdLambda.Create("Array($2 <= 9, $2, $2, $2)"))

'"a","aa","aaa","aaaa","aaaaa"
set enumerator = stdEnumerator.CreateFromCallable(stdLambda.Create("Array($2 <= 5, $2, if $2 = 1 then ""a"" else $1 & ""a"", $2)"))
```

> Note: Currently `stdEnumerator` DOES NOT implement `IEnumVARIANT` and for this reason `for each item in myEnumerator` syntax **will not work**.

### GUIDANCE

#### Procedural Enumeration

In later patches enumeration methods were added to the `stdEnumerator` library. These allow for enumeration methods similar to `for-each` syntax. This works as follows:

```vb
Dim eValues as stdEnumerator: set eValues = stdEnumerator.CreateFromArray(Array(1,2,3))

Debug.print "Beg"

Dim iValue as Long
While eValues.enumNext(iValue)
  Debug.Print iValue
Wend

Debug.print "End"
```

The result of which will be printed in the console:

```
Beg
1
2
3
End
```

With these methods you can easily and efficiently iterate over the `stdEnumerator` set and perform whatever manipulation/extraction. This is advantageous in situations where loading the entire dataset into would be slow or impossible. It is suggested that before looping over the collection you use `stdEnumerator#enumRefresh`. This will ensure the enumerator is set at the beginning of the dataset.

```vb
Dim iValue as Long
Call eValues.enumRefresh()
While eValues.enumNext(iValue)
  Debug.Print iValue
Wend
```

### INSTANCE PROPERTIES

#### `Item(ByVal index as Variant, optional byval byIndex as boolean = false) as variant`

Retrieve an item from the enumerator object. If `byIndex` is `false` keys are prioritised over keys.

> Note: Keys of collections are not currently preserved. This is something which will be [added at a later date](https://stackoverflow.com/questions/5702362/vba-collection-list-of-keys/5702524).

```vb
Debug.Print enumerator.item(1)
```

#### `Length() as Long`

Obtain the length of the enumerator contents.

```vb
Debug.Print enumerator.length
```

### INSTANCE METHODS

#### `enumNext(ByRef vOut as Variant) as Boolean`

Gets the next element from the enumerator relative to the cursor. If this function returns `true` a new value was retrieved. Else if it returns `false` the enumerator has reached the end of the collection.

The next element is set to the inputted parameter `vOut`.

#### `enumRefresh()`

Set the enumerator cursor to the beginning of the collection.

#### `asCollection()`

Get this enumerator as a collection.

```vb
For each o in enumerator.AsCollection()
  '...
Next
```

#### `asArray(iType as vbVarType)`

Get this enumerator as an array of type `iType`.

```vb
set enumerator = stdEnumerator.CreateFromArray(Array(1,2,3,4,5))
Debug.Print join(enumerator.AsArray(vbString),",") = "1,2,3,4,5"
```

#### `asDictionary()`

Get this enumerator as a dictionary.

```vb
set enumerator = stdEnumerator.CreateFromArray(Array(1,5,6,10,25))
Debug.Print enumerator.asDictionary().exists(10)
```

#### `Sort(Optional ByVal cb as stdICallable<(item: Variant, key: Variant)=>Variant> = nothing) as stdEnumerator`

Sort the enumerator contents by the value retrieved by either the values in the enumerator or the values returned by `cb`.

#### `Reverse() as stdEnumerator`

Reverse the elements in the enumerator.

#### `ForEach(ByVal cb As stdICallable<(item: Variant, key: Variant)=>Variant>) As stdEnumerator`

Call callback `cb` on each item in the enumerator.

#### `Map(ByVal cb As stdICallable<(item: Variant, key: Variant)=>Variant>) As stdEnumerator`

Call callback `cb` on each item in the enumerator. Creates a new enumerator from each value returned from `cb`.

#### `Unique(optional byval cb as stdICallable<(item: Variant, key: Variant)=>Variant> = nothing) as stdEnumerator`

Returns a new enumerator by removing duplicate values from `Me`.

If a callback `cb` is given, it will use the return value of the callback for comparison.

#### `Filter(ByVal cb as stdICallable<(item: Variant, key: Variant)=>Boolean>) as stdEnumerator`

Returns a new enumerator containing all elements of `Me` for which the given callback returns `true`.

#### `Concat(ByVal obj as stdEnumerator) as stdEnumerator`

Returns a new enumerator which contains all elements of `Me` and with all elements of `obj` appended on the end.

#### `Join(Optional ByVal sDelimiter as string = ",") as string`

Returns a string created by converting each element of the enumerator to a string, separated by the given delimiter. If the delimiter is missing, it uses `","`.

#### `indexOf(ByVal tv as variant) as long`

Obtains the first index of the value `tv`.

#### `lastIndexOf(ByVal tv as variant) as long`

Obtains the last index of the value `tv`.

#### `includes(ByVal tv as variant) as boolean`

If the enumerator contains `tv` the return value is `true`, else `false` is returned.

#### `reduce(ByVal cb as stdICallable<(accumulator: Variant, item: Variant, key: Variant)=>Variant>, Optional ByVal vInitialValue as variant = 0) as variant`

The `reduce()` method executes a reducer function (that you provide) on each element of the array, resulting in single output value.

The reducer function takes four arguments:

1. Accumulator
2. Current Value
3. Current Index/Key

Your reducer function's returned value is assigned to the accumulator, whose value is remembered across each iteration throughout the array, and ultimately becomes the final, single resulting value.

#### `countBy(ByVal cb as stdICallable<(item: Variant, key: Variant)=>Boolean>) as long`

Counts the number of elements for which the callback `cb` returns `true`.

#### `groupBy(ByVal cb as stdICallable<(item: Variant, key: Variant)=>Variant>) as Dictionary`

Groups the enumerator by result of the callback `cb`. Returns a `Dictionary` where the keys are the evaluated result from the callback and the values are enumerators of elements in the collection that correspond to the key.

#### `max(Optional ByVal cb as stdICallable<(item: Variant, key: Variant)=>Double> = nothing) as variant`

Obtains the maximum value from the enumerator. If a callback is given the item which returns the largest value from the callback is returned.

#### `min(Optional ByVal cb as stdICallable<(item: Variant, key: Variant)=>Double> = nothing) as variant`

Obtains the minimum value from the enumerator. If a callback is given the item which returns the smallest value from the callback is returned.

#### `sum(Optional ByVal cb as stdICallable<(item: Variant, key: Variant)=>Double> = nothing) as variant`

Obtains the sum of all items in the enumerator. If a callback is given the values which are summed are the values returned by the callback function.

#### `Flatten() as stdEnumerator`

Returns a new enumerator that is a one-dimensional flattening of `Me` (recursively).

That is, for every element that is an enumerator, extract its elements into the new enumerator.

#### `Cycle(ByVal iTimes as long, ByVal cb as stdICallable<(item: Variant, key: Variant)=>Variant>) as stdEnumerator`

Calls the given callback `cb` for each element in the enumerator, `iTimes` times.

#### `FindFirst(ByVal cb as stdICallable<(item: Variant, key: Variant)=>Boolean>) as variant`

The `FindFirst()` method returns the value of the first element in the provided enumerator that satisfies the provided testing function. If no values satisfy the testing function, `null` is returned.

#### `checkAll(ByVal cb as stdICallable<(item: Variant, key: Variant)=>Boolean>) as boolean`

Evaluates callback `cb` on all items of the enumerator. If all items return `true`, `true` is returned. Else `false` is returned.

#### `checkAny(ByVal cb as stdICallable<(item: Variant, key: Variant)=>Boolean>) as boolean`

Evaluates callback `cb` on all items of the enumerator. If any of the items return `true`, `true` is returned. Else `false` is returned.

#### `checkNone(ByVal cb as stdICallable<(item: Variant, key: Variant)=>Boolean>) as boolean`

Evaluates callback `cb` on all items of the enumerator. If none of the items return `true`, `true` is returned. Else `false` is returned.

#### `checkOnlyOne(ByVal cb as stdICallable<(item: Variant, key: Variant)=>Boolean>) as boolean`

Evaluates callback `cb` on all items of the enumerator. If only one item returns `true`, `true` is returned. Else `false` is returned.

### PROTECTED METHODS

#### `protInit(ByVal iEnumeratorType as long, ByVal iMaxLength as long, ParamArray v() As Variant)`

Can be used to instantiate the class. Do not use this method unless you know what you are doing.
