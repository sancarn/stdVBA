# `stdLambda`

## Index

* Introduction
    * What is a Lambda expression
* Spec
    * Syntax
        * Expression evaluation
        * Parameter Access
        * Dictionary member access
        * Usage of in-built functions
        * Multi-line lambda expressions.
        * If statements and Inline-if statements
        * Variable definition
        * Function definition
    * Internal Methods and Variables
        * ...
    * Constructors
        * Create()
        * CreateMultiline()
    * stdICallable
        * Run()
        * RunEx()
        * Bind()
        * SendMessage()
    * Instance Methods
        * Bind()
        * BindEx()
        * BindGlobal()
        * Run()
        * RunEx()

## Introduction

### What is a Lambda expression?

A lambda expression/anonymous function is a function definition that is not bound to a name. Lambda expressions are usually "1st class citizens" which means they can be passed to other functions for evaluation.

Lambda expressions are often best described with an example. Imagine we wanted to filter the `Worksheets` object for a Worksheet by Name, Visibility and the Data in `Range("A1")`. Typically in VBA we'd do this as follows:

```vb
Function FilterEnumByName(ByRef col as collection, ByVal sNameLike as string) as collection
  Dim ret as Collection: set ret = new Collection
  Dim item as object
  For each item in col
    if item.name like sNameLike then
      ret.add item
    end if
  next 
  set FilterEnumByName = ret
End Function
Function FilterEnumByVisible(ByRef col as collection, ByVal bVisible as Boolean) as collection
  Dim ret as Collection: set ret = new Collection
  Dim item as object
  For each item in col
    if item.name = bVisible then
      ret.add item
    end if
  next 
  set FilterEnumByVisible = ret
End Function
Function FilterEnumByRange(ByRef col as collection, ByVal v as Variant) as collection
  Dim ret as Collection: set ret = new Collection
  Dim item as object
  For each item in col
    if item.range("A1").value = v then
      ret.add item
    end if
  next 
  set FilterEnumByRange = ret
End Function
```

In this scenario we have made 3 seperate functions, and each function is operationally, virtually the same but performing slightly different operations. Furthermore the operation of each function is extremely limited. This also breaks the [DRY](https://en.wikipedia.org/wiki/Don%27t_repeat_yourself) principle in programming.

Wouldn't it be better if we could get all that power while only writing a single function? Well, we can, with lambda expressions:

```vb
Sub Test
  'Filter by name
  set col = FilterEnum(col, stdLambda.Create("$1.name like ""Hello*"""))

  'Filter by Visible
  set col = FilterEnum(col, stdLambda.Create("not $1.Visible"))

  'Filter by Range
  set col = FilterEnum(col, stdLambda.Create("$1.Range(""A1"").value = ""Test"""))
End Sub

Function FilterEnum(ByRef col as Collection, ByVal lambda as stdICallable) as Collection
  Dim ret as Collection: set ret = new Collection
  Dim item as object
  For each item in col
    if lambda.run(item) then
      ret.add item
    end if
  next 
  set FilterEnumByRange = ret
End Function
```

Bare in mind, however, that this new function is significantly more powerful! If you suddenly also need to filter where `range("A3").value < 10` this is something that can be easily applied:

```vb
'...
set col = FilterEnum(col, stdLambda.Create("$1.Range(""A3"").Value < 10"))
'...
```

This is the role that Lambda expressions provide, and is what `stdLambda` provides within VBA.

One of the core use cases of `stdLambda` in `stdVBA` can be found in `stdArray`, an array class declared in `src/stdArray.cls`. The following examples can be used to see the use of the `stdLambda` class with `stdArray`:

```vb
'Create an array
Dim arr as stdArray
set arr = stdArray.Create(1,2,3,4,5,6,7,8,9,10) 'Can also call CreateFromArray

'More advanced behaviour when including callbacks! And VBA Lamdas!!
Debug.Print arr.Map(stdLambda.Create("$1+1")).join          '2,3,4,5,6,7,8,9,10,11
Debug.Print arr.Reduce(stdLambda.Create("$1+$2"))           '55 ' I.E. Calculate the sum
Debug.Print arr.Reduce(stdLambda.Create("Max($1,$2)"))      '10 ' I.E. Calculate the maximum
Debug.Print arr.Filter(stdLambda.Create("$1>=5")).join      '5,6,7,8,9,10

'Execute property accessors with Lambda syntax
Debug.Print arr.Map(stdLambda.Create("ThisWorkbook.Sheets($1)")) _ 
               .Map(stdLambda.Create("$1.Name")).join(",")            'Sheet1,Sheet2,Sheet3,...,Sheet10

'Execute methods with lambda:
Call stdArray.Create(Workbooks(1),Workbooks(2)).forEach(stdLambda.Create("$1.Save")

'We even have if statement!
With stdLambda.Create("if $1 then ""lisa"" else ""bart""")
  Debug.Print .Run(true)                                              'lisa
  Debug.Print .Run(false)                                             'bart
End With
```

## Spec

### Syntax

One of the core additions in `stdLambda` is the custom syntax/language embedded in VBA strings.

#### Expression Evaluation

`stdLambda`'s primary focus is expression evaluation.

#### Parameter Access


To define a function which takes multiple arguments $# should be used where # is the index of the argument. E.G. $1 is the first argument, $2 is the 2nd argument and $n is the nth argument.

```vb
Sub test()
    Dim average as stdLambda
    set average = stdLambda.Create("($1+$2)/2")
End Sub
```

You can also define functions which call members of objects. E.G.

```vb
Sub test()
    'Call properties or methods with `.`
    Debug.Print stdLambda.Create("$1.Name")(ThisWorkbook)  'returns ThisWorkbook.Name
    Call stdLambda.Create("$1.Save")(ThisWorkbook)         'calls ThisWorkbook.Save

    'If you absolutely need to use `VbGet` use `.$` or if you need to use `VbMethod` use `.#` instead.
    Debug.Print stdLambda.Create("$1.$Name")(ThisWorkbook)  'returns ThisWorkbook.Name
    Call stdLambda.Create("$1.#Save")(ThisWorkbook)         'calls ThisWorkbook.Save
End Sub

```

#### Dictionary member access

If a dictionary is passed into a `stdLambda` then you can use `Dictionary.Key` syntax to access members:

```vb
'Both:
stdLambda.create("$1.item(""someVar"")").run(myDict)
'and
stdLambda.create("$1.someVar").run(myDict)
'return the same results
```

#### Usage of in-built functions

The lambda syntax comes with many in-built functions many of which call directly to native VBA functions. These can greatly help productivity.

```vb
Sub test()
    Debug.Print stdLambda.Create("Mid($1,1,5)")("hello world")        'returns "hello"
    Debug.Print stdLambda.Create("$1 like ""hello*""")("hello world") 'returns true
End Sub
```

#### Multi-line lambda expressions.

Lambda expressions may be defined as multi-line expressions using either the `:` end-of-line character or using `CreateMultiLine` constructor. In the cases where a mutli-line function is used, the last expression evaluated is returned from the lambda. I.E. the following expressions will return `10`:

```vb
Call stdLambda.Create("2+2: 5*2").Run()

'... or ...

Call stdLambda.CreateMultiline(array( _ 
  "2+2", _ 
  "5*2" 
)).Run()
```

#### If statements and Inline-if statements 

`stdLambda` has an if statement as well as an inline-if statement. The if statement looks as follows:

```vb
With stdLambda.CreateMultiline(array( _ 
  "if $1 = 1 then", _ 
  "  5*1", _ 
  "else if $1 = 2 then", _ 
  "  5*2", _ 
  "else", _
  "  5*3", _ 
  "end" 
))
  Debug.Print .Run(0) '15
  Debug.Print .Run(1) '5
  Debug.Print .Run(2) '10
  Debug.Print .Run(3) '15
End With
```

The inline-if acts as a regular if statement, It will only execute the `then ...` part or `else ...` part, but it will 'return' the result of the executed block. This means it can be used inline like vba's iif can be used, but it doesn't have to compute both the `then` and `else` part like `iif` does.

```vb
Sub test()
    Debug.Print stdLambda.Create("if $1 then 1 else 2")(true)        'returns 1
    Debug.Print stdLambda.Create("if $1 then 1 else 2")(false)       'returns 2
End Sub

' only evaluates doSmth(), does not evaluate doAnother() when $1 is true, and visa versa
stdLambda.Create("(if $1 then doSmth() else doAnother() end)").Run(True) 
```

> note: if statements will only evaluate the part which is required. This is extremely beneficial and far superior in comparrison to `iif()`

#### Variable definition

Variables can be defined and assigned, e.g. `oranges = 2`. This can make definition of extensive formula easier to read. Assignment results in their value.

```vb
'the last assignment is redundant, just used to show that assignments result in their value
Debug.Print stdLambda.CreateMultiline(array( _
  "count = $1", _
  "footPrint = count * 2 ^ count" _
)).Run(2) ' -> 8
```

#### Function definition

You can also define functions:

```vb
stdLambda.CreateMultiline(Array( _
  "fun fib(v)", _
  "  if v<=1 then", _
  "    v", _
  "  else ", _
  "    fib(v-2) + fib(v-1)", _
  "  end", _
  "end", _
  "fib($1)" _
)).Run(20) '->6765
```

### Watch-outs

* Currently the main "caveat" to using this library is performance. This will absolutely not be as fast as pure VBA code and doesn't intend to be.
* Currently there is a lack of error handling.
* Lack of intellisense or syntax highlighting. This can be fixed in IDEs like VSCode.

### Built in functions and variables

In `stdLambda` functions and variables are treated as the same thing. Variable names which don't exist in an assignment, won't make it onto the variable table, and thus won't make it onto the stack, and are interpreted as `iType.oFunc` types instead of `iType.oAccess`. At runtime `iType.oFunc` code is passed to the `evaluateFunc` function with arguments also received from the stack, if a variable is evaluated with `evaluateFunc`, typically no arguments are passed. Within `evaluateFunc`, first the function extension table is checked to see if the variable name exists in there, otherwise the variable name is looked up in a large select-case tree.

Co-incidentally `BindGlobal` just adds to the function extension table in `evaluateFunc`. Also this means that `bindGlobal` can be used to bind a function to a lambda expression, and also can be used to override built-in functions and variables.

All function names in `stdLambda` are case insensitive, a design decision to make the language more like VBA.

#### `ThisWorkbook`

Evaluates to the workbook which instantiated the `stdLambda` class.

##### Example:

```vb
ThisWorkbook.name
```

#### `Application`

Evaluates to the Excel/Word/Powerpoint/App's `Application` object.

##### Example:

```vb
if Application.version < 15 then 1 else 2
```


#### `eval(expression as string) as variant`

Evaluates an expression passed as a string. Note: Internally `eval` uses `stdLambda`. The definition in source is `stdLambda.Create(firstArg).Run()`

##### Example:

```vb
eval("1+1")   '==> 2
```

#### `Abs(val as double) as double`

A call to VBA's `Abs()` function.

#### `int(val as variant) as integer`

A call to VBA's `Int()` function.

#### `fix(val as variant) as double`

A call to VBA's `Fix()` function.

#### `exp(val as double) as double`

A call to VBA's `Exp()` function.

#### `log(val as double) as double`

A call to VBA's `Log()` function.

#### `sqr(val as double) as double`

A call to VBA's `Sqr()` function.

#### `sgn(val as double) as double`

A call to VBA's `Sgn()` function.

#### `rnd(val as double) as double`

A call to VBA's `Rnd()` function.

#### `cos(angle as variant) as double`

A call to VBA's `Cos()` function. `angle` should be defined in radians.

#### `sin(angle as variant) as double`

A call to VBA's `Sin()` function. `angle` should be defined in radians.

#### `tan(angle as variant) as double`

A call to VBA's `Tab()` function. `angle` should be defined in radians.

#### `atn(angle as variant) as double`

A call to VBA's `Atn()` function. Calculates the Inverse tangent. `angle` should be defined in radians.

#### `asin(angle as variant) as double`

Calculates the Inverse sin. `angle` should be defined in radians.

#### `acos(angle as variant) as double`

Calculates the Inverse cosine. `angle` should be defined in radians.

#### `Array`

Creates an array from the supplied parameters:

```vb
Array(1,2,3) '--> Array(1,2,3)
```

#### `CreateObject(class as string, optional server as string) as object`

A call to VBA's `CreateObject()` function. 

```vb
CreateObject("Scripting.Dictionary")
```

#### `GetObject`

A call to VBA's `GetObject()` function. 

```vb
GetObject("InternetExplorer.Application")
```

#### `iff(cond as boolean, valIfTrue as variant, valIfFalse as variant) as variant`

A call to VBA's `Iff()` function. It is suggested that you avoid calls to this function and instead use the in-line if of `stdLambda`.

#### `TypeName`

A call to VBA's `TypeName(v as variant) as string` function. 

#### `CBool(v as variant) as boolean`

A call to VBA's `cbool` function. 

#### `CByte(v as variant) as Byte`

A call to VBA's `cbyte` function. 

#### `CCur(v as variant) as Currency`

A call to VBA's `ccur` function. 

#### `CDate(v as variant) as Date`

A call to VBA's `cdate` function. 

#### `CSng(v as variant) as Single`

definitionA call to VBA's `csng` function. 

#### `CDbl(v as variant) as Double`

A call to VBA's `cdbl` function. 

#### `CInt(v as variant) as Integer`

A call to VBA's `cint` function. 

#### `CLng(v as variant) as Long`

A call to VBA's `clng` function. 

#### `CStr(v as variant) as String`

A call to VBA's `cstr` function. 

#### `CVar(v as variant) as Variant`

A call to VBA's `cvar` function. 

#### `CVErr(errnum as long) as Error`

A call to VBA's `cverr` function. 

#### `Asc`

A call to VBA's `Asc` function. 

#### `Chr`

A call to VBA's `Chr` function. 

#### `Format`

A call to VBA's `Format` function. 

#### `Hex`

A call to VBA's `Hex` function. 

#### `Oct`

A call to VBA's `Oct` function. 

#### `Str`

A call to VBA's `Str` function. 

#### `Val`

A call to VBA's `Val` function. 

#### `Trim`

A call to VBA's `Trim` function. 

#### `LCase`

A call to VBA's `LCase` function. 

#### `UCase`

A call to VBA's `UCase` function. 

#### `Right`

A call to VBA's `Right` function. 

#### `Left`

A call to VBA's `Left` function. 

#### `Mid`

A call to VBA's `Mid` function. 

#### `Len`

A call to VBA's `Len` function. 

#### `Now`

A call to VBA's `Now` function. 

#### `Switch`

A call to VBA's `Switch` function. 

#### `Any`

A call to VBA's `Any` function. 


#### `vbCrLf`

Returns a carriage-return line-feed.

#### `vbCr`

Returns a carriage-return.

#### `vbLf`

Returns a line-feed.

#### `vbNewline`

Returns a carriage-return line-feed.

#### `vbNullChar`

Returns a null character.

#### `vbNullString`

Returns a null string.

#### `vbObjectError`

Equivalent of `vbObjectError` in VBA

#### `vbTab`

Returns a tab character.

#### `vbBack`

Returns a backspace character.

#### `vbFormFeed`

Returns a form feed character

#### `vbVerticalTab`

Returns a vertical tab character

### Constructors

#### `Create(ByVal sEquation As String, Optional ByVal bUsePerformanceCache As Boolean = False, Optional ByVal bSandboxExtras As Boolean = False) As stdLambda`

Creates and returns a `stdLambda` object which will execute the supplied equation body, when run.

```vb
Debug.Print stdLambda.Create("1+3*8/2*(2+2+3)").Execute()
With stdLambda.Create("$1+1+3*8/2*(2+2+3)")
    Debug.Print .Execute(10)
    Debug.Print .Execute(15)
    Debug.Print .Execute(20)
End With
Debug.Print stdLambda.Create("$1.Range(""A1"")").Execute(Sheets(1)).Address(True, True, xlA1, True)
Debug.Print stdLambda.Create("$1.join("","")").Execute(stdArray.Create(1,2))
```

Use `bUsePerformanceCache` when looping over a large dataset (e.g. many rows of a table) and filtering on something very specific with little variance. Example:

```vb
'Data like
'| ID | Type    | Status   | ...
'|----|---------|----------| ...
'| 1  | Message | Archived | ...
'| 2  | Note    | Active   | ...
'| 3  | Message | Active   | ...
'| 4  | Message | Active   | ...
'| 5  | Note    | Active   | ...
'...

'Assuming our filterByColumns function passes only the columns values into the lambda as parameters, the following will be much faster:
myTable.filterByColumns(Array("Type", "Status"), stdLambda.Create("$1=""Message"" and $2=""Active""", true))
'than
myTable.filterByColumns(Array("Type", "Status"), stdLambda.Create("$1=""Message"" and $2=""Active"""))
```

**Note:**

Avoid using Performance cache when lambda is called with very dynamic arguments and/or on small sets of data. E.G.

```vb
    'Data like
    '| Freq  | Risk    | ...
    '|-------|---------| ...
    '| 1.2   | 5.0     | ...
    '| 2.1   | 5.0     | ...
    '| 34.2  | 2.4     | ...
    '| 43.2  | 2.7     | ...
    '| 2.0   | 3.1     | ...
    '...

    'This is a poor use of performance cache because all input parameters are different from one call to the next.
    'Therefore no performance gains will be observed, and rather performance decreases are likely due to memory consumption.
    table.filterByColumns(Array("Freq", "Risk"), stdLambda.Create("$1*$2 > 60",true))


    'Data like
    '| ID | Type    | Status   | ...
    '|----|---------|----------| ...
    '| 1  | Message | Archived | ...
    '| 2  | Note    | Active   | ...
    '| 3  | Message | Active   | ...
    '|----|---------|----------| ...
    'Again this is a poor use of cache, as the table is so small. Performance cache is mainly only useful on large datasets.
    myTable.filterByColumns(Array("Type", "Status"), stdLambda.Create("$1=""Message"" and $2=""Active""", true))

    'Data like
    '| ID | Type    | Status   | ...
    '|----|---------|----------| ...
    '| 1  | Message | Archived | ...
    '| 2  | Note    | Active   | ...
    '| 3  | Message | Active   | ...
    '| 4  | Message | Active   | ...
    '| 5  | Note    | Active   | ...
    '...

    'Unfortunately this is also a poor use of performance cache. $1 is the only argument and this will be a different row each time
    'this is called. Therefore no performance benefits will be observed.
    'Note: There is room to improve this behaviour at a later date.
    myTable.filter(stdLambda.Create("$1.Type=""Message"" and $1.Status=""Active""", true))
```

Use `bSandboxExtras` when you want strict control over the functions the user can call.

#### `CreateMultiline(ByRef sEquation As Variant, Optional ByVal bUsePerformanceCache As Boolean = False, Optional ByVal bSandboxExtras As Boolean = False) As stdLambda`

Creates and returns a `stdLambda` object which will execute the supplied equation body, when run.

```vb
stdLambda.CreateMultiline(Array( _ 
  "if $1 = 0 then ""Test1""", _ 
  "else if $1 = 1 then ""Test2""", _ 
  "else ""Test3""" _  
))
```

See `Create` for usage of `bUsePerformanceCache` and `bSandboxExtras`

### Implemented Interfaces

#### `stdICallable`

`stdICallable` is shared between `stdLambda` and `stdCallback` in `stdVBA`, however the intention is anyone can implement this same interface. Bare in mind that `SendMessage` will be implemented differently for each system.

```vb
'Call will call the passed function with param array
Public Function Run(ParamArray params() as variant) as variant: End Function

'Call function with supplied array of params
Public Function RunEx(ByVal params as variant) as variant: End Function

'Bind a parameter to the function
Public Function Bind(ParamArray params() as variant) as stdICallable: End Function

'Making late-bound calls to stdICallable members
'@protected
'@param {ByVal String} - Message to send
'@param {ByRef Boolean} - Whether the call was successful
'@param {ByVal Variant} - Any variant, typically parameters as an array. Passed along with the message.
'@returns {Variant} - Any return value.
Public Function SendMessage(ByVal sMessage as string, ByRef success as boolean, ByVal params as variant) as Variant: End Function
```

In this scenario 3 methods are implemented onto `stdLambda`:

```
obj          - returns the object itself, can be used for casting to base type.
className    - returns the object's class name.
bindGlobal   - ICallable alias for bindGlobal.
```

### Instance Methods

#### `Bind(ParamArray v())`

The `bind()` method creates a new ICallable that, when called, supplies the given sequence of arguments preceding any provided when the new function is called.

```vb
set lambda1 = stdLambda.Create("$1.name")
Debug.Print lambda1(ThisWorkbook)

set lambda2 = lambda1.bind(ThisWorkbook)
Debug.Print lambda2()                     'same result as above
```

Typically this is used when looping through a set of data, while relating back to something else. It is also significantly more optimal than continual recompiles. For instance:

```vb
For i = 1 to 10
    'FAST - no need to continually recompile
    stdArray.Create(1,2,3).map(stdLambda.Create("$2 / $1").bind(i)) 'note this is like "$1 / i"

    'SLOW - Recompile on each loop
    stdArray.Create(1,2,3).map(stdLambda.Create("$1 / " & i))
Next

'another example of usage:
set GetRecordsByDate = records.filter(stdLambda.Create("$2.Date = $1").bind(dt))
```

#### `BindEx(vArray as variant)`

Equivalent of `Bind()`, but with a passed array instead of a param array.

#### `BindGlobal(sName as string, vFunctionOrVariable as variant)`

`BindGlobal` will bind a global function or variable onto the lambda it's called against. In order to bind a function `vFunctionOrVariable` must be an object which implements `stdICallable`.

> Note: globals can be bound to `stdLambda` base class, e.g. `stdLambda.bindGlobal("superGlobalCollection", new Collection)`, which will add the function to all newly generated lambdas.

> Note: `stdCallback` can be used to produce an `stdICallable` from a Module or Class function.
```vb
'Typical usage of bind as an enum-style type:
Sub test1Main()
  Debug.Print test1(stdLambda.Create("Status.Red")) = 1 'True
  Debug.Print test1(stdLambda.Create("Status.Green")) = 3 'True
End Sub
Function test1(lambda as stdLambda) as long
  Static oStatus as object
  if oStatus is nothing then
    set oStatus = CreateObject("Scripting.Dictionary")
    oStatus("Red") = 1
    oStatus("Amber") = 2
    oStatus("Green") = 3
  end if
  lambda.bindGlobal("Status", oStatus)

  test1 = lambda.run()
End Function

'Usage to assign custom functions to a `stdLambda`.
Sub test2Main()
  Debug.print test2(stdLambda.Create("addOne(21)")) = 22   'True
End Sub
Function test1(lambda as stdLambda) as long
  Static addOne as stdICallable: if addOne is nothing then set addOne = stdLambda.Create("$1+1")
  lambda.bindGlobal("addOne", addOne)
  test2 = lambda.run()
End Function
```

#### `Run(paramarray v()) as variant`

Evaluates the lambda expression.

#### `RunEx(vArray as Variant) as variant`

See `Run`. Takes an array of parameters instead of a paramarray.
