# stdVBA

A Collection of libraries to form a common standard layer for modern VBA applications.

## Benefits

* Code faster!
* Improve code maintainability.
* Let the library handle the complicated stuff, you focus on the process 
* Heavily inspired by JavaScript APIs - More standard
* Open Source - Means the libraries are continually maintained by the community. Want something added, help us make it!

[The full roadmap](https://github.com/sancarn/stdVBA/projects/1) has more detailed information than here.

## Short example

```vb
sub Main()
  'Create an array
  Dim arr as stdArray
  set arr = stdArray.Create(1,2,3,4,5,6,7,8,9,10) 'Can also call CreateFromArray

  'Demonstrating join, join will be used in most of the below functions
  Debug.Print arr.join()                                                 '1,2,3,4,5,6,7,8,9,10
  Debug.Print arr.join("|")                                              '1|2|3|4|5|6|7|8|9|10

  'Basic operations
  arr.push 3
  Debug.Print arr.join()                                                 '1,2,3,4,5,6,7,8,9,10,3
  Debug.Print arr.pop()                                                  '3
  Debug.Print arr.join()                                                 '1,2,3,4,5,6,7,8,9,10
  Debug.Print arr.concat(stdArray.Create(11,12,13)).join                 '1,2,3,4,5,6,7,8,9,10,11,12,13
  Debug.Print arr.join()                                                 '1,2,3,4,5,6,7,8,9,10 'concat doesn't mutate object
  Debug.Print arr.includes(3)                                            'True
  Debug.Print arr.includes(34)                                           'False

  'More advanced behaviour when including callbacks! And VBA Lamdas!!
  Debug.Print arr.Map(stdLambda.Create("$1+1")).join          '2,3,4,5,6,7,8,9,10,11
  Debug.Print arr.Reduce(stdLambda.Create("$1+$2"))           '55 ' I.E. Calculate the sum
  Debug.Print arr.Reduce(stdLambda.Create("Max($1,$2)"))      '10 ' I.E. Calculate the maximum
  Debug.Print arr.Filter(stdLambda.Create("$1>=5")).join      '5,6,7,8,9,10
  
  'Execute property accessors with Lambda syntax
  Debug.Print arr.Map(stdLambda.Create("ThisWorkbook.Sheets($1)")) _ 
                 .Map(stdLambda.Create("$1.Name")).join(",")            'Sheet1,Sheet2,Sheet3,...,Sheet10
  
  'Execute methods with lambdas and enumerate over enumeratable collections:
  Call stdEnumerator.Create(Application.Workbooks).forEach(stdLambda.Create("$1#Save")
  
  'We even have if statement!
  With stdLambda.Create("if $1 then ""lisa"" else ""bart""")
    Debug.Print .Run(true)                                              'lisa
    Debug.Print .Run(false)                                             'bart
  End With
  
  'Execute custom functions
  Debug.Print arr.Map(stdCallback.CreateFromModule("ModuleMain","CalcArea")).join  '3.14159,12.56636,28.274309999999996,50.26544,78.53975,113.09723999999999,153.93791,201.06176,254.46879,314.159

  'Let's move onto regex:
  Dim oRegex as stdRegex
  set oRegex = stdRegex.Create("(?<county>[A-Z])-(?<city>\d+)-(?<street>\d+)","i")

  Dim oRegResult as object
  set oRegResult = oRegex.Match("D-040-1425")
  Debug.Print oRegResult("county") 'D
  Debug.Print oRegResult("city")   '040
  
  'And getting all the matches....
  Dim sHaystack as string: sHaystack = "D-040-1425;D-029-0055;A-100-1351"
  Debug.Print stdEnumerator.CreateFromEnumVARIANT(oRegex.MatchAll(sHaystack)).map(stdLambda.Create("$1.item(""county"")")).join 'D,D,A
  
  'Dump regex matches to range:
  '   D,040,040-1425
  '   D,029,029-0055
  '   A,100,100-1351
  Range("A3:C6").value = oRegex.ListArr(sHaystack, Array("$county","$city","$city-$street"))
  
  'Copy some data to the clipboard:
  Range("A1").value = "Hello there"
  Range("A1").copy
  Debug.Print stdClipboard.Text 'Hello there
  stdClipboard.Text = "Hello world"
  Debug.Print stdClipboard.Text 'Hello world

  'Copy files to the clipboard.
  Dim files as collection
  set files = new collection
  files.add "C:\File1.txt"
  files.add "C:\File2.txt"
  set stdClipboard.files = files

  'Save a chart as a file
  Sheets("Sheet1").ChartObjects(1).copy
  Call stdClipboard.Picture.saveAsFile("C:\test.bmp",false,null) 'Use IPicture interface to save to disk as image
End Sub

Public Function CalcArea(ByVal radius as Double) as Double
  CalcArea = 3.14159*radius*radius
End Function
```

## Motivation

VBA first appeared in 1993 (over 25 years ago) and the language's age is definitely noticable. VBA has a lot of specific libraries for controlling Word, Excel, Powerpoint etc. However the language massively lacks in generic modern libraries for accomplishing common programming tasks. VBA projects ultimately become a mish mash of many different technologies and programming styles. Commonly for me that means calls to Win32 DLLs, COM libraries via late-binding, calls to command line applications and calls to .NET assemblies.

Over time I have been building my own libraries and have gradually built my own layer above the simple VBA APIs.

The VBA Standard Library aims to give users a set of common libraries, maintained by the community, which aid in the building of VBA Applications.

## Road Map

This project is has been majorly maintained by 1 person, so progress is generally very slow. This said, generally the road map corresponds with what I need at the time, or what irritates me. In general this means `fundamental` features are more likely to be complete first, more complex features will be integrated towards the end. This is not a rule, i.e. `stdSharepoint` is mostly complete without implementation of `stdXML` which it'd use. But as a general rule of thumb things will be implemented in the following order:

* Types - `stdArray`, `stdDictionary`, `stdRegex`, `stdDate`, `stdLambda`, ... 
* Data  - `stdJSON`, `stdXML`, `stdOXML`, `stdCSON`, `stdIni`, `stdZip` 
* File  - `stdShell` 
* Automation - `stdHTTP`, `stdAcc`, `stdWindow`, `stdKernel`
* Excel specific - `xlFileWatcher`, `xlProjectBuilder`, `xlTimer`, `xlShapeEvents`, `xlTable`
* Runtimes - `stdCLR`, `stdPowershell`, `stdJavascript`, `stdOfficeJSBridge`

## Planned Global Objects:

<!--
  docs/assets/Status_G.png - Ready
  docs/assets/Status_Y.png - WIP
  docs/assets/Status_R.png - Hold
-->


|Color                                                     | Status | Type       |Name             | Description  |
|----------------------------------------------------------|--------|------------|-----------------|--------------|
|![l](docs/assets/Status_G.png) | READY  | Debug      |stdError         | Better error handling, including stack trace and error handling diversion and events.
|![l](docs/assets/Status_G.png) | READY  | Type       |stdArray         | A library designed to re-create the Javascript dynamic array object.
|![l](docs/assets/Status_G.png) | READY  | Type       |stdEnumerator    | A library designed to wrap enumerable objects providing additional functionality.
|![l](docs/assets/Status_Y.png) | WIP    | Type       |stdDictionary    | A drop in replacement for VBScript's dictionary.
|![l](docs/assets/Status_G.png) | READY  | Type       |stdDate          | A standard date parsing library. No more will you have to rely on Excel's interpreter. State the format, get the data.
|![l](docs/assets/Status_G.png) | READY  | Type       |stdRegex         | A regex library with more features than standard e.g. named capture groups and free-spaces.
|![l](docs/assets/Status_G.png) | READY  | Type       |stdLambda        | Build and create in-line functions. Execute them at a later stage.
|![l](docs/assets/Status_G.png) | READY  | Type       |stdCallback      | Link to existing functions defined in VBA code, call them at a later stage.
|![l](docs/assets/Status_G.png) | READY  | Type       |stdCOM           | A wrapper around a COM object which provides Reflection (through ITypeInfo), Interface querying, Calling interface methods (via DispID) and more. 
|![l](docs/assets/Status_G.png) | READY  | Automation |stdClipboard     | Clipboard management library. Set text, files, images and more to the clipboard.
|![l](docs/assets/Status_R.png) | HOLD   | Automation |stdHTTP          | A wrapper around Win HTTP libraries.
|![l](docs/assets/Status_G.png) | READY  | Automation |stdWindow        | A handy wrapper around Win32 Window management APIs.
|![l](docs/assets/Status_G.png) | READY  | Automation |stdProcess       | Create and manage processes.
|![l](docs/assets/Status_G.png) | READY  | Automation |stdAcc           | Use Microsoft Active Accessibility framework within VBA - Very useful for automation.
|![l](docs/assets/Status_Y.png) | WIP    | Excel      |xlTable          | Better tables for VBA, e.g. Map rows etc.

[The full roadmap](https://github.com/sancarn/stdVBA/projects/1) has more detailed information than here.

### Statuses

#### ![_](docs/assets/Status_G.png) READY

APIs which are ready to use, and although are not fully featured are in a good enough working state.

#### ![_](docs/assets/Status_Y.png) WIP

APIs which are WIP are not necessarily being worked on currently but at least are recognised for their importance to the library. These will be lightly worked on/thought about continuously even if no commits are made.

As of Oct 2020, this status typically consists of:
* data types, e.g. stdEnumerator, stdDictionary, stdTable;
* Unit testing; 
* Tasks difficult to automate otherwise e.g. stdClipboard, stdAccessibility;

#### ![_](docs/assets/Status_R.png) HOLD

APIs where progress has been temporarily halted, and/or is currently not a priority.

In the early days we'll see this more with things which do already have existing work arounds and are not critical, so projects are more likely to fit into this category.

#### ![_](https://via.placeholder.com/15/aaaaaa/000000?text=+) UNK

APIs which have been indefinitely halted. We aren't sure whether we need these or if they really fit into the project. They are nice to haves but not necessities for the project as current. These ideas may be picked up later. All feature requests will fit into this category initially.

## Structure

All modules or classes will be prefixed by `std` if they are generic libraries.

Application specific libraries to be prefixed with `xl`, `wd`, `pp`, `ax` representing their specific application.

Commonly implementations will use the factory class design pattern:

```vb
Class stdClass
  Private bInitialised as boolean

  'Creates an object from the given parameters
  '@constructor
  Public Function Create(...) As stdClass
    if not bInitialised then
      Set Create = New stdClass
      Call Create.init(...)
    else
      Call CriticalRaise("Constructor called on object not class")
    End If
  End Function

  'Initialises the class. This method is meant for internal use only. Use at your own risk.
  '@protected
  Public Sub init(...)
    If bInitialised Then
      Call CriticalRaise("Cannot run init() on initialised object")
    elseif Me is stdClass then
      Call CriticalRaise("Cannot run init() on static class")
    else
      'initialise with params...

      'Make sure bInitialised is set
      bInitialised=true
    End If
  End Sub

  Private Sub CriticalRaise(ByVal sMsg as string)
    if isObject(stdError) then
      stdError.Raise sMsg
    else
      Err.Raise 1, "stdClass", sMsg
    end if
  End Sub
  
  '...
End Class
```

With the above example, the Regex class is constructed with the `Create()` method, which can only be called on the `stdRegex` static class. We will try to keep this structure across all STD VBA classes.

# Contributing

If you are looking to contribute to the VBA standard library codebase, the best place to start is the [GitHub "issues" tab](https://github.com/sancarn/VBA-STD-Library/issues). This is also a great place for filing bug reports and making suggestions for ways in which we can improve the code and documentation. A list of options of different ways to contribute are listed below:

* If you have a Feature Request - Create a new issue
* If you have found a bug - Create a new issue
* If you have written some code which you want to contribute see the Contributing Code section below.

## Contributing Code

There are several ways to contribute code to the project:

* Opening pull requests is the easiest way to get code intergrated with the standard library.
* Create a new issue and providing the code in a code block - Bare in mind, it will take us a lot longer to pick this up than a standard pull request.

Please make sure code contributions follow the following guidelines:

* `stdMyClass.cls` should have `Attribute VB_PredeclaredId = True`. 
* `Attribute VB_Name` should follow the STD convention e.g. `"stdMyClass"`
* Follow the STD constructor convention `stdMyClass.Create(...)`.
* Ensure there are plenty of comments where required.
* Ensure lines end in `\r\n` and not `\n` only.

As long as these standard conventions are met, the rest is up to you! Just try to be as general as possible! We're not necessarily looking for optimised code, at the moment we're just looking for code that works!

> Note: Ensure that all code is written by you. If the code is not written by you, you will be responsible for any repercussions!

## Inspiration documents

Inspiration was initially stored in this repository, however the vast swathes of examples, knowledge and data became too big for this repository, therefore it was moved to:

https://github.com/sancarn/VBA-STD-Lib-Inspiration
