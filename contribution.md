Hello, and thanks for looking into contributing to stdVBA! ⭐

In this file I want to document the main ideas behind stdVBA, and critical watchouts to getting your PRs accepted.

## Installation

stdVBA modules can be dragged from windows explorer and dropped into the VBA project. Similarly if you want to export from VBA you can drag the file into explorer, or right click on a class/module and click "Export to file"

## Changelog

The changelog `./changelog.md` is fairly unique for this repository. It lists all major changes to stdVBA. Whenever a breaking change is made it needs to be logged in the changelog. New features and bug-fixes ideally also need to be logged in the changelog to. If you are mainly changing style of a class without changing functionality, changelog entries can be ignored.

## Class-first

stdVBA is class-first meaning that all libraries aim to be a class unless either:

1. They absolutely must be a module.
2. They are supplimentary (i.e. the class can operate without the modules)

Take `stdHTTP` for example. `stdHTTP.cls` is a class which can handle basic authentication mechanisms, however more complex authentication mechanisms can be found in `stdHTTPAuthenticators.bas`.

## Class/Module naming

All classes and modules should be prefixed with `std`. e.g. `stdWindow`

All interfaces should be prefixed with `stdI` e.g. `stdICallable`

## Multi-application

stdVBA is designed as a 1 stop shop for all VBA across all applications `Word`,`Excel`, `PowerPoint`, `Access` etc. Modules should avoid using features which are only available in 1 application, and where available provide generic methods which work across all applications where this isn't available. See `stdWindow.CreateFromApplication` for an example:

```vb
if oApp is nothing then set oApp = Application
select case oApp.Name
  case "Microsoft Excel"
    Set CreateFromApplication = CreateFromHwnd(oApp.hwnd)
  case "Microsoft Word"
    Set CreateFromApplication = CreateFromHwnd(oApp.ActiveWindow.Hwnd)
  case "Microsoft PowerPoint"
    set CreateFromApplication = CreateFromIAccessible(oApp.CommandBars("Status Bar")).AncestralRoot
  case "Microsoft Access"
    set CreateFromApplication = CreateFromHwnd(oApp.hWndAccessApp)
  Case "Outlook"
    Set CreateFromApplication = CreateFromIAccessible(oApp.ActiveWindow.CommandBars("Status Bar")).AncestralRoot
  Case "Microsoft Publisher"
    Set CreateFromApplication = CreateFromIAccessible(oApp.CommandBars("Status Bar")).AncestralRoot
  case else
    Err.Raise 1, "stdAcc::CreateFromApplication()", "No implementation for getting application window of " & Application.name
end select
```

## [Predeclared ID](https://github.com/sancarn/stdVBA/blob/master/src/stdFiber.cls#L8)

This is all about user experience really, with stdVBA you can guarantee when you import a class you can immediately type `stdWhatever.` and see all the methods and properties which come with the class. In many libraries you need to create new instances first, which is often not too clean.

---

```vb
Dim clip as new stdClipboard
clip.text = "hello"
```

vs the `stdVBA` way:

```vb
stdClipboard.text = "hello"
```

---

```vb
Dim proc as new stdProcess
Call proc.init(processID)
Debug.Print proc.name
```

vs the `stdVBA` way:

```vb
Debug.Print stdProcess.CreateFromProcessId(processID).name
```

## Minimise direct-dependencies as much as possible.

Again this is mainly about user experience. An imported class should immediately work without the need of importing another class.

As a result it can lead to messy modules - e.g. stdHTTP being both the [creator of the request](https://github.com/sancarn/stdVBA/blob/master/src/stdHTTP.cls#L216-L219) and the [response of the request](https://github.com/sancarn/stdVBA/blob/master/src/stdHTTP.cls#L336-L385) simultaneously.

> This can lead to lack of code reuse, which im not too happy about, but it makes the user experience great. I'm open to better mechanisms here if anyone has ideas.

### Exception - Generic Cross-Class Interfaces

One exception is the use of a generic cross class interface. `stdICallable` is a good example of this. This 1 interface is used across many of the classes in stdVBA, and merely provides a single interface which many classes can implement as receivers of functionality or producers of functionality.

## All constructors (methods which create an object) will start with the word `Create`.

This bundles all constructors under intellisense. [Example in stdProcess](https://github.com/sancarn/stdVBA/blob/master/src/stdProcess.cls#L230-L314).

## Protected methods

VBA doesn't strictly have protected methods. The closest thing is Friend methods, but even those don't function like traditional protected methods. The constructor methodology described above requires at least 1 protected initialiser (which binds all instance vars to their values).

- All protected methods will be given `prot` prefix, indicating to users that these aren't for them.
- Protected initialisers will have `protInit` prefix. You may have different types of initialisers, see [`stdCOM`](https://github.com/sancarn/stdVBA/blob/master/src/stdCOM.cls#L881-L892)

Constructors will require one public (protected) initialisation method which will at least start with [`protInit`](https://github.com/sancarn/stdVBA/blob/master/src/stdProcess.cls#L327)

## Store state in `TThis`

Historically most of my state has been stored in `pXXXX` variables, but recently I've been migrating all code to the following pattern. It is advisable to keep with this style for contributions.

```vb
Private Type TThis
  State1 as string
  isStateEnabled as boolean
End Type
Private This as TThis
```

Class level state should be stored in a `TSingleton` type within `TThis`:

```vb
Private Type TSingleton
  Cache as object
End Type
Private Type TThis
  Singleton as TSingleton

  State1 as string
  isStateEnabled as boolean
End Type
Private This as TThis
```

## Module-level AI disclosure

Document AI involvement using the [0rwx framework](https://0rwx.com/) in the existing module-level documentation block of each new or changed class/module:

```vb
'@module
'@description Describe the class or module here.
'@ai AI:W
'@aidate 2026-10-31
'@aidesc AI contributed timeout handling, reviewed by the maintainer. The existing transport implementation predates that change and is not attributed to AI by this evidence.
```

Use one `@ai` mark when its classification is established:

- `AI:0`: handmade; AI has neither read nor written the work.
- `AI:R`: AI-inspired; AI has read-only access and a human writes the code.
- `AI:W`: AI collaboration; AI writes code that a human reads and scrutinises.
- `AI:X`: AI-made; a human curates output without being able to scrutinise all generated work.

These marks describe involvement and scrutiny, not percentages or a quality rating. A small AI-written fix does not make the entire class AI-made. For mixed provenance, describe the affected component and the older/upstream work in `@aidesc`. A maintainer may classify a module `AI:W` overall while identifying a specific component as `AI:X`; do not automatically elevate the entire module or attribute its older/upstream code to AI. Do not reset known involvement simply because the latest edit was manual.

- `@aidate YYYY-MM-DD` records the earliest evidenced AI involvement in the class/module described by `@aidesc`, including historical experimental versions when relevant. State when that work is absent from the current implementation. Use an actual edit timestamp where available, otherwise a matched commit date; record the basis and uncertainty in a remark or provenance audit. This is a first known date, not proof that no earlier AI use occurred. Do not advance it for follow-up changes. Omit it for `AI:0` or when no date can be established.
- `@aidesc <description>` explains what AI wrote or changed, what predates those contributions or comes from credited upstream authors, what is not attributed to AI by the evidence, and uncertainty about human scrutiny. Preserve existing descriptions, examples, attributions and member documentation.

For retrospective disclosures, cross-check source comments, `git blame`, dated commit diffs and conversation records. Check deletions as well as surviving lines; a deleted line will not appear in blame. Copied upstream code and unmerged experiments must not be counted as AI-generated current implementation. Dates, model names and commit frequency can guide investigation but cannot establish human scrutiny. A CI bot or the absence of an AI co-author is not evidence of authorship either way.

Do not use `AI:X` as a fallback for unknown review history, or infer `AI:W` merely because a fix is small. Where AI authorship is established but W/X is unresolved, retain `@aidate` and `@aidesc` and leave `@ai` absent until the classification can be supported. Leave modules without sufficient provenance unmarked; do not invent a fifth mark or claim `AI:0`. A disclosure-only edit does not itself establish AI authorship of the implementation. Use each tag at most once per module and regenerate `docs.json` after changing documented source. The generator exposes these as optional `ai`, `aiDate` and `aiDescription` module fields.

## Function-level documentation

A critical area which I am developing is the documentation. All methods should be documented using a `JSDoc`-like equivalent. Docs are analysed by `./Tools/VBDocsGen` (which later will be run on a github action).

````vb
'Search the Window tree for elements which match a certain criteria. Return the first element found.
'@param query as stdICallable<(stdWindow,depth)=>EWndFindResult> - Callback returning `EWndFindResult` options:
'
' _ `EWndFindResult.NoMatchFound`/`0`/`False` - Not found, countinue walking
' _ `EWndFindResult.MatchFound`/`1`/`-1`/`True` - Found, return this element
' _ `EWndFindResult.NoMatchCancelSearch`/`2` - Not found, cancel search
' _ `EWndFindResult.NoMatchSkipDescendents`/`3`, - Not found, don't search descendents
' _ `EWndFindResult.MatchFoundSearchDescendents`/`4` - Same as `EWndFindResult.MatchFound` in this case.
'@param searchType - The type of search, 0 for Breadth First Search (BFS) and 1 for Depth First Search (DFS).
' To understand the difference between BFS and DFS take this tree:
'```
'        A
'       / \
'      B   C
'     /   / \
'    D   E   F
'```
' _ A BFS will walk this tree in the following order: A, B, C, D, E, F
' \* A DFS will walk this tree in a different order: A, C, F, E, B, D
'@param iStaticDepthLimit - Static depth limit. E.G. if you want to search children only, set this value to 1
'@returns - First element found, or `Nothing` if no element found.
'@examples
' ```vb
' 'Find where name is "hello" and class is "world":
' el.FindFirst(stdLambda.Create("$1.name=""hello"" and $1.class=""world"""))
' 'Find first element named "hello" at depth > 4:
' el.FindFirst(stdLambda.Create("$1.name = ""hello"" AND $2 > 4"))
' ```

Public Function FindFirst(ByVal query As stdICallable, Optional ByVal searchType As EWndFindType = EWndFindType.BreadthFirst, Optional ByVal iStaticDepthLimit As Long = -1) As stdWindow
````

### Typical comment layour

```vb
'<description>
'<tags>
<declare>
```

### Key documentation tags:

- `@constructor` - Indicates that a certain method is a constructor.
- `@protected` - Indicates that a certain method is protected.
- `@defaultMember` - Indicate that a certain method is the default member of a class
- `@deprecated` - Indicate that a certain method/property is deprecated
- `@param <name> [as <type>] - <description>` - Document parameter purpose. Type information of parameter only required if type is ambiguous. See `Types` section below.
- `@returns [<type>] - <description>` - Description of returned data. Type information of parameter only required if type is ambiguous. See `Types` section below.
- `@example <markdown>` - Document an example in markdown.
- `@remark <markdown>` - Document a remark / watchout / tip here.
- `@devRemark <markdown>` - Document remarks specifically for devs maintaining the class. This might include source (url) of where you found the algorithm.
- `@TODO: <action description>` - Document actions for future devs to look into.
- `@throws <class>[::|#]<method> - <message>` - Infrequently documented, but sometimes errors will be documented like this. `::` to be used for singleton calls. `#` to be used on instance calls.
- `@complexity <time complexity (average)> - <space complexity>`
- `@requires <library>` - Explicityly declare a dependency.

**Not a 100% sure on this yet, but we might document the classes/modules themselves with**

- `@class` - Class specific documentation

### Repeated calls to the same tag

Repeated calls to the same tag will add an additional tag of that type to the function. I.E.

```vb
'@remark Hello
'@remark Goodbye
```

Will add 2 seperate remarks to the function

### Types

VBA is a typed language therefore documentation of types is usually not required. However there are some cases where you might want to document type information

#### Type is ambiguous

Take the following function:

```vb
Public Function MMult(ByVal m1 as variant, ByVal m2 as variant) as variant

End Function
```

Here we might want to document that `m1` and `m2` are 2d arrays:

```vb
'@param m1 as Variant<Array2D<Float>> - 1st matrix
'@param m2 as Variant<Array2D<Float>> - 2nd matrix
'@returns Variant<Array2D<Float>>
Public Function MMult(ByVal m1 as variant, ByVal m2 as variant) as variant

End Function
```

#### Type can be ambiguous 2

stdCallback can be made from Module functions (where the name of the module is required) or Classes where the object is required. As a result the Parent is of type variant:

```vb
Public Property Get Parent() as Variant
```

In this instance it is important to document this:

```vb
'@returns Variant<String|Object>
Public Property Get Parent() as Variant
```

#### Type is a stdICallable

Type may be a callback, which isn't documentable in VBA; Take `Array#Filter`:

```vb
Public Function Filter(ByVal cb As stdICallable) As stdArray
```

`stdICallable` isn't very descriptive. The reality is we want a callback which, takes a Variant parameter and returns a boolean. We can document this within the documentation as `stdICallable<(element: Variant)=>Boolean>`:

```vb
'Filter the array based on a condition
'@param cb as stdICallable<(element: Variant)=>Boolean> - Callback to run on each element. If the callback returns true, the element is included in the returned array.
'@returns - A new array containing only the elements which passed the filter
Public Function Filter(ByVal cb As stdICallable) As stdArray
```
