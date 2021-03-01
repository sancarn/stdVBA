# `stdAcc`

This merge request is to finally lift `stdAcc` from WIP into SRC with the new query system based on `stdICallable` interface.

## Spec

### Constructors

#### `CreateFromPoint(X,Y)`

Creates an `stdAcc` object from an `X` and `Y` point location on the screen.

```vb
Debug.Print TypeName(stdAcc.CreateFromPoint(0,0))
```

#### `CreateFromHwnd(hwnd)`

Creates an `stdAcc` object from a window handle.

```vb
Debug.Print TypeName(FindWindowA(MSH_WHEELMODULE_CLASS,MSH_WHEELMODULE_TITLE))
```

#### `CreateFromApplication()`

Creates an `stdAcc` object from the current running application (e.g. Excel / Word / Powerpoint).

```vb
'Print name of app window
Debug.Print stdAcc.CreateFromApplication().Name
```

> Note: Implementation as current relies on `Application.hwnd`. An application agnostic method is required. As such this method is only guaranteed to work in Excel.

#### `CreateFromDesktop()`

Creates an `stdAcc` object from the desktop.

```vb
'Loop over all windows
Dim accWnd as stdAcc
For each accWnd in stdAcc.CreateFromDesktop().children
  Debug.print accWnd.name
next
```

> Note: Implementation as current relies `stdAcc.CreateFromApplication()`, and as such has the same limitations.

#### `CreateFromIAccessible(IAccessible)`

Creates an `stdAcc` object from an object which implements `IAccessible`.

```vb
Dim obj As IAccessible
Dim v As Variant
Call AccessibleObjectFromPoint(x, y, obj, v)
Debug.Print CreateFromIAccessible(obj).name
```

#### `CreateFromMouse()`

Creates an `stdAcc` object for the element the mouse currently hovers over.

```vb
While True
    Dim obj as stdAcc
    set obj = stdAcc.CreateFromMouse()
    if not obj is nothing then Debug.Print obj.Name & " - " & obj.Role 
    DoEvents
Wend
```

#### `CreateFromPath(sPath)`

Creates an `stdAcc` object for the element at a given path from the current element.

```vb
Debug.Print stdAcc.CreateFromApplication().CreateFromPath("3.1").name
```

### Instance Methods

#### `FindFirst(query as stdIAccessible, Optional searchType as ESearchType = ESearchType.DepthFirst)`

Finds the first element which satisfies query. Query is implemented as an object which implements `stdIAccessible`. Typically `stdLambda` or `stdCallback` would be used for this. Query's signiature should be of the form `(element: stdAcc, depth: long)=>ESearchResult`

```vb
'Find where name is "hello" and class is "world":
el.FindFirst(stdLambda.Create("$1.name=""hello"" and $1.class=""world"""))

'Find first element named "hello" at depth > 4:
el.FindFirst(stdLambda.Create("$1.name = ""hello"" AND $2 > 4"))
```



#### `FindAll(query, searchType?=ESearchType.DepthFirst)`

Finds all elements which satisfy the query. Query is implemented as an object which implements `stdIAccessible`. Typically `stdLambda` or `stdCallback` would be used for this.

```vb
'Find where name is "hello" and class is "world":
el.FindFirst(stdLambda.Create("$1.name=""hello"" and $1.class=""world"""))

'Find first element named "hello" at depth > 4:
el.FindFirst(stdLambda.Create("$1.name = ""hello"" AND $2 > 4"))
```


#### Other instance methods


* `GetDescendents()` - Get all descendents of the stdAcc control






### Enums

#### `EAccSearchType`
Used while walking the Accessibility tree. Can be used to toggle between a Breadth first search (BFS) and a depth first search (DFS).

To understand the difference between BFS and DFS take this tree:
```
       A
      / \
     B   C
    /   / \
   D   E   F
```
A BFS will walk this tree in the following order: `A, B, C, D, E, F`
A DFS will walk this tree in a different order:   `A, C, F, E, B, D`

```vb
Public Enum
    BreadthFirst = 0
    DepthFirst = 1
End Enum
```

#### `ESearchResult`

Used while walking the Accessibility tree. Can be used to discard entire trees of elements, to increase speed of walk algorithms.

```vb
Public Enum EAccSearchResult
    MatchFound = 1                 'Matched                                    
    MatchFoundSearchDescendents=4  'Same as `ESearchResult.MatchFound`         
    NoMatchFound = 0               'Not found, continue searching descendents  
    NoMatchCancelSearch= 2         'Not found, cancel search                   
    NoMatchSkipDescendents= 3      'Not found, don't search descendents        
End Enum
```

#### `EAccRoles`

See [Microsoft Docs](https://docs.microsoft.com/en-us/windows/win32/winauto/object-roles) for details.

```vb
Public Enum EAccRoles
    ROLE_TITLEBAR = &H1&
    ROLE_MENUBAR = &H2&
    ROLE_SCROLLBAR = &H3&
    ROLE_GRIP = &H4&
    ROLE_SOUND = &H5&
    ROLE_CURSOR = &H6&
    ROLE_CARET = &H7&
    ROLE_ALERT = &H8&
    ROLE_WINDOW = &H9&
    ROLE_CLIENT = &HA&
    ROLE_MENUPOPUP = &HB&
    ROLE_MENUITEM = &HC&
    ROLE_TOOLTIP = &HD&
    ROLE_APPLICATION = &HE&
    ROLE_DOCUMENT = &HF&
    ROLE_PANE = &H10&
    ROLE_CHART = &H11&
    ROLE_DIALOG = &H12&
    ROLE_BORDER = &H13&
    ROLE_GROUPING = &H14&
    ROLE_SEPARATOR = &H15&
    ROLE_TOOLBAR = &H16&
    ROLE_STATUSBAR = &H17&
    ROLE_TABLE = &H18&
    ROLE_COLUMNHEADER = &H19&
    ROLE_ROWHEADER = &H1A&
    ROLE_COLUMN = &H1B&
    ROLE_ROW = &H1C&
    ROLE_CELL = &H1D&
    ROLE_LINK = &H1E&
    ROLE_HELPBALLOON = &H1F&
    ROLE_CHARACTER = &H20&
    ROLE_LIST = &H21&
    ROLE_LISTITEM = &H22&
    ROLE_OUTLINE = &H23&
    ROLE_OUTLINEITEM = &H24&
    ROLE_PAGETAB = &H25&
    ROLE_PROPERTYPAGE = &H26&
    ROLE_INDICATOR = &H27&
    ROLE_GRAPHIC = &H28&
    ROLE_STATICTEXT = &H29&
    ROLE_TEXT = &H2A&
    ROLE_PUSHBUTTON = &H2B&
    ROLE_CHECKBUTTON = &H2C&
    ROLE_RADIOBUTTON = &H2D&
    ROLE_COMBOBOX = &H2E&
    ROLE_DROPLIST = &H2F&
    ROLE_PROGRESSBAR = &H30&
    ROLE_DIAL = &H31&
    ROLE_HOTKEYFIELD = &H32&
    ROLE_SLIDER = &H33&
    ROLE_SPINBUTTON = &H34&
    ROLE_DIAGRAM = &H35&
    ROLE_ANIMATION = &H36&
    ROLE_EQUATION = &H37&
    ROLE_BUTTONDROPDOWN = &H38&
    ROLE_BUTTONMENU = &H39&
    ROLE_BUTTONDROPDOWNGRID = &H3A&
    ROLE_WHITESPACE = &H3B&
    ROLE_PAGETABLIST = &H3C&
End Enum
```

#### `EAccRoles`

```vb
Public Enum EAccStates
    STATE_NORMAL = &H0
    STATE_UNAVAILABLE = &H1
    STATE_SELECTED = &H2
    STATE_FOCUSED = &H4
    STATE_PRESSED = &H8
    STATE_CHECKED = &H10
    STATE_MIXED = &H20
    STATE_INDETERMINATE = &H99
    STATE_READONLY = &H40
    STATE_HOTTRACKED = &H80
    STATE_DEFAULT = &H100
    STATE_EXPANDED = &H200
    STATE_COLLAPSED = &H400
    STATE_BUSY = &H800
    STATE_FLOATING = &H1000
    STATE_MARQUEED = &H2000
    STATE_ANIMATED = &H4000
    STATE_INVISIBLE = &H8000
    STATE_OFFSCREEN = &H10000
    STATE_SIZEABLE = &H20000
    STATE_MOVEABLE = &H40000
    STATE_SELFVOICING = &H80000
    STATE_FOCUSABLE = &H100000
    STATE_SELECTABLE = &H200000
    STATE_LINKED = &H400000
    STATE_TRAVERSED = &H800000
    STATE_MULTISELECTABLE = &H1000000
    STATE_EXTSELECTABLE = &H2000000
    STATE_ALERT_LOW = &H4000000
    STATE_ALERT_MEDIUM = &H8000000
    STATE_ALERT_HIGH = &H10000000
    STATE_PROTECTED = &H20000000
    STATE_VALID = &H7FFFFFFF
End Enum
```

* Write Spec

* Changed all From... methods to CreateFrom... methods
* Removed weird text query system for `FindFirst` and `FindAll` and instead implemented `ICallable` system. `FindFirst` and `FindAll` will now have DFS and BFS optional parameter allowing for more optimal searching. The query callable will be able to return 4 values, each which have a means to optimise searches across the accessible tree.
* Add EMSAARoles and EMSAAStates Enums
* Added Lookups and removed ACC_STATES and ACC_ROLES objects, which now lie under lookups("roles") and lookups("states"). Note only 1 lookups object is created on the main stdAcc object. From there it is distributed. Population of lookups data improved also.
* Change GetIAccessible and SetIAccessible out for public protAccessible object. Hopefully naming convention is enough to stop people from overwriting it.
* Move to tFindNode for FindFirst and FindAll stack, and changing algorithms to compensate.
* Changing HWND as Long to LongPtr everywhere
* Renamed CreateFromExcel to CreateFromApplication
* Added StateData() function which can be used to obtain multiple states at once.
* Changed code format of Text() to a list of names and values in a JSON like format.
* Added `PrintChildTexts` and `PrintDescTexts` which are useful for debugging purposes.

* Added build once to testBuilder

* stdAcc Tests

Rudimentary stdAcc Tests

* Spelling mistake