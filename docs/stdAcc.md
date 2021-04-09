# `stdAcc`

`stdAcc` is a library built largely for Windows (and in the future Mac) window automation. The intention is to give users full flexibility over the windows on the operating system, allowing you to filter them, obtain information from them, and automate them.

## Spec

### Constructors

#### `CreateFromPoint(ByVal x As Long, ByVal y As Long) as stdAcc`

Creates an `stdAcc` object from an `X` and `Y` point location on the screen.

```vb
Debug.Print TypeName(stdAcc.CreateFromPoint(0,0))
```

#### `CreateFromHwnd(ByVal hwnd As LongPtr) as stdAcc`

Creates an `stdAcc` object from a window handle.

```vb
Debug.Print TypeName(FindWindowA(MSH_WHEELMODULE_CLASS,MSH_WHEELMODULE_TITLE))
```

#### `CreateFromApplication() as stdAcc`

Creates an `stdAcc` object from the current running application (e.g. Excel / Word / Powerpoint).

```vb
'Print name of app window
Debug.Print stdAcc.CreateFromApplication().Name
```

> Note: Implementation as current relies on `Application.hwnd`. An application agnostic method is required. As such this method is only guaranteed to work in Excel.

#### `CreateFromDesktop() as stdAcc`

Creates an `stdAcc` object from the desktop.

```vb
'Loop over all windows
Dim accWnd as stdAcc
For each accWnd in stdAcc.CreateFromDesktop().children
  Debug.print accWnd.name
next
```

> Note: Implementation as current relies `stdAcc.CreateFromApplication()`, and as such has the same limitations.

#### `CreateFromIAccessible(ByRef obj As IAccessible) as stdAcc`

Creates an `stdAcc` object from an object which implements `IAccessible`.

```vb
Dim obj As IAccessible
Dim v As Variant
Call AccessibleObjectFromPoint(x, y, obj, v)
Debug.Print CreateFromIAccessible(obj).name
```

#### `CreateFromMouse() as stdAcc`

Creates an `stdAcc` object for the element the mouse currently hovers over.

```vb
While True
    Dim obj as stdAcc
    set obj = stdAcc.CreateFromMouse()
    if not obj is nothing then Debug.Print obj.Name & " - " & obj.Role 
    DoEvents
Wend
```

#### `CreateFromPath(ByVal sPath As String) as stdAcc`

Creates an `stdAcc` object for the element at a given path from the current element.

```vb
Debug.Print stdAcc.CreateFromApplication().CreateFromPath("3.1").name
```

#### `InitWithProxy(ByRef oParent as stdAcc, ByVal index as long)`

**PROTECTED METHOD - DO NOT CALL UNLESS YOU KNOW WHAT YOU ARE DOING**

Initialises an stdAcc object as a `Proxy` object, who's methods are implemented on the parent instead of on the element itself

```vb
Dim x as new stdAcc
Call x.InitWithProxy(oParent,1)
```


----------------------------------------------------------------------------------------------

### Instance Methods

#### `GetDescendents() as Collection<stdAcc>`

Get all descendents of the stdAcc control

```vb
el.getDescendents()
```

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

#### `DoDefaultAction()`

Performs the default action of the IAccessible object

#### `SendMessage(ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long` 

**WARNING - THIS METHOD IS SCHEDULED TO BE REMOVED. Use `stdWindow#SendMessage(...)` instead**

Send a win32 message to the control


#### `PrintChildTexts()`

Prints all children texts and paths. Useful for debugging.

```vb
stdAcc.CreateFromApplication().PrintDescTexts()                      'Print to immediate window
stdAcc.CreateFromApplication().PrintDescTexts("D:\vba.log",false)    'Print to file only
```

#### `PrintDescTexts(Optional ByVal sToFilePath as string = "", Optional ByVal bPrintToDebug as boolean = true, Optional ByVal sPath As String = "P", Optional ByVal fileNum as long = 0)`

Prints all descendent texts and paths. Useful for debugging.

```vb
stdAcc.CreateFromApplication().PrintDescTexts()                      'Print to immediate window
stdAcc.CreateFromApplication().PrintDescTexts("D:\vba.log",false)    'Print to file only
```



#### `getPath(Optional toAccessible As stdAcc = Nothing) As String`

**WARNING - THERE ARE STILL A NUMBER OF KNOWN BUGS WITH THIS METHOD**

Returns the path to an element

```vb
Debug.Print el1.getPath()                'D.W.1.4.2.4

Debug.Print el1.children(2).getPath(el1) '2
```

#### `toJSON()`

Returns this element and all descendents as a JSON string. Useful for debugging

#### `highlight`

**Planned but not implemented**


#### Other instance methods


* `GetDescendents()` - Get all descendents of the stdAcc control

### Instance Properties

#### R `Parent() As stdAcc`

Return the parent of the IAccessible object

#### R `children() As Collection`

Return the children of the IAccessible object

#### R `hwnd() As LongPtr`

Return the hwnd of the IAccessible object

#### R `Location() As Collection`

Return the location of the element as a collection. Has 5 named keys: "Width", "Height", "Left", "Top" and "Parent"

```vb
With el.Location
    Debug.Print "Center: " .item("Left") + .item("Width")/2 & "," & .item("Top") + .item("Height")/2
End WIth
```

#### RW `value() As Variant`

Read or Write the value of an element.

#### R `name() As String`

Get the name of an element.

#### R `DefaultAction() As String`

Get the default action name of the element.

#### R `Role() As String`

Get the [Accessibility role](https://docs.microsoft.com/en-us/windows/win32/winauto/object-roles) of the object.

#### R `State() As String`

Get the [state](https://docs.microsoft.com/en-us/windows/win32/winauto/object-state-constants) of an element.

#### R `StateData() As Long`

Get the union of [states](https://docs.microsoft.com/en-us/windows/win32/winauto/object-state-constants) of the object. This could be several of the states OR-ed together.

#### R `Description() As String`

Gets the description of the element.

#### R `KeyboardShortcut() As String`

Gets the keyboard shortcut of the element.

#### RW `Focus() As Boolean`

Gets whether the element is focussed or not.

#### R `Help() As String`

Gets the help text of the element.

#### R `HelpTopic() As String`

Gets the help topic of the element.

#### R `Text() As String`

Gets a string which contains numerous information about the element. This can almost be seen as a descriptor for the element.

#### R `HitTest(ByVal x As Long, ByVal y As Long) As stdAcc`

Return the element under the specified location

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

| Code         | Value |
|--------------|-------|
| BreadthFirst | 0     |
| DepthFirst   | 1     |


#### `EFindResult`

Used while walking the Accessibility tree. Can be used to discard entire trees of elements, to increase speed of walk algorithms.

| Code                          | Value | Comment                                                                                                                   |      
|-------------------------------|-------|---------------------------------------------------------------------------------------------------------------------------|
| MatchFound                    |  1    | Matched                                                                                                                   |  
| MatchFoundSearchDescendents   |  4    | Same as `EFindResult.MatchFound` while using find first. In `FindAll` this will match the element and search descendents. |
| NoMatchFound                  |  0    | Not found, continue searching descendents                                                                                 |
| NoMatchCancelSearch           |  2    | Not found, cancel search                                                                                                  |  
| NoMatchSkipDescendents        |  3    | Not found, don't search descendents                                                                                       |  

#### `EAccRoles`

See [Microsoft Docs](https://docs.microsoft.com/en-us/windows/win32/winauto/object-roles) for details.

| Code                     | Value |
|--------------------------|-------|
| ROLE_TITLEBAR            | &H1&  |
| ROLE_MENUBAR             | &H2&  |
| ROLE_SCROLLBAR           | &H3&  |
| ROLE_GRIP                | &H4&  |
| ROLE_SOUND               | &H5&  |
| ROLE_CURSOR              | &H6&  |
| ROLE_CARET               | &H7&  |
| ROLE_ALERT               | &H8&  |
| ROLE_WINDOW              | &H9&  |
| ROLE_CLIENT              | &HA&  |
| ROLE_MENUPOPUP           | &HB&  |
| ROLE_MENUITEM            | &HC&  |
| ROLE_TOOLTIP             | &HD&  |
| ROLE_APPLICATION         | &HE&  |
| ROLE_DOCUMENT            | &HF&  |
| ROLE_PANE                | &H10& |
| ROLE_CHART               | &H11& |
| ROLE_DIALOG              | &H12& |
| ROLE_BORDER              | &H13& |
| ROLE_GROUPING            | &H14& |
| ROLE_SEPARATOR           | &H15& |
| ROLE_TOOLBAR             | &H16& |
| ROLE_STATUSBAR           | &H17& |
| ROLE_TABLE               | &H18& |
| ROLE_COLUMNHEADER        | &H19& |
| ROLE_ROWHEADER           | &H1A& |
| ROLE_COLUMN              | &H1B& |
| ROLE_ROW                 | &H1C& |
| ROLE_CELL                | &H1D& |
| ROLE_LINK                | &H1E& |
| ROLE_HELPBALLOON         | &H1F& |
| ROLE_CHARACTER           | &H20& |
| ROLE_LIST                | &H21& |
| ROLE_LISTITEM            | &H22& |
| ROLE_OUTLINE             | &H23& |
| ROLE_OUTLINEITEM         | &H24& |
| ROLE_PAGETAB             | &H25& |
| ROLE_PROPERTYPAGE        | &H26& |
| ROLE_INDICATOR           | &H27& |
| ROLE_GRAPHIC             | &H28& |
| ROLE_STATICTEXT          | &H29& |
| ROLE_TEXT                | &H2A& |
| ROLE_PUSHBUTTON          | &H2B& |
| ROLE_CHECKBUTTON         | &H2C& |
| ROLE_RADIOBUTTON         | &H2D& |
| ROLE_COMBOBOX            | &H2E& |
| ROLE_DROPLIST            | &H2F& |
| ROLE_PROGRESSBAR         | &H30& |
| ROLE_DIAL                | &H31& |
| ROLE_HOTKEYFIELD         | &H32& |
| ROLE_SLIDER              | &H33& |
| ROLE_SPINBUTTON          | &H34& |
| ROLE_DIAGRAM             | &H35& |
| ROLE_ANIMATION           | &H36& |
| ROLE_EQUATION            | &H37& |
| ROLE_BUTTONDROPDOWN      | &H38& |
| ROLE_BUTTONMENU          | &H39& |
| ROLE_BUTTONDROPDOWNGRID  | &H3A& |
| ROLE_WHITESPACE          | &H3B& |
| ROLE_PAGETABLIST         | &H3C& |

#### `EAccRoles`

| Code                   | Value      |
|------------------------|------------|
| STATE_NORMAL           | &H0        |
| STATE_UNAVAILABLE      | &H1        |
| STATE_SELECTED         | &H2        |
| STATE_FOCUSED          | &H4        |
| STATE_PRESSED          | &H8        |
| STATE_CHECKED          | &H10       | 
| STATE_MIXED            | &H20       | 
| STATE_INDETERMINATE    | &H99       | 
| STATE_READONLY         | &H40       | 
| STATE_HOTTRACKED       | &H80       | 
| STATE_DEFAULT          | &H100      |  
| STATE_EXPANDED         | &H200      |  
| STATE_COLLAPSED        | &H400      |  
| STATE_BUSY             | &H800      |  
| STATE_FLOATING         | &H1000     |   
| STATE_MARQUEED         | &H2000     |   
| STATE_ANIMATED         | &H4000     |   
| STATE_INVISIBLE        | &H8000     |   
| STATE_OFFSCREEN        | &H10000    |    
| STATE_SIZEABLE         | &H20000    |    
| STATE_MOVEABLE         | &H40000    |    
| STATE_SELFVOICING      | &H80000    |    
| STATE_FOCUSABLE        | &H100000   |     
| STATE_SELECTABLE       | &H200000   |     
| STATE_LINKED           | &H400000   |     
| STATE_TRAVERSED        | &H800000   |     
| STATE_MULTISELECTABLE  | &H1000000  |      
| STATE_EXTSELECTABLE    | &H2000000  |      
| STATE_ALERT_LOW        | &H4000000  |      
| STATE_ALERT_MEDIUM     | &H8000000  |      
| STATE_ALERT_HIGH       | &H10000000 |       
| STATE_PROTECTED        | &H20000000 |       
| STATE_VALID            | &H7FFFFFFF |       


### Protected Methods and Properties

#### `protGetLookups()`

**PROTECTED METHOD - DO NOT CALL UNLESS YOU KNOW WHAT YOU ARE DOING**

Returns the lookups object