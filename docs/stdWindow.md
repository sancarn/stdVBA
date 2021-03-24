# `stdWindow`

## Spec

### Constructors

#### `CreateFromDesktop() as stdWindow`

Creates a `stdWindow` object from the Desktop. This object is usually used in order to enumerate over all open windows.

```vb
Dim window as stdWindow
For each window in stdWindow.CreateFromDesktop.children
    Debug.Print window.name
next
```

#### `CreateFromHwnd(ByVal hwnd as LongPtr/Long) as stdWindow`

Creates a `stdWindow` object from a supplied hwnd. It is unlikely this function will be needed unless you already have an hwnd from a different library.

```vb
Dim window as stdWindow
set window = stdWindow.CreateFromHwnd(Application.Hwnd)
```

#### `CreateFromPoint(ByVal x as long, ByVal y as Long) as stdWindow`

Creates a `stdWindow` for the window underneath a supplied point.

```vb
With MouseGetPos()
  set window = stdWindow.Create(.x, .y)
End With
```

#### `CreateFromIUnknown(ByVal obj as IUnknown) as stdWindow`

Creates a `stdWindow` object for a supplied `IUnknown` object which implements either [`IOleWindow`](https://docs.microsoft.com/en-us/windows/win32/api/oleidl/nn-oleidl-iolewindow), [`IInternetSecurityMgrSite`](https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/ms537098(v=vs.85)) or [`IShellView`](https://docs.microsoft.com/en-us/windows/win32/api/shobjidl_core/nn-shobjidl_core-ishellview). Uses shell API's `IUnknown_GetWindow` function internally. Useful for obtaining the window object of a `UserForm` object.

```vb
Dim uf as UserForm1: set uf = new UserForm1
uf.show false
With stdWindow.CreateFromIUnknown(uf)
  .resizable = true
End With
```

#### `CreateFromContextMenu() as stdWindow`

Creates a `stdWindow` object which represents the currently open windows context menu (right-click menu).

```vb
stdAcc.CreateFromHwnd(stdWindow.CreateFromContextMenu.hwnd).children(1).DoDefaultAction
```

### INSTANCE PROPERTIES

* [X] Get     handle() as LongPtr
* [X] Get     hDC() as LongPtr
* [X] Get     Exists as Boolean
* [X] Get/Let Visible() as Boolean        
* [X] Get/Let State() as EWndState    'Normal,Minimised,Maximised
* [X] Get     IsFrozen() as Boolean
* [X] Get/Let Caption() as string
* [X] Get     Class() as string
* [X] Get     RectClient() as Long()
* [X] Get/Let RectWindow() as Long()
* [X] Get/Let X() as Long
* [X] Get/Let Y() as Long
* [X] Get/Let Width() as Long
* [X] Get/Let Height() as Long
* [X] Get     ProcessID() as long
* [X] Get     ProcessName() as string
* [X] Get/Set Parent() as stdWindow
* [X] Get     AncestralRoot() as stdWindow
* [X] Get/Let Style() as Long
* [X] Get/Let StyleEx() as Long
* [X] Get/Let UserData() as LongPtr
* [X] Get/Let WndProc() as LongPtr
* [X] Get/Let Resizable() as Boolean
* [X] Get     Children() as Collection
* [X] Get     Descendents() as Collection
* [ ] Get/Let AlwaysOnTop() as Boolean

### INSTANCE METHODS

* [ ] SetHook(idHook, hook, hInstance, dwThreadID) as LongPtr
* [X] Redraw()
* [X] SendMessage(wMsg, wParam, lParam)
* [X] PostMessage(wMsg, wParam, lParam)
* [ ] SendMessageTimeout(wMsg, wParam, lParam, TimeoutMilliseconds)
* [ ] ClickInput(x?, y?, Button?)
* [X] ClickEvent(x?, y?, Button?, isDoubleClick?, wParam?)
* [ ] SendKeysInput(sKeys, bRaw?, keyDelay?)
* [X] SendKeysEvent(sKeys, bRaw?, keyDelay?)
* [X] Show()
* [X] Hide()
* [X] Maximize()
* [X] Minimize()
* [X] Activate()
* [X] Close()
* [X] FindFirst(query)
* [X] FindAll(query)
* [ ] Screenshot()

### PROTECTED METHODS

* [X] protInit(Byval hwnd as Longptr/Long)
* [X] protGetNextDescendent(stack, DFS, Prev) as stdWindow
* [X] protGetLookups()

## stdVBA Developer docs

### Multi-platform

This feature is unimplemented and thus this section will be mostly developer docs

In an ideal world this library would work on all operating systems. The library is definitely possible to implement on all operating systems however currently I mostly lack the time to build and test it on all OSes.

The implementation details of implementing this library on Mac OS X involve use of either ObjC directly (via Obj C runtime C API), or use of a JXA daemon which accesses the `System Events` application. For example, looping through all windows (via Desktop window) can be done with:

```js
Application("System Events").processes().filter(e=>e.windows().length > 0)
```

which returns

```js
[
  Application("System Events").applicationProcesses.byName("Terminal"),
  Application("System Events").applicationProcesses.byName("Microsoft Excel"),
  Application("System Events").applicationProcesses.byName("TextEdit"), 
  Application("System Events").applicationProcesses.byName("GitHub Desktop"), 
  Application("System Events").applicationProcesses.byName("Activity Monitor"),
  Application("System Events").applicationProcesses.byName("Google Chrome")
]
```

There might have to be some abstractions cross platform and some methods may be system specific. however