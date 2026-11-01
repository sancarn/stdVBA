# `stdWindow`

## Window customization and issue #86

Owner changes (`owner`, `setOwnerHandle`), title-bar buttons
(`isMinimiseButtonVisible`, `isMaximiseButtonVisible`, `isSystemMenuVisible`),
and icons (`HICON`) were already implemented. Style setters now refresh the
non-client frame immediately. `SendMessage` and `PostMessage` preserve native
pointer widths in VBA7, including 64-bit message results and icon handles.

`owner` reads the owner rather than a child window's parent. It returns `Nothing`
for an unowned window; assigning `Nothing` removes the owner. `setOwnerHandle`
rejects child windows, whose relationship must be changed through `Parent`.

### Taskbar visibility

```vb
Dim wnd As stdWindow
Set wnd = stdWindow.CreateFromIUnknown(MyForm)
wnd.isTaskbarVisible = True
wnd.isMinimiseButtonVisible = True
wnd.isMaximiseButtonVisible = True
```

`isTaskbarVisible` configures a top-level window's taskbar eligibility using
`WS_EX_APPWINDOW` and `WS_EX_TOOLWINDOW`. It hides a visible window before changing
the styles, then shows it without activation so the Shell can refresh its button.
A hidden window stays hidden, and its owner is preserved. This getter reports
configured eligibility, including for hidden windows; it does not query Explorer's
buttons, grouping, or pinned applications. Child windows are ineligible and the
setter rejects them. Hiding a taskbar button also removes the window from Alt+Tab.

The existing `isAppWindow` and `isToolWindow` properties remain lower-level style
switches. They do not perform the Shell's hide/change/show sequence. This approach
keeps `stdWindow` independent of `stdCOM`. The taskbar COM example in `stdCOM` is
still a useful starting point for future thumbnail, overlay, or progress APIs.

### Callable window subclassing

Keep the wrapper in a module or form field for as long as it should receive messages:

```vb
'Standard module MyModule:
Private watched As stdWindow

Public Sub StartWatching(ByVal form As Object)
  Set watched = stdWindow.CreateFromIUnknown(form)
  watched.Subclass stdCallback.CreateFromModule("MyModule", "OnWindowMessage")
End Sub

Public Function OnWindowMessage(ByVal hwnd As LongPtr, ByVal message As Long, _
                                ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Variant
  Const WM_APP_PING As Long = &H8001&
  If message = WM_APP_PING Then OnWindowMessage = 123
  'Implicit Empty forwards every other message through DefSubclassProc.
End Function

Public Sub StopWatching()
  If Not watched Is Nothing Then watched.Unsubclass
  Set watched = Nothing
End Sub
```

`Subclass(callback As stdICallable, Optional enableIDEsafety As Boolean = True)` accepts `stdCallback`, `stdLambda`, or any
implementation of the interface. Each invocation receives a zero-based array of
`hwnd`, message, `wParam`, and `lParam` through `RunEx`. Return `Empty` to forward
the message, or an integer representable as `LongPtr` to handle it. Zero is a
handled result. A handled result bypasses the remaining chain. Use explicit Long
constants for messages such as `&H8001&`; VBA otherwise interprets some hexadecimal
literals as signed Integers.

Only windows belonging to the current VBA thread can be subclassed. Each wrapper
can install one callback; calling `Subclass` again replaces it. Independent
wrappers on the same HWND form a chain using `SetWindowSubclass` and
`DefSubclassProc`. `Unsubclass` removes only that wrapper and is idempotent.

`WM_NCDESTROY` detaches the wrapper before notifying its callback, then always
forwards the destruction message. The wrapper also detaches on normal object
termination. Removal, nested messages, destruction, and releasing the caller's
last wrapper reference from inside a callback are supported. Avoid reference
cycles (for example, a bound callback retaining the same window); explicitly
unsubscribe to break them.

Callback exceptions and invalid return-value conversions are contained and
forwarded. Inspect `SubclassErrorNumber` and `SubclassErrorDescription` for the
last error since installation; reinstalling clears them. Callbacks should return
quickly and avoid modal UI or `DoEvents` unless reentrancy is intended.

The thunk checks the editor mode **before touching the VBA callback pointer**:

* In Run mode it dispatches the callback normally.
* In Break mode it forwards messages without entering VBA and keeps the subclass
  installed, allowing Continue to resume callbacks. A dying window still detaches.
* When Design/reset mode (or an unknown mode) is observed, it removes the native
  subclass and forwards. `isSubclassed` reconciles that removal when queried again.

This follows the native EbMode guard pattern. A missing named export does not
make EbMode inaccessible: when `GetProcAddress` cannot find it, stdWindow locates
the internal function through the VBE Proc/MethCallEngine template, using the
discovery technique described by [The trick's CTrickTimer](https://github.com/thetrik/VbTrickTimer/blob/master/Class/CTrickTimer.cls).
The resolver checks PE headers and architecture, scans readable executable
sections with a masked full signature, and requires matching templates to agree
on one committed executable target inside the same module. Unsupported or
ambiguous runtime layouts fail installation; no guessed address is called.

The actual EbMode pointer is cached and called directly by the native thunk.
This needs no VBProject access, VBIDE reference, PDB download, or helper class.
The ordinary window/UserForm tests and resolver tests also pass in a fresh Excel
instance where AccessVBOM is disabled. The existing stdWebView helper only tries
the named export and does not currently perform this internal-function discovery.
VBAStack demonstrates another route through PDB symbols; no symbol resolver is
required by this implementation. Its separate [pure VBA stack reader](https://github.com/Ariche2/VBAStack/blob/8928cfad2d54bd1afee5627d32f9a16d012131e7/VBAStack/VBAStack.bas)
uses the Err object's runtime-thread and stack-frame structure offsets without
PDBs; those offsets locate stack data rather than the EbMode function.

The `enableIDEsafety` argument defaults to True. In the tested Excel host a
modeless starter macro returning can itself put the engine in Design mode, so
the default guard then detaches on the next message. Keep VBA running (for
example, during a modal form or an active event loop) when using this guard.
For callbacks between macros, opt out explicitly:

```vb
wnd.Subclass stdCallback.CreateFromModule("MyModule", "OnWindowMessage"), False
```

This opt-out restores ordinary modeless/background callbacks and disables the
break/reset guard. You must call `Unsubclass` before breaking, resetting, using
End, or editing running code in that mode. Changing the argument on a subsequent
Subclass call updates only that wrapper; other subclasses keep their own policy.

Explicit `Unsubclass` remains the preferred shutdown path. A mode guard reacts to
messages and cannot retroactively protect an already executing callback or promise
to detect a reset if no non-Run mode is observed before later code runs. Avoid
`End`, Reset, and code edits inside a callback. An abrupt End/reset can bypass VBA
cleanup and leave the 8 KB native guard allocation until
Office exits; ordinary removal releases them. Keep VBE Error Trapping on **Break
on Unhandled Errors** when intentionally testing handled callback errors.

The native bridge is embedded in `stdWindow`; no helper `.bas` or `stdCOM` import
is needed. It calls the typed `stdICallable.Run` vtable entry with stack-only
arguments, which forwards to the supplied callable's `RunEx`, without method-name
lookup. The user callable receives a normal VBA-owned array so it can resize or
retain its arguments. The bridge is retained across the complete COM call. x86 and x64 code is
generated from `Tools/stdWindow-subclass/thunk32.asm` and `thunk64.asm`; x64 unwind
metadata is registered with Windows. The shared code page is read/execute, with a
separate writable state page with active-call and registration counts. Normal removal frees the allocation when
idle. Removal inside the last callback defers freeing until a later removal or
normal singleton termination. Raw `SetWindowsHookEx` hooks remain unimplemented;
subclassing receives messages for a particular existing window.

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
