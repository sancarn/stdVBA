Attribute VB_Name = "stdWindowTests"
'Spec:
' CONSTRUCTORS
'   [ ] Create(sClassName,sCaption,dwStyle, x, y, Width, Height, hWndParent, hMenu, hInstance, lpParam) as stdWindow
'   [ ] CreateStaticPopup(x, y, Width, Height, BorderWidth, BorderColor) as stdWindow
'   [X] CreateFromDesktop() as stdWindow
'   [X] CreateFromHwnd(hwnd) as stdWindow
'   [X] CreateFromPoint(x, y) as stdWindow
'   [X] CreateFromEvent() as stdWindow
'   [X] CreateFromIUnknown(obj) as stdWindow
' STATIC METHODS
'   [?] Requires()
' INSTANCE PROPERTIES
'   [X] Get     handle() as LongPtr
'   [X] Get     hDC() as LongPtr
'   [X] Get     Exists as Boolean
'   [X] Get     IsVisible() as Boolean
'   [X] Get     IsMinimised() as Boolean
'   [X] Get     IsMaximised() as Boolean
'   [X] Get     IsFrozen() as Boolean
'   [X] Get/Set Caption() as string
'   [X] Get     Class() as string
'   [X] Get     RectClient() as Long()
'   [X] Get/Set RectWindow() as Long()
'   [X] Get/Set X() as Long
'   [X] Get/Set Y() as Long
'   [X] Get/Set Width() as Long
'   [X] Get/Set Height() as Long
'   [X] Get     ProcessID() as long
'   [X] Get     ProcessName() as string
'   [X] Get/Set Parent() as stdWindow
'   [X] Get     AncestralRoot() as stdWindow
'   [X] Get/Set Style() as Long
'   [X] Get/Set StyleEx() as Long
'   [X] Get/Set UserData() as LongPtr
'   [X] Get/Set WndProc() as LongPtr
'   [X] Get/Set Resizable() as Boolean
'   [X] Get     Children() as stdEnumerator
'   [X] Get     Descendents(DFS?) as stdEnumerator

' INSTANCE METHODS
'   [ ] SetHook(idHook, hook, hInstance, dwThreadID) as LongPtr
'   [X] Redraw()
'   [X] SendMessage(wMsg, wParam, lParam)
'   [X] PostMessage(wMsg, wParam, lParam)
'   [ ] SendMessageTimeout(wMsg, wParam, lParam, TimeoutMilliseconds)
'   [ ] ClickInput(x?, y?, Button?)
'   [X] ClickEvent(x?, y?, Button?, isDoubleClick?, wParam?)
'   [ ] SendKeysInput(sKeys, bRaw?, keyDelay?)
'   [X] SendKeysEvent(sKeys, bRaw?, keyDelay?)
'   [X] Show()
'   [X] Hide()
'   [X] Maximize()
'   [X] Minimize()
'   [X] Activate()
' PROTECTED METHODS
'   [X] zProtGetNextDescendent(stack, DFS, Prev) as stdWindow
Sub testAll()
  Test.Topic "stdWindow"
  
  Dim wnd As stdWindow
  Set wnd = stdWindow.CreateFromDesktop()
  Test.Assert "Desktop hwnd", wnd.handle > 0
  Test.Assert "Desktop class", wnd.Class = "#32769"
  Test.Assert "Desktop children", wnd.Children.Count > 10
  Test.Assert "Desktop no parent", wnd.Parent Is Nothing
  
  
  Set wnd = stdWindow.CreateFromHwnd(Application.hwnd)
  Test.Assert "stdWindow::CreateFromHwnd()", wnd.Class = "XLMAIN"
  
End Sub


