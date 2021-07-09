VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
'======================================
'TODO:    Submit merge request
'======================================

'@class stdWindow
'@description A class for managing windows
'@example:
'   With stdWindow.CreateFromDesktop()
'     Dim notepad as stdWindow
'     set notepad = .Find(stdLambda.Create("$1.Caption = ""Untitled - Notepad"" and $1.ProcessName = ""notepad.exe"""))
'     nodepad.SendKeysInput("hello world")
'     nodepad.SendKeysInput("^a")
'     nodepad.SendKeysInput("^c")
'     Debug.Print stdClipboard.Text
'   End With
'
'   'Make a userform resizable
'   MyForm.show
'   stdWindow.CreateFromIUnknown(MyForm).resizable = true
'
'Spec:
' CONSTRUCTORS
'   [ ] Create(sClassName,sCaption,dwStyle, x, y, Width, Height, hWndParent, hMenu, hInstance, lpParam) as stdWindow
'   [ ] TODO:CreateStaticPopup(x, y, Width, Height, BorderWidth, BorderColor) as stdWindow
'   [X] CreateFromDesktop() as stdWindow
'   [X] CreateFromHwnd(hwnd) as stdWindow
'   [X] CreateFromPoint(x, y) as stdWindow
'   [ ] CreateFromEvent() as stdWindow
'   [X] CreateFromIUnknown(obj) as stdWindow
'   [X] CreateFromContextMenu() as stdWindow    'Class == "#32768"
' STATIC METHODS
'   [?] Requires()
' INSTANCE PROPERTIES
'   [X] Get     handle() as LongPtr
'   [X] Get     hDC() as LongPtr
'   [X] Get     Exists as Boolean
'   [X] Get/Let Visible() as Boolean        
'   [X] Get/Let State() as EWndState    'Normal,Minimised,Maximised
'   [X] Get     IsFrozen() as Boolean
'   [X] Get/Let Caption() as string
'   [X] Get     Class() as string
'   [X] Get     RectClient() as Long()
'   [X] Get/Let RectWindow() as Long()
'   [X] Get/Let X() as Long
'   [X] Get/Let Y() as Long
'   [X] Get/Let Width() as Long
'   [X] Get/Let Height() as Long
'   [X] Get     ProcessID() as long
'   [X] Get     ProcessName() as string
'   [X] Get/Set Parent() as stdWindow
'   [X] Get     AncestralRoot() as stdWindow
'   [X] Get/Let Style() as Long
'   [X] Get/Let StyleEx() as Long
'   [X] Get/Let UserData() as LongPtr
'   [X] Get/Let WndProc() as LongPtr
'   [X] Get/Let Resizable() as Boolean
'   [X] Get     Children() as Collection
'   [X] Get     Descendents() as Collection
'   [ ] Get/Let AlwaysOnTop() as Boolean
'
' INSTANCE METHODS
'   [ ] SetHook(idHook, hook, hInstance, dwThreadID) as LongPtr
'   [X] Redraw()
'   [X] SendMessage(wMsg, wParam, lParam)
'   [X] PostMessage(wMsg, wParam, lParam)
'   [ ] TODO: SendMessageTimeout(wMsg, wParam, lParam, TimeoutMilliseconds)
'   [ ] ClickInput(x?, y?, Button?)
'   [X] ClickEvent(x?, y?, Button?, isDoubleClick?, wParam?)
'   [ ] SendKeysInput(sKeys, bRaw?, keyDelay?)
'   [X] SendKeysEvent(sKeys, bRaw?, keyDelay?)
'   [X] Activate()
'   [X] Close()
'   [X] FindFirst(query)
'   [X] FindAll(query)
'   [ ] Screenshot()
' PROTECTED METHODS
'   [X] zProtGetNextDescendent(stack, DFS, Prev) as stdWindow
' GENERIC
'   [ ] TODO: Mac compatibility







'--------------------------------------------------------------------------------
'Win API Declares
'--------------------------------------------------------------------------------
Private Type apiRect
  left As Long
  top As Long
  right As Long
  bottom As Long
End Type

Private Type apiWindowInfo
  cbSize As Integer         'DWORD
  rcWindow As apiRect       'RECT
  rcClient As apiRect       'RECT
  dwStyle As Integer        'DWORD
  dwExStyle As Integer      'DWORD
  dwWindowStatus As Integer 'DWORD
  cxWindowBorders As Long   'UINT
  cyWindowBorders As Long   'UINT
  atomWindowType As Long    'ATOM
  wCreatorVersion As Long   'WORD
End Type
Public Enum apiWindowHookType
  WH_MSGFILTER = -1
  WH_JOURNALRECORD = 0
  WH_JOURNALPLAYBACK = 1
  WH_KEYBOARD = 2
  WH_GETMESSAGE = 3
  WH_CALLWNDPROC = 4
  WH_SYSMSGFILTER = 6
  WH_MOUSE = 7
  WH_SHELL = 10
  WH_CALLWNDPROCRET = 12
  WH_KEYBOARD_LL = 13
  WH_MOUSE_LL = 14
  WH_CBT = 5
  WH_DEBUG = 9
  WH_FOREGROUNDIDLE = 11
End Enum

'https://www.autohotkey.com/docs_1.0/misc/Styles.htm
Public Enum EWndStyles
  WS_BORDER         = &H800000
  WS_CAPTION        = &HC00000
  WS_CHILD          = &H40000000
  WS_CHILDWINDOW    = &H40000000
  WS_CLIPCHILDREN   = &H2000000
  WS_CLIPSIBLINGS   = &H4000000
  WS_DISABLED       = &H8000000
  WS_DLGFRAME       = &H400000
  WS_GROUP          = &H20000
  WS_HSCROLL        = &H100000
  WS_ICONIC         = &H20000000
  WS_MAXIMIZE       = &H1000000
  WS_MAXIMIZEBOX    = &H10000
  WS_MINIMIZE       = &H20000000
  WS_MINIMIZEBOX    = &H20000
  WS_OVERLAPPED     = &H0
  WS_POPUP          = &H80000000
  WS_SIZEBOX        = &H40000
  WS_SYSMENU        = &H80000
  WS_TABSTOP        = &H10000
  WS_THICKFRAME     = &H40000
  WS_TILED          = &H0
  WS_VISIBLE        = &H10000000
  WS_VSCROLL        = &H200000

  WS_OVERLAPPEDWINDOW = WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX
  WS_POPUPWINDOW = WS_POPUP Or WS_BORDER Or WS_SYSMENU
End Enum
Public Enum EWndExStyles
  WS_EX_ACCEPTFILES          = &H00000010 
  WS_EX_APPWINDOW            = &H00040000 
  WS_EX_CLIENTEDGE           = &H00000200 
  WS_EX_COMPOSITED           = &H02000000 
  WS_EX_CONTEXTHELP          = &H00000400 
  WS_EX_CONTROLPARENT        = &H00010000 
  WS_EX_DLGMODALFRAME        = &H00000001 
  WS_EX_LAYERED              = &H00080000 
  WS_EX_LAYOUTRTL            = &H00400000 
  WS_EX_LEFT                 = &H00000000 
  WS_EX_LEFTSCROLLBAR        = &H00004000 
  WS_EX_LTRREADING           = &H00000000 
  WS_EX_MDICHILD             = &H00000040 
  WS_EX_NOACTIVATE           = &H08000000 
  WS_EX_NOINHERITLAYOUT      = &H00100000 
  WS_EX_NOPARENTNOTIFY       = &H00000004 
  WS_EX_NOREDIRECTIONBITMAP  = &H00200000 
  WS_EX_RIGHT                = &H00001000 
  WS_EX_RIGHTSCROLLBAR       = &H00000000 
  WS_EX_RTLREADING           = &H00002000 
  WS_EX_STATICEDGE           = &H00020000 
  WS_EX_TOOLWINDOW           = &H00000080 
  WS_EX_TOPMOST              = &H00000008 
  WS_EX_TRANSPARENT          = &H00000020 
  WS_EX_WINDOWEDGE           = &H00000100 
  WS_EX_OVERLAPPEDWINDOW     = WS_EX_WINDOWEDGE or WS_EX_CLIENTEDGE 
  WS_EX_PALETTEWINDOW        = WS_EX_WINDOWEDGE or WS_EX_TOOLWINDOW or WS_EX_TOPMOST
End Enum
Public Enum EWndState
  Normal
  Maximised
  Minimised
End Enum
Public Enum EWndRectType
  RectTypeWindow
  RectTypeClient
End Enum

'SendInput() API helpers
'TODO: Not used yet
'========================================
Private Enum apiTagInputType
  INPUT_MOUSE = 0
  INPUT_KEYBOARD = 1
  INPUT_HARDWARE = 2
End Enum
Private Type apiKeyboardInput
  wVirtualKey As Integer
  wScanCode As Integer
  dwFlags As Long
  time As Long
  #if VBA7 then
    dwExtraInfo As LongPtr
  #else
    dwExtraInfo As Long
  #end if
End Type
Private Type apiTagInput
  iType As apiTagInputType
  kbd As apiKeyboardInput
End Type
Private Enum KeyState
  Up
  Down
End Enum
Private Enum KeyModifier
  Ctrl = 1
  Shift = 2
  Alt = 4
  Meta = 8
End Enum
Private Type KeyToken
  wVirtualKey As Integer  'https://docs.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes
  wScanCode As Integer    '
  iKeyState As KeyState
  iModifiers() As KeyModifier
End Type
'========================================


'Or use EnumChildWindows
Private Enum apiWindowRelationship
  GW_CHILD = 5
  GW_ENABLEDPOPUP = 6
  GW_HWNDFIRST = 0
  GW_HWNDLAST = 1
  GW_HWNDNEXT = 2
  GW_HWNDPREV = 3
  GW_OWNER = 4
End Enum

Private Enum EWndShowStyle
  SW_HIDE              = 0   
  SW_SHOWNORMAL        = 1   'Shows/Restores + Activates
  SW_SHOWMINIMIZED     = 2   'Activates the window and displays it as a minimized window.
  SW_MAXIMIZE          = 3   'Maximize
  SW_SHOWNOACTIVATE    = 4   'Shows in most recent size + position but doesn't activate
  SW_SHOW              = 5   'Activate
  SW_MINIMIZE          = 6   'Minimize
  SW_SHOWMINNOACTIVE   = 7   'Minimize no activate
  SW_SHOWNA            = 8   'Show in current size and position, no activate
  SW_RESTORE           = 9   'Restore
  SW_SHOWDEFAULT       = 10  'Default window state at start of program
  SW_FORCEMINIMIZE     = 11  
End Enum

Private Enum apiWindowLongType
  GWL_WNDPROC = -4
  GWL_HINSTANCE = -6
  GWL_HWNDPARENT = -8
  GWL_ID = -12
  GWL_STYLE = -16
  GWL_EXSTYLE = -20
  GWL_USERDATA = -21
  
  'If HWND is a dialog box
  DWL_MSGRESULT = 0
  'DWL_DLGPROC = DWLP_MSGRESULT + sizeof(LRESULT)
  'DWL_USER = DWL_DLGPROC + sizeof(DLGPROC)
End Enum

Private Enum apiWindowAncestorType
  GA_PARENT = 1
  GA_ROOT = 2
  GA_ROOTOWNER = 3
End Enum

Public Enum EWndMouseButton
  LButton
  RButton
  MButton
End Enum

'Used while walking the Window tree. Can be used to toggle between a Breadth first search and a depth first search.
Public Enum EWndFindType
  BreadthFirst = 0
  DepthFirst = 1
End Enum

'Used while walking the Window tree. Can be used to discard entire trees of elements, to increase speed of walk algorithms.
Public Enum EWndFindResult
  MatchFound = 1                   'Matched                                    
  MatchFoundSearchDescendents = 4  'Same as `ESearchResult.MatchFound`         
  NoMatchFound = 0                 'Not found, continue searching descendents  
  NoMatchCancelSearch = 2          'Not found, cancel search                   
  NoMatchSkipDescendents = 3       'Not found, don't search descendents        
End Enum

Private Type tFindNode
  initialised as boolean
  depth as long
  element as object
End Type

#if VBA7 then
  'Constructors
  Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
  Private Declare PtrSafe Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByVal hwnd As LongPtr) As Long
  Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal x As Long, ByVal y As Long) As LongPtr

  'Getting window data
  Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
  Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As LongPtr) As Long
  Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hwnd As LongPtr) As Long
  Private Declare PtrSafe Function IsHungAppWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
  Private Declare PtrSafe Function IsZoomed Lib "user32" (ByVal hwnd As LongPtr) As Long
  Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
  Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
  Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hwnd As LongPtr, ByRef pRect As apiRect) As Long
  Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, ByRef pRect As apiRect) As Long
  Private Declare PtrSafe Function GetWindowInfo Lib "user32" (ByVal hwnd As LongPtr, ByRef pInf As apiWindowInfo) As Long
  Private Declare PtrSafe Function GetParent Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
  Private Declare PtrSafe Function SetParent Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndParent As LongPtr) As LongPtr
  Private Declare PtrSafe Function GetWindowDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
  Private Declare PtrSafe Function GetWindowLongA Lib "user32" (ByVal hwnd As LongPtr, ByVal nIndex As apiWindowLongType) As Long
  Private Declare PtrSafe Function GetWindowLongPtrA Lib "user32" (ByVal hwnd As LongPtr, ByVal nIndex As apiWindowLongType) As LongPtr
  Private Declare PtrSafe Function SetWindowLongA Lib "user32" (ByVal hwnd As LongPtr, ByVal nIndex As apiWindowLongType, ByVal dwNewLong As Long) As Long
  Private Declare PtrSafe Function SetWindowLongPtrA Lib "user32" (ByVal hwnd As LongPtr, ByVal nIndex As apiWindowLongType, ByVal dwNewPtr As LongPtr) As Long
  Private Declare PtrSafe Function GetAncestor Lib "user32" (ByVal hwnd As LongPtr, ByVal nIndex As apiWindowAncestorType) As LongPtr
  Private Declare PtrSafe Function FindWindowExA Lib "user32" (ByVal hwnd as LongPtr, ByVal hwndChildAfter as LongPtr, ByVal lpszClass as string, ByVal lpszWindow as string) as LongPtr
  Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hwndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal flags As Long) As Long
  Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal bRepaint As Boolean) As Long

  'Redrawing window, UpdateWindow can also be used but isn't as safe...
  Private Declare PtrSafe Function RedrawWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal lprcUpdate As LongPtr, ByVal hrgnUpdate As LongPtr, ByVal flags As Long) As Long

  'Get children / siblings / parent
  Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal wCmd As apiWindowRelationship) As LongPtr

  'Get process related data
  Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As LongPtr, ByRef ldpwProcessId As Long) As Long
  Private Declare PtrSafe Function GetWindowModuleFileName Lib "user32" Alias "GetWindowModuleFileNameA" (ByVal hwnd As LongPtr, ByVal pszFileName As String, ByVal cchFileNameMax As Long) As Long

  'Setting window data
  Private Declare PtrSafe Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String) As Long

  'Automating windows
  Private Declare PtrSafe Function apiSendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare PtrSafe Function apiPostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As EWndShowStyle) As Long
  Private Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hwnd As LongPtr) As Long
  Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hwnd As LongPtr) As Long

  'SendKeys
  Private Declare PtrSafe Function GetMessageExtraInfo Lib "user32" () As LongPtr
#else
  'Constructors
  Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As Long
  Private Declare PtrSafe Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByVal hwnd As Long) As Long
  Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

  'Getting window data
  Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare PtrSafe Function IsHungAppWindow Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare PtrSafe Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
  Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
  Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hwnd As Long, ByRef pRect As apiRect) As Long
  Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As Long, ByRef pRect As apiRect) As Long
  Private Declare PtrSafe Function GetWindowInfo Lib "user32" (ByVal hwnd As Long, ByRef pInf As apiWindowInfo) As Long
  Private Declare PtrSafe Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare PtrSafe Function SetParent Lib "user32" (ByVal hwnd As Long, ByVal hWndParent As Long) As Long
  Private Declare PtrSafe Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare PtrSafe Function GetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As apiWindowLongType) As Long
  Private Declare PtrSafe Function GetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As apiWindowLongType) As Long
  Private Declare PtrSafe Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As apiWindowLongType, ByVal dwNewLong As Long) As Long
  Private Declare PtrSafe Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As apiWindowLongType, ByVal dwNewPtr As Long) As Long
  Private Declare PtrSafe Function GetAncestor Lib "user32" (ByVal hwnd As Long, ByVal nIndex As apiWindowAncestorType) As Long
  Private Declare PtrSafe Function FindWindowExA Lib "user32" (ByVal hwnd as Long, ByVal hwndChildAfter as Long, ByVal lpszClass as string, ByVal lpszWindow as string) as Long
  Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hwndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal flags As Long) As Long
  Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal bRepaint As Boolean) As Long

  'Redrawing window, UpdateWindow can also be used but isn't as safe...
  Private Declare PtrSafe Function RedrawWindow Lib "user32" (ByVal hwnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal flags As Long) As Long

  'Get children / siblings / parent
  Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As apiWindowRelationship) As Long

  'Get process related data
  Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, ByRef ldpwProcessId As Long) As Long
  Private Declare PtrSafe Function GetWindowModuleFileName Lib "user32" Alias "GetWindowModuleFileNameA" (ByVal hwnd As Long, ByVal pszFileName As String, ByVal cchFileNameMax As Long) As Long

  'Setting window data
  Private Declare PtrSafe Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long

  'Automating windows
  Private Declare PtrSafe Function apiSendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare PtrSafe Function apiPostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As EWndShowStyle) As Long
  Private Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

  'SendKeys
  Private Declare PtrSafe Function GetMessageExtraInfo Lib "user32" () As Long
#end if


#if VBA7 then
  Private pHandle As LongPtr
#else
  Private pHandle As Long
#end if

Private pInitialized As Boolean
Private Lookups as Object


'================================================================================================
'=      PUBLIC CONSTRUCTORS
'================================================================================================
'TODO: Create a window object from information passed in by this function
'@constructor
'@param {ByVal String} The class name can be any name registered with RegisterClass or RegisterClassEx, provided that the module that registers the class is also the module that creates the window. The class name can also be any of the predefined system class names. For a list of system class names, see the Remarks section.
'@param {ByVal String} The name/caption of the window
'@param {ByVal Long (DWORD)} The window style for the window
'@param {ByVal Long} The x coordinate of the window
'@param {ByVal Long} The y coordinate of the window
'@param {ByVal Long} The width of the window
'@param {ByVal Long} The height of the window
'@param {ByVal LongPtr} Parent window handle. Can be 0 for pop-up windows.
'@param {ByVal LongPtr} Menu handle. Can be 0 for pop-up windows.
'@param {ByVal LongPtr} Module Instance handle.
'@param {ByVal lpParam} Pointer to a location where extra information is stored. Or ClientCreateStruct (for MDI windows), or null if no extra data required
'@returns {stdWindow} The created window
'@remarks System Class Names: BUTTON, COMBOBOX, EDIT, LISTBOX, MDICLIENT, RICHEDIT, RICHEDIT_CLASS, SCROLLBAR, STATIC
#if VBA7 then
Private Function Create(ByVal sClassName As String, ByVal sCaption As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, ByVal lpParam As Long) As stdWindow
#else
Private Function Create(ByVal sClassName As String, ByVal sCaption As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByVal lpParam As Long) As stdWindow
#end if
  Err.Raise 1, "Create", "Not Implemented"
End Function

'TODO: Create a window object used mainly for highlighting areas
'@constructor
'@param {ByVal Long} The x coordinate of the window
'@param {ByVal Long} The y coordinate of the window
'@param {ByVal Long} The width of the window
'@param {ByVal Long} The height of the window
'@param {ByVal Long} The width of the colored border
'@param {ByVal Long} The color of the colored border
'@returns {stdWindow} The created highlighting box
'@remarks https://stackoverflow.com/questions/3970066/creating-a-transparent-window-in-c-win32
Private Function CreateStaticPopup(ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal BorderWidth As Long, ByVal BorderColor As Long) As stdWindow
  Err.Raise 1, "Create", "Not Implemented"
End Function

'Create a window from the desktop window
'@constructor
'@returns {stdWindow} Desktop window
Public Function CreateFromDesktop() As stdWindow
  Set CreateFromDesktop = New stdWindow
  Call CreateFromDesktop.protInit(GetDesktopWindow())
End Function

'Create a window object from a window handle
'@constructor
'@param {ByVal LongPtr} Handle to window (hwnd) to create window object for
'@returns {stdWindow} Specificed window.
#if VBA7 then
Public Function CreateFromHwnd(ByVal hwnd As LongPtr) As stdWindow
#else
Public Function CreateFromHwnd(ByVal hwnd As Long) As stdWindow
#end if
  Set CreateFromHwnd = New stdWindow
  Call CreateFromHwnd.protInit(hwnd)
End Function


'Find and Create a window object for a window displayed intersecting a point on the screen.
'@constructor
'@param {ByVal Long} X of Point to find window at
'@param {ByVal Long} Y of Point to find window at
'@returns {stdWindow} Window intersecting point.
Public Function CreateFromPoint(ByVal x As Long, ByVal y As Long) As stdWindow
  Set CreateFromPoint = New stdWindow
  Call CreateFromPoint.protInit(WindowFromPoint(x, y))
End Function

''Create from Shell event
'Public Function CreateFromEvent() as stdWindow
'
'End Function

'Create a window object from an object which implements either IOleWindow, IInternetSecurityMgrSite or IShellView. Uses shell API's IUnknown_GetWindow internally.
'@constructor
'@param {ByVal IUnknown} Object which implements GetWindow() method
'@returns {stdWindow} Window specified by object
Public Function CreateFromIUnknown(ByVal obj As IUnknown) As stdWindow
  #if VBA7 then
    Dim hwnd As LongPtr
  #else
    Dim hwnd As Long
  #end if
  Dim hResult As Long
  hResult = IUnknown_GetWindow(obj, VarPtr(hwnd))
  If hResult = 0 Then
    Set CreateFromIUnknown = New stdWindow
    Call CreateFromIUnknown.protInit(hwnd)
  ElseIf hResult = -2147467262 Then
    Err.Raise 1, "CreateFromIUnknown", "This object does not implement IOleWindow, IInternetSecurityMgrSite or IShellView, and thus cannot retrieve the window assosciated with the object."
  Else
    Err.Raise 1, "CreateFromIUnknown", "An unknown error has occurred.", hResult
  End If
End Function

'Create a window object from the context menu. There should only ever be 1 context menu on the system at one time.
'@constructor
'@returns {stdWindow} Window object representing the ContextMenu
Public Function CreateFromContextMenu() as stdWindow
  #if VBA7 then
    Dim hwnd as LongPtr
  #else
    Dim hwnd as Long
  #end if
  hwnd = FindWindowExA(null, null, MakeIntAtom(&H8000&), vbNullString)
  if hwnd <> 0 then 
    set CreateFromContextMenu = CreateFromHwnd(hwnd)
  else
    set CreateFromContextMenu = nothing
  end if
End Function



'Notes:
'Windows are hierarchical therefore CreateManyFromQuery and CreateFromQuery makes less sense than FindFirst() and FindAll() methods

'================================================================================================
'=      PROTECTED CONSTRUCTORS / DESTRUCTORS
'================================================================================================
'Initialize a window object
'@constructor
'@protected
'@param {ByVal LongPtr} Handle to window (hwnd) to create window object for
'@returns {stdWindow} Specificed window.
#if VBA7 then
Friend Sub protInit(ByVal hwnd As LongPtr)
#else
Friend Sub protInit(ByVal hwnd As Long)
#end if
  pHandle = hwnd
  pInitialized = True
End Sub

'Whenever we initialise ensure to get lookups from `stdWindow`
Private Sub Class_Initialize()
    set Lookups = stdWindow.protGetLookups()
End Sub


'--------------------------------------------------------------------------------
'Access window information
'--------------------------------------------------------------------------------

'Get the hWND / window ID of the window
#if VBA7 then
Public Property Get handle() As LongPtr
#else
Public Property Get handle() As Long
#end if
  handle = pHandle
End Property

'Get the handle to the display context for the window
#if VBA7 then
Public Property Get hDC() As LongPtr
#else
Public Property Get hDC() As Long
#end if
  hDC = GetWindowDC(pHandle)
End Property

'Detect if the window exists
Public Property Get Exists() As Boolean
  Exists = IsWindow(pHandle)
End Property

'Detect if the window is hanging/frozen
Public Property Get IsFrozen() As Boolean
  If Exists Then
    IsFrozen = IsHungAppWindow(pHandle)
  Else
    Err.Raise 1, "IsFrozen", "Window does not exist."
  End If
End Property

'Get/Set the window caption
Public Property Get Caption() As String
  If Exists Then
    Dim sCharBuffer As String, iNumChars As Long
    sCharBuffer = Space(256)
    iNumChars = GetWindowText(pHandle, sCharBuffer, 256)
    Caption = Mid(sCharBuffer, 1, iNumChars)
  Else
    Err.Raise 1, "Caption", "Window does not exist."
  End If
End Property
Public Property Let Caption(ByVal s As String)
  If Exists Then
    If Not CBool(SetWindowText(pHandle, s)) Then
      Err.Raise 1, "Caption [Let]", "Window text could not be set."
    End If
  Else
    Err.Raise 1, "Caption [Let]", "Window does not exist."
  End If
End Property

'Get the window class
Public Property Get Class() As String
  If Exists Then
    Dim sCharBuffer As String, iNumChars As Long
    sCharBuffer = Space(256)
    iNumChars = GetClassName(pHandle, sCharBuffer, 256)
    Class = Mid(sCharBuffer, 1, iNumChars)
  Else
    Err.Raise 1, "Class", "Window does not exist."
  End If
End Property

'Get/Let visibility of window
Public Property Get Visible() as Boolean
  if Exists then
    Visible = IsWindowVisible(pHandle)
  else
    Err.Raise 1, "Visible", "Window does not exist."
  end if
End Property
Public Property Let Visible(ByVal x as Boolean)
  if Exists then
    if x then
      Call ShowWindow(pHandle, EWndShowStyle.SW_SHOW)
    else
      Call ShowWindow(pHandle, EWndShowStyle.SW_HIDE)
    end if
  else
    Err.Raise 1, "Visible", "Window does not exist."
  end if
End Property

'Get/Let windowState of window
Public Property Get State() as EWndState
  if Exists then
    if isZoomed(pHandle) then
      State = EWndState.Maximised
    elseif IsIconic(pHandle) then
      State = EWndState.Minimised
    else
      State = EWndState.Normal
    end if
  else
    Err.Raise 1, "State", "Window does not exist."
  end if  
End Property
Public Property Let State(ByVal x as EWndState)
  if Exists then
    if Visible then
      select case x
        case EWndState.Normal
          Call ShowWindow(pHandle, EWndShowStyle.SW_RESTORE)
        case EWndState.Maximised
          Call ShowWindow(pHandle, EWndShowStyle.SW_MAXIMIZE)
        case EWndState.Minimised
          Call ShowWindow(pHandle, EWndShowStyle.SW_FORCEMINIMIZE)
      end select
    else
      Err.Raise 1, "State", "Cannot set window state of a hidden window."
    end if
  else
    Err.Raise 1, "State", "Window does not exist."
  end if
End Property

'Get the child windows of this window
Public Property Get Children() As Collection
  If Exists Then
    'Define collection to return
    Dim ret As Collection
    Set ret = New Collection

    'Attempt to get a child window
    #if VBA7 then
      Dim childHandle As LongPtr
    #else
      Dim childHandle As Long
    #end if
    childHandle = GetWindow(pHandle, GW_CHILD)

    'If a child window exists, find all children
    if  childHandle <> 0 then
      'Quickly add all hwnds to an array
      'This is a fast operation, object creation is anticipated to be a slower operation, thus is left till afterwards
      #if VBA7 then
        Dim childHandles() As LongPtr
      #else
        Dim childHandles() As Long
      #end if

      Dim i As Long: i = -1
      Do While (childHandle <> 0)
          i = i + 1
          ReDim Preserve childHandles(i)
          childHandles(i) = childHandle
          childHandle = GetWindow(childHandle, GW_HWNDNEXT)
      Loop
      
      'Create stdWindow objects from hwnds array and add them to collection.
      For i = LBound(childHandles) To UBound(childHandles)
        Call ret.Add(stdWindow.CreateFromHwnd(childHandles(i)))
      Next
    end if

    'Return child objects
    Set Children = ret
  Else
    Err.Raise 1, "Children", "Window does not exist."
  End If
End Property

'Get the Client rect - I.E. Position and Size of Window's Client area
Public Property Get RectClient() As Double()
  If Exists Then
    Dim rect As apiRect
    If Not GetClientRect(pHandle, rect) Then
      Dim fRet(0 To 3) As Double
      With rect
        fRet(0) = .left
        fRet(1) = .top
        fRet(2) = .right - .left
        fRet(3) = .bottom - .top
        RectClient = fRet
      End With
    Else
      Err.Raise 1, "RectClient", "Cannot get client rect.", Err.LastDllError
    End If
  Else
    Err.Raise 1, "RectClient", "Window does not exist."
  End If
End Property

'Get/Set the Window rect - I.E. Position and Size of window
Public Property Get RectWindow() As Variant
  If Exists Then
    Dim rect As apiRect
    If Not GetWindowRect(pHandle, rect) Then
      Dim fRet As Variant
      ReDim fRet(0 To 3)
      With rect
        fRet(0) = .left
        fRet(1) = .top
        fRet(2) = .right - .left
        fRet(3) = .bottom - .top
        RectWindow = fRet
      End With
    Else
      Err.Raise 1, "RectWindow", "Cannot get window rect.", Err.LastDllError
    End If
  Else
    Err.Raise 1, "RectWindow", "Window does not exist."
  End If
End Property
Public Property Let RectWindow(ByVal rect As Variant)
  If IsArray(rect) Then
    If Exists Then
      Call MoveWindow(pHandle, rect(0), rect(1), rect(2), rect(3), True)
    Else
      Err.Raise 1, "RectWindow [Let]", "Window does not exist."
    End If
  Else
    Err.Raise 1, "RectWindow [Let]", "Value must be an array of 4 doubles - left, top, width, height."
  End If
End Property
Public Property Get RectByType(optional iClWnd as EWndRectType) as Variant
  select case iClWnd
    case EWndRectType.RectTypeClient
      RectByType = RectClient
    case EWndRectType.RectTypeWindow
      RectByType = RectWindow
  end select
End Property
Public Property Let RectByType(optional iClWnd as EWndRectType, ByVal rect As Variant)
  select case iClWnd
    case EWndRectType.RectTypeClient
      RectClient = Rect
    case EWndRectType.RectTypeWindow
      RectWindow = Rect
  end select
End Property

'Get/Set the X position of this window
Public Property Get x(optional ByVal iClWnd as EWndRectType) As Long
  If Exists Then
    x = Me.RectByType(iClWnd)(0)
  Else
    Err.Raise 1, "stdWindow#x", "Window does not exist."
  End If
End Property
Public Property Let x(optional ByVal iClWnd as EWndRectType, ByVal vX As Long)
  If Exists Then
    Dim rect As Variant: rect = RectByType(iClWnd)
    Call MoveWindow(pHandle, vX, rect(1), rect(2), rect(3), True)
  Else
    Err.Raise 1, "stdWindow#x [Let]", "Window does not exist."
  End If
End Property

'Get/Set the Y position of this window
Public Property Get y(optional ByVal iClWnd as EWndRectType) As Long
  If Exists Then
    y = RectByType(iClWnd)(1)
  Else
    Err.Raise 1, "stdWindow#y", "Window does not exist."
  End If
End Property
Public Property Let y(optional ByVal iClWnd as EWndRectType, ByVal vY As Long)
  If Exists Then
    Dim rect As Variant: rect = RectByType(iClWnd)
    Call MoveWindow(pHandle, rect(0), vY, rect(2), rect(3), True)
  Else
    Err.Raise 1, "stdWindow#y [Let]", "Window does not exist."
  End If
End Property

'Get/Set the width of this window
Public Property Get width(optional ByVal iClWnd as EWndRectType) As Long
  If Exists Then
    width = RectByType(iClWnd)(2)
  Else
    Err.Raise 1, "stdWindow#width", "Window does not exist."
  End If
End Property
Public Property Let width(optional ByVal iClWnd as EWndRectType, ByVal vW As Long)
  If Exists Then
    Dim rect As variant:: rect = RectByType(iClWnd)
    Call MoveWindow(pHandle, rect(0), rect(1), vW, rect(3), True)
  Else
    Err.Raise 1, "stdWindow#width [Let]", "Window does not exist."
  End If
End Property

'Get/Set the height of this window
Public Property Get height(optional ByVal iClWnd as EWndRectType) As Long
  If Exists Then
    height = RectByType(iClWnd)(3)
  Else
    Err.Raise 1, "stdWindow#height", "Window does not exist."
  End If
End Property
Public Property Let height(optional ByVal iClWnd as EWndRectType, ByVal vH As Long)
  If Exists Then
    Dim rect As variant:: rect = RectByType(iClWnd)
    Call MoveWindow(pHandle, rect(0), rect(1), rect(2), vH, True)
  Else
    Err.Raise 1, "stdWindow#height [Let]", "Window does not exist."
  End If
End Property

'Get the ID of the process running this window
Public Property Get ProcessID() As Long
  If Exists Then
    Call GetWindowThreadProcessId(pHandle, ProcessID)
  Else
    Err.Raise 1, "ProcessID", "Window does not exist."
  End If
End Property

'Get the name of the process running this window
Public Property Get ProcessName() As String
  If Exists Then
    Dim sCharBuffer As String, iNumChars As Long
    sCharBuffer = space(256)
    iNumChars = GetWindowModuleFileName(pHandle, sCharBuffer, 256)
    ProcessName = Mid(sCharBuffer, 1, iNumChars)
  Else
    Err.Raise 1, "ProcessName", "Window does not exist."
  End If
End Property

'Get/Set the window's parent window:
Public Property Get Parent() As stdWindow
  #if VBA7 then
    Dim pHwnd As LongPtr
  #else
    Dim pHwnd As Long
  #end if
  pHwnd = GetParent(pHandle)
  If pHwnd <> 0 Then
    Set Parent = stdWindow.CreateFromHwnd(GetParent(pHandle))
  ElseIf Class <> "#32769" Then
    Set Parent = stdWindow.CreateFromDesktop()
  End If
End Property
Public Property Set Parent(ByVal win As stdWindow)
  If Not Exists Then
    Err.Raise 1, "Parent [Set]", "Window does not exist."
  ElseIf Not win.Exists Then
    Err.Raise 1, "Parent [Set]", "New parent window no longer exists."
  Else
    #if VBA7 then
      Dim hOldParent As LongPtr
    #else
      Dim hOldParent As Long
    #end if
    hOldParent = SetParent(pHandle, win.handle)
  End If
End Property

'Get the root window of this Window/ChildWindow
Public Property Get AncestralRoot() As stdWindow
  If Exists Then
    Set AncestralRoot = stdWindow.CreateFromHwnd(GetAncestor(pHandle, apiWindowAncestorType.GA_ROOT))
  Else
    Err.Raise 1, "AncestralRoot", "Window does not exist."
  End If
End Property

'Get/Set the style of the window
Public Property Get Style() As Long
  If Exists Then
    Style = GetWindowLongA(pHandle, apiWindowLongType.GWL_STYLE)
  Else
    Err.Raise 1, "Style", "Window does not exist."
  End If
End Property
Public Property Let Style(ByVal newStyle As Long)
  If Exists Then
    'Clear Error. See return value at https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowlonga
    Err.Clear

    'Set window long
    Dim hResult As Long: hResult = SetWindowLongA(pHandle, apiWindowLongType.GWL_STYLE, newStyle)

    'Check for errors
    If hResult = 0 And Err.LastDllError <> 0 Then Err.Raise 1, "Style [Let]", "Unexpected error in SetWindowLongA", Err.LastDllError
  Else
    Err.Raise 1, "Style [Let]", "Window does not exist."
  End If
End Property

'Get/Set the extended style of the window
Public Property Get StyleEx() As Long
  If Exists Then
    StyleEx = GetWindowLongA(pHandle, apiWindowLongType.GWL_EXSTYLE)
  Else
    Err.Raise 1, "StyleEx", "Window does not exist."
  End If
End Property
Public Property Let StyleEx(ByVal newStyle As Long)
  If Exists Then
    'Clear Error. See return value at https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowlonga
    Err.Clear

    'Set window long
    Dim hResult As Long: hResult = SetWindowLongA(pHandle, apiWindowLongType.GWL_EXSTYLE, newStyle)

    'Check for errors
    If hResult = 0 And Err.LastDllError <> 0 Then Err.Raise 1, "StyleEx [Let]", "Unexpected error in SetWindowLongA", Err.LastDllError
  Else
    Err.Raise 1, "StyleEx [Let]", "Window does not exist."
  End If
End Property

#if VBA7 then
  'Get/Set a pointer to userdata/metadata
  Public Property Get UserData() As LongPtr
    If Exists Then
      UserData = GetWindowLongPtrA(pHandle, apiWindowLongType.GWL_USERDATA)
    Else
      Err.Raise 1, "UserData", "Window does not exist."
    End If
  End Property

  Public Property Let UserData(ByVal newUserData As LongPtr)
    If Exists Then
      'Clear Error. See return value at https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowlonga
      Err.Clear

      'Set window long
      Dim hResult As LongPtr: hResult = SetWindowLongPtrA(pHandle, apiWindowLongType.GWL_USERDATA, newUserData)

      'Check for errors
      If hResult = 0 And Err.LastDllError <> 0 Then Err.Raise 1, "UserData [Let]", "Unexpected error in SetWindowLongA", Err.LastDllError
    Else
      Err.Raise 1, "UserData [Let]", "Window does not exist."
    End If
  End Property

  'Get/Set the WndProc of the window
  Public Property Get WndProc() As LongPtr
    If Exists Then
      WndProc = GetWindowLongPtrA(pHandle, apiWindowLongType.GWL_WNDPROC)
    Else
      Err.Raise 1, "WndProc", "Window does not exist."
    End If
  End Property
  Public Property Let WndProc(ByVal newWndProc As LongPtr)
    If Exists Then
      'Clear Error. See return value at https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowlonga
      Err.Clear

      'Set window long
      Dim hResult As LongPtr: hResult = SetWindowLongPtrA(pHandle, apiWindowLongType.GWL_WNDPROC, newWndProc)

      'Check for errors
      If hResult = 0 And Err.LastDllError <> 0 Then Err.Raise 1, "WndProc [Let]", "Unexpected error in SetWindowLongA", Err.LastDllError
    Else
      Err.Raise 1, "WndProc [Let]", "Window does not exist."
    End If
  End Property
#else
  'Get/Set a pointer to userdata/metadata
  Public Property Get UserData() As Long
    If Exists Then
      UserData = GetWindowLongPtrA(pHandle, apiWindowLongType.GWL_USERDATA)
    Else
      Err.Raise 1, "UserData", "Window does not exist."
    End If
  End Property
  Public Property Let UserData(ByVal newUserData As Long)
    If Exists Then
      'Clear Error. See return value at https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowlonga
      Err.Clear

      'Set window long
      Dim hResult As Long: hResult = SetWindowLongPtrA(pHandle, apiWindowLongType.GWL_USERDATA, newUserData)

      'Check for errors
      If hResult = 0 And Err.LastDllError <> 0 Then Err.Raise 1, "UserData [Let]", "Unexpected error in SetWindowLongA", Err.LastDllError
    Else
      Err.Raise 1, "UserData [Let]", "Window does not exist."
    End If
  End Property

  'Get/Set the WndProc of the window
  Public Property Get WndProc() As Long
    If Exists Then
      WndProc = GetWindowLongPtrA(pHandle, apiWindowLongType.GWL_WNDPROC)
    Else
      Err.Raise 1, "WndProc", "Window does not exist."
    End If
  End Property
  Public Property Let WndProc(ByVal newWndProc As Long)
    If Exists Then
      'Clear Error. See return value at https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowlonga
      Err.Clear

      'Set window long
      Dim hResult As Long: hResult = SetWindowLongPtrA(pHandle, apiWindowLongType.GWL_WNDPROC, newWndProc)

      'Check for errors
      If hResult = 0 And Err.LastDllError <> 0 Then Err.Raise 1, "WndProc [Let]", "Unexpected error in SetWindowLongA", Err.LastDllError
    Else
      Err.Raise 1, "WndProc [Let]", "Window does not exist."
    End If
  End Property
#end if
  



'Get/Set ability to resize
Public Property Get Resizable() As Boolean
  'THICK FRAME style is used to determine if a window is resizable
  Resizable = Style And EWndStyles.WS_THICKFRAME
End Property
Public Property Let Resizable(ByVal v As Boolean)
  If v Then
    Style = Style Or EWndStyles.WS_THICKFRAME
  Else
    Style = Style And (Not EWndStyles.WS_THICKFRAME)
  End If
End Property


'================================================================================================
'=      PUBLIC INSTANCE METHODS
'================================================================================================


'Set hooks for a window
#if VBA7 then
Public Function SetHook(ByVal idHook As apiWindowHookType, ByVal hook As LongPtr, hInstance As LongPtr, dwThreadID As Long) As LongPtr  
#else
Public Function SetHook(ByVal idHook As apiWindowHookType, ByVal hook As Long, hInstance As Long, dwThreadID As Long) As Long
#end if
  'TODO:
  Err.Raise 1, "SetHook", "Error: Not implemented."
End Function

'Search the Window tree for elements which match a certain criteria. Return the first element found.
'@param {ByVal stdICallable<(stdWindow,depth)=>EWndFindResult>} Callback returning
'  EWndFindResult options:
'    EWndFindResult.NoMatchFound/0/False             - Not found, countinue walking
'    EWndFindResult.MatchFound/1/-1/True             - Found, return this element
'    EWndFindResult.NoMatchCancelSearch/2            - Not found, cancel search
'    EWndFindResult.NoMatchSkipDescendents/3,else    - Not found, don't search descendents
'    EWndFindResult.MatchFoundSearchDescendents/4    - Same as EWndFindResult.MatchFound
'@param {EWndFindType=1} - The type of search, 0 for Breadth First Search (BFS) and 1 for Depth First Search (DFS).
' To understand the difference between BFS and DFS take this tree:
'        A
'       / \
'      B   C
'     /   / \
'    D   E   F
' A BFS will walk this tree in the following order: A, B, C, D, E, F
' A DFS will walk this tree in a different order:   A, C, F, E, B, D
'@examples
' ```
' 'Find where name is "hello" and class is "world":
' el.FindFirst(stdLambda.Create("$1.name=""hello"" and $1.class=""world"""))
' 'Find first element named "hello" at depth > 4:
' el.FindFirst(stdLambda.Create("$1.name = ""hello"" AND $2 > 4"))
' ```
Public Function FindFirst(ByVal query As stdICallable, optional byval searchType as EWndFindType=EWndFindType.BreadthFirst) As stdWindow
    Dim stack() As tFindNode
    ReDim stack(0 To 0)
    stack(0).initialised = true
    stack(0).depth = 0
    Set stack(0).element = Me
    
    Dim length As Long: length = 1
    Dim index As Long: index = -1

    'Bind globals to query
    Call BindGlobals(query)

    'Loop over the stack/array
    While length > 0 And index < length
        Dim part As tFindNode
        select case searchType
            case EWndFindType.DepthFirst
                'Depth first search, so pop the item out of the stack
                part = stackPopV(stack,length)
            case EWndFindType.BreadthFirst
                'Breadth first search, get item directly out of array, no need to change array size
                index = index + 1
                part = stack(index)
            case else
                Err.Raise 1, "stdWindow#FindFirst", "Invalid search type given. Please use EWndFindType"
        end select
        
        With part
            If not .initialised Then Exit Function
            
            'Run query and test options
            Select Case query.Run(.element, .depth)
                Case EWndFindResult.NoMatchFound
                    'Nothing found, search descendents
                    Dim child As stdWindow
                    For Each child In part.element.children
                        Call stackPushV(stack, length, CreateFindNode(.depth + 1, child))
                    Next
                Case EWndFindResult.MatchFound, True, EWndFindResult.MatchFoundSearchDescendents
                    'Found, return element
                    Set FindFirst = .element
                    Exit Function
                Case EWndFindResult.NoMatchCancelSearch
                    'Nothing found, cancel function
                    Set FindFirst = Nothing
                    Exit Function
                case EWndFindResult.NoMatchSkipDescendents
                '    Nothing found, don't search descendents
            End Select
        End With
        'Just make sure no freezing occurs
        DoEvents
    Wend
    
    'Else set to nothing
    Set FindFirst = Nothing
End Function

'Search the Window tree for elements which match a certain criteria. Return all elements found.
'@param {ByVal stdICallable<(stdWindow,depth)=>EWndFindResult>} Callback returning
'  EWndFindResult options:
'    EWndFindResult.NoMatchFound/0/False             - Not found, countinue walking
'    EWndFindResult.MatchFound/1/-1/True             - Found, return this element, won't search descendents of elements found
'    EWndFindResult.NoMatchCancelSearch/2            - Not found, cancel search
'    EWndFindResult.NoMatchSkipDescendents/3,else    - Not found, don't search descendents
'    EWndFindResult.MatchFoundSearchDescendents/4    - Found, return this element, but continue searching descendents
'@param {EWndFindType} - The type of search, 0 for Breadth First Search (BFS) and 1 for Depth First Search (DFS).
' To understand the difference between BFS and DFS take this tree:
'        A
'       / \
'      B   C
'     /   / \
'    D   E   F
' A BFS will walk this tree in the following order: A, B, C, D, E, F
' A DFS will walk this tree in a different order:   A, C, F, E, B, D
'@examples
' ```
' 'Find where name is "hello" and class is "world":
' el.FindAll(stdLambda.Create("$1.name=""hello"" and $1.class=""world"""))
' 'Find all elements with depth <= 4:
' el.FindAll(stdLambda.Create("if $2 < 4 then 4 else if $2 = 4 then 1 else 3"))
' ```
Public Function FindAll(ByVal query As stdICallable, optional byval searchType as EWndFindType=EWndFindType.BreadthFirst) As Collection
    Dim stack() As tFindNode
    ReDim stack(0 To 0)
    stack(0).initialised = true
    stack(0).depth = 0
    Set stack(0).element = Me
    
    Dim length As Long: length = 1
    Dim index As Long: index = -1
    
    'Bind globals to query
    Call BindGlobals(query)

    'Initialise collection
    set FindAll = new Collection

    'Loop over the stack/array
    While length > 0 And index < length
        Dim part as tFindNode
        select case searchType
            case EWndFindType.DepthFirst
                'Depth first search, so pop the item out of the stack
                part = stackPopV(stack,length)
            case EWndFindType.BreadthFirst
                'Breadth first search, get item directly out of array, no need to change array size
                index = index + 1
                part = stack(index)
            case else
                Err.Raise 1, "stdWindow#FindAll", "Invalid search type given. Please use EWndFindType"
        end select
        
        With part
            'When hitting the edge of the stack quit
            if not .initialised then Exit Function

            'Run query and test options
            Dim child as stdWindow
            select case query.Run(.element,.depth)
                Case EWndFindResult.NoMatchFound
                    'Nothing found, search descendents
                    for each child in .element.children
                        Call stackPushV(stack,length,CreateFindNode(.depth+1,child))
                    next
                Case EWndFindResult.MatchFound, True
                    'Found, add element
                    Call FindAll.add(.element)
                Case EWndFindResult.NoMatchCancelSearch
                    'Nothing found, cancel function
                    Exit Function
                Case EWndFindResult.NoMatchSkipDescendents
                    'Nothing found, don't search descendents
                Case EWndFindResult.MatchFoundSearchDescendents
                    Call FindAll.add(.element)
                    for each child in .element.children
                        Call stackPushV(stack,length,CreateFindNode(.depth+1,child))
                    next
            end select
        End with
        'Just make sure no freezing occurs
        DoEvents
    Wend
End Function



'--------------------------------------------------------------------------------
'Automate the window
'--------------------------------------------------------------------------------

Public Sub Quit()
  if Exists then
    if DestroyWindow(pHandle) = 0 then
      Err.Raise Err.LastDllError, "Close", "Failed to close window"
    end if
  else
    Err.Raise 1, "Close", "Window does not exist."
  end if
End Sub

Public Sub Activate()
  If Exists Then
    Call ShowWindow(pHandle, EWndShowStyle.SW_SHOW)
  Else
    Err.Raise 1, "Minimize", "Window does not exist."
  End If
End Sub

'Get all descendents of the stdWindow
'@returns {Collection<stdWindow>} Collection of descendents
Public Function GetDescendents() As Collection
  'Create collection which will be returned
  Dim c As Collection
  Set c = New Collection
  
  Dim child As stdWindow, desc As stdWindow
  For Each child In Children
    'Add children to collection
    c.Add child
    
    'Add all descendents to collection
    For Each desc In child.GetDescendents
        c.Add desc
    Next
  Next
  
  'Return descendents
  Set GetDescendents = c
End Function

Public Sub Redraw()
  Const RDW_INVALIDATE = &H1
  Call RedrawWindow(pHandle, 0&, 0&, RDW_INVALIDATE)
End Sub

Public Function SendMessage(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
  If Exists Then
    SendMessage = apiSendMessage(pHandle, wMsg, wParam, lParam)
  Else
    Err.Raise 1, "SendMessage", "Window does not exist."
  End If
End Function
Public Sub PostMessage(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
  If Exists Then
    'If desktop then broadcasts
    #if VBA7 then
      Dim hwnd As LongPtr
    #else
      Dim hwnd As Long
    #end if
    hwnd = pHandle
    If pHandle = GetDesktopWindow Then hwnd = &HFFFF
    
    If Not apiPostMessage(hwnd, wMsg, wParam, lParam) Then
      Err.Raise 1, "PostMessage", "An unexpected error occurred while posting the message.", Err.LastDllError
    End If
  Else
    Err.Raise 1, "PostMessage", "Window does not exist."
  End If
End Sub

Public Function SendMessageTimeout(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal TimeoutMilliseconds As Long)
  If Exists Then
    'TODO:
    Err.Raise 1, "SendMessageTimeout", "Error: Not implemented."
  Else
    Err.Raise 1, "SendMessageTimeout", "Window does not exist."
  End If
End Function



Public Sub ClickInput(Optional ByVal x As Long = &HFFFF, Optional ByVal y As Long = &HFFFF, Optional ByVal Button As EWndMouseButton)
  'TODO:
  'If x or y are missing, then interpret them as center of window
  If x = &HFFFF Or y = &HFFFF Then
    Dim r() As Long: r = RectWindow
    If x = &HFFFF Then x = CLng((r(2) - r(0)) / 2)
    If y = &HFFFF Then y = CLng((r(3) - r(1)) / 2)
  End If

  Select Case Button
    Case LButton
    Case RButton
    Case MButton

  End Select

  'TODO: Use SendInput()
  Err.Raise 1, "ClickInput", "Error: Not implemented."
End Sub

Public Sub ClickEvent(Optional ByVal x As Long = &H10000, Optional ByVal y As Long = &H10000, Optional ByVal Button As EWndMouseButton, Optional ByVal isDoubleClick As Boolean = False, Optional ByVal wParam As Long = 0)
  'If x or y are missing, then interpret them as center of window
  If x = &HFFFF Or y = &HFFFF Then
    Dim r() As Long: r = RectWindow
    If x > &HFFFF Then x = CLng((r(2) - r(0)) / 2)
    If y > &HFFFF Then y = CLng((r(3) - r(1)) / 2)
  End If

  Const WM_LBUTTONDOWN = &H201
  Const WM_LBUTTONUP = &H202
  Const WM_LBUTTONDBLCLK = &H203
  Const WM_RBUTTONDOWN = &H204
  Const WM_RBUTTONUP = &H205
  Const WM_RBUTTONDBLCLK = &H206
  Const WM_MBUTTONDOWN = &H207
  Const WM_MBUTTONUP = &H208
  Const WM_MBUTTONDBLCLK = &H209

  Dim lParam As Long: lParam = MakeDWord(x, y)

  Select Case Button
    Case LButton
      If isDoubleClick Then
        Call PostMessage(WM_LBUTTONDBLCLK, wParam, lParam)
      Else
        Call PostMessage(WM_LBUTTONDOWN, wParam, lParam)
        Call PostMessage(WM_LBUTTONUP, wParam, lParam)
      End If
    Case RButton
      If isDoubleClick Then
        Call PostMessage(WM_RBUTTONDBLCLK, wParam, lParam)
      Else
        Call PostMessage(WM_RBUTTONDOWN, wParam, lParam)
        Call PostMessage(WM_RBUTTONUP, wParam, lParam)
      End If
    Case MButton
      If isDoubleClick Then
        Call PostMessage(WM_MBUTTONDBLCLK, wParam, lParam)
      Else
        Call PostMessage(WM_MBUTTONDOWN, wParam, lParam)
        Call PostMessage(WM_MBUTTONUP, wParam, lParam)
      End If
  End Select
End Sub

'Uses `SendInput` to send keystrokes to a window.
'@param {ByVal String} Keys to send to the window
'@param {Optional ByVal Boolean} Whether to ignore special chars or not e.g. `{Enter}`
'@param {Optional ByVal Long} Delay between each keystroke
Public Sub SendKeysInput(ByVal sKeys As String, Optional ByVal bRaw As Boolean = False, Optional ByVal keyDelay As Long = 0)
  Dim Keys() As KeyToken: Keys = TokeniseKeys(sKeys)
  If Exists Then
    'TODO:
    Err.Raise 1, "SendKeysInput", "Error: Not implemented."
  Else
    Err.Raise 1, "SendKeysInput", "Window does not exist."
  End If
End Sub

'Uses `SendMessage` to send keystrokes to a window.
'@param {ByVal String} Keys to send to the window
'@param {Optional ByVal Boolean} Whether to ignore special chars or not e.g. `{Enter}`
'@param {Optional ByVal Long} Delay between each keystroke
Public Sub SendKeysEvent(ByVal sKeys As String, Optional ByVal bRaw As Boolean = False, Optional ByVal keyDelay As Long = 0)
  Dim Keys() As KeyToken: Keys = TokeniseKeys(sKeys)
  If Exists Then
    'TODO:
    Err.Raise 1, "SendKeysEvent", "Error: Not implemented."
  Else
    Err.Raise 1, "SendKeysEvent", "Window does not exist."
  End If
End Sub

'Parses a set of keys and converts them into a KeyToken array
Private Function TokeniseKeys(ByVal sKeys As String) As KeyToken()
  'TODO:
  Err.Raise 1, "TokeniseKeys", "Error: Not implemented."
End Function


'================================================================================================
'=      PROTECTED METHODS
'================================================================================================

'Obtain the next window given a stack
Public Function protGetNextDescendent(ByVal stack As Collection, ByVal DFS As Boolean, ByVal Prev As stdWindow) As stdWindow
  If stack.count > 0 Then
    'Get the next window, use popCol if we want to do Depth First Search, else use shiftCol
    Dim oNext As stdWindow
    If DFS Then
      Set oNext = PopCol(stack)
    Else
      Set oNext = ShiftCol(stack)
    End If

    'Add all children to stack
    Dim windows As Collection: Set windows = oNext.Children.Object
    For Each child In windows
      stack.Add child
    Next

    'Return oNext
    Set protGetNextDescendent = oNext
  Else
    protGetNextDescendent = Null
  End If
End Function

'Returns the lookups object
'@returns {Dictionary<Dictionary<Dictionary<string|long>>>}
Public Function protGetLookups()
    if Lookups is nothing then
        'Notes:
        'EWndShowStyle is not needed, use isVisible, isMinimized and isMaximised

        set Lookups = CreateObject("Scripting.Dictionary")
        set Lookups("EWndStyles") = CreateLookupDict(Array( _ 
          "WS_BORDER"            , &H00800000 ,       "WS_CAPTION"        ,  &H00C00000 , _ 
          "WS_CHILD"             , &H40000000 ,       "WS_CHILDWINDOW"    ,  &H40000000 , _ 
          "WS_CLIPCHILDREN"      , &H02000000 ,       "WS_CLIPSIBLINGS"   ,  &H04000000 , _ 
          "WS_DISABLED"          , &H08000000 ,       "WS_DLGFRAME"       ,  &H00400000 , _ 
          "WS_GROUP"             , &H00020000 ,       "WS_HSCROLL"        ,  &H00100000 , _ 
          "WS_ICONIC"            , &H20000000 ,       "WS_MAXIMIZE"       ,  &H01000000 , _ 
          "WS_MAXIMIZEBOX"       , &H00010000 ,       "WS_MINIMIZE"       ,  &H20000000 , _ 
          "WS_MINIMIZEBOX"       , &H00020000 ,       "WS_OVERLAPPED"     ,  &H00000000 , _ 
          "WS_POPUP"             , &H80000000 ,       "WS_SIZEBOX"        ,  &H00040000 , _ 
          "WS_SYSMENU"           , &H00080000 ,       "WS_TABSTOP"        ,  &H00010000 , _ 
          "WS_THICKFRAME"        , &H00040000 ,       "WS_TILED"          ,  &H00000000 , _ 
          "WS_VISIBLE"           , &H10000000 ,       "WS_VSCROLL"        ,  &H00200000 , _ 
          "WS_OVERLAPPEDWINDOW"  , WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX, _ 
          "WS_POPUPWINDOW"       , WS_POPUP Or WS_BORDER Or WS_SYSMENU _ 
        ))
        set Lookups("EWndExStyles") = CreateLookupDict(Array( _ 
          "WS_EX_ACCEPTFILES"          , &H00000010 ,           "WS_EX_APPWINDOW"            , &H00040000 , _ 
          "WS_EX_CLIENTEDGE"           , &H00000200 ,           "WS_EX_COMPOSITED"           , &H02000000 , _ 
          "WS_EX_CONTEXTHELP"          , &H00000400 ,           "WS_EX_CONTROLPARENT"        , &H00010000 , _ 
          "WS_EX_DLGMODALFRAME"        , &H00000001 ,           "WS_EX_LAYERED"              , &H00080000 , _ 
          "WS_EX_LAYOUTRTL"            , &H00400000 ,           "WS_EX_LEFT"                 , &H00000000 , _ 
          "WS_EX_LEFTSCROLLBAR"        , &H00004000 ,           "WS_EX_LTRREADING"           , &H00000000 , _ 
          "WS_EX_MDICHILD"             , &H00000040 ,           "WS_EX_NOACTIVATE"           , &H08000000 , _ 
          "WS_EX_NOINHERITLAYOUT"      , &H00100000 ,           "WS_EX_NOPARENTNOTIFY"       , &H00000004 , _ 
          "WS_EX_NOREDIRECTIONBITMAP"  , &H00200000 ,           "WS_EX_RIGHT"                , &H00001000 , _ 
          "WS_EX_RIGHTSCROLLBAR"       , &H00000000 ,           "WS_EX_RTLREADING"           , &H00002000 , _ 
          "WS_EX_STATICEDGE"           , &H00020000 ,           "WS_EX_TOOLWINDOW"           , &H00000080 , _ 
          "WS_EX_TOPMOST"              , &H00000008 ,           "WS_EX_TRANSPARENT"          , &H00000020 , _ 
          "WS_EX_WINDOWEDGE"           , &H00000100                                          , _ 
          "WS_EX_OVERLAPPEDWINDOW"     , WS_EX_WINDOWEDGE or WS_EX_CLIENTEDGE                , _ 
          "WS_EX_PALETTEWINDOW"        , WS_EX_WINDOWEDGE or WS_EX_TOOLWINDOW or WS_EX_TOPMOST _ 
        ))
        set Lookups("EWndFindResult") = CreateLookupDict(Array( _ 
          "MatchFound"                  , EWndFindResult.MatchFound, _ 
          "MatchFoundSearchDescendents" , EWndFindResult.MatchFoundSearchDescendents, _ 
          "NoMatchFound"                , EWndFindResult.NoMatchFound, _ 
          "NoMatchCancelSearch"         , EWndFindResult.NoMatchCancelSearch, _ 
          "NoMatchSkipDescendents"      , EWndFindResult.NoMatchSkipDescendents _ 
        ))
        set Lookups("EWndState") = CreateLookupDict(Array( _ 
          "Normal", EWndState.Normal, _ 
          "Maximised", EWndState.Maximised, _ 
          "Minimised", EWndState.Minimised _ 
        ))
    end if
    set protGetLookups = lookups
End Function


'================================================================================================
'=      PRIVATE HELPERS
'================================================================================================
Private Function PopCol(ByRef Col As Collection) As stdWindow
  Set PopCol = Col(Col.count)
  Call Col.remove(Col.count)
End Function
Private Function ShiftCol(ByRef Col As Collection) As stdWindow
  Set ShiftCol = Col(1)
  Call Col.remove(1)
End Function
Private Function MakeDWord(ByVal wHi As Integer, ByVal wLo As Integer) As Long
  If wHi And &H8000& Then
    MakeDWord = (((wHi And &H7FFF&) * (&HFFFF& + 1)) Or (wLo And &HFFFF&)) Or &H80000000
  Else
    MakeDWord = (wHi * &HFFFF) + wLo
  End If
End Function

'****************************************
'* HELPERS FOR FindFirst() and FindAll()
'****************************************
'Binds Enumerators to FindFirst
'@param {stdICallable} A callable that may 
'@returns {VOID}
Private Sub BindGlobals(ByRef query as stdICallable)
    Dim bSuccess as boolean
    Dim keys: keys = Lookups.keys()
    For each key in keys
      Call query.SendMessage("bindGlobal", bSuccess, Array(key, Lookups(key)("S2N")))
    next
End Sub

'Constructor for a tFindNode type
'@param {Long} depth of element
'@param {stdWindow} element
'@returns {tFindNode} FindNode struct
Private Function CreateFindNode(ByVal depth as long, ByVal element as stdWindow) as tFindNode
    CreateFindNode.initialised = true
    CreateFindNode.depth = depth
    set CreateFindNode.element = element
End Function

'Given an array `stack`, and length `index`, append `item` onto the stack 
'@param {tFindNode()} stack of elements
'@param {Long} current index in stack
'@param {tFindNode} Item to insert into stack
Private Sub stackPushV(ByRef stack() As tFindNode, ByRef index As Long, ByRef item As tFindNode)
    Dim ub As Long: ub = UBound(stack)
    Dim size As Long: size = ub + 1
    If index > ub Then
        ReDim Preserve stack(0 To size * 2)
    End If
    stack(index).initialised = item.initialised
    stack(index).depth = item.depth
    Set stack(index).element = item.element
    index = index + 1
End Sub

'Given an array `stack`, and `index`=>`n`, pop the nth item off the stack. Used in DFS
'@param {tFindNode()} stack of elements
'@param {Long} current index in stack
'@returns {tFindNode} FindNode found at index
Private Function stackPopV(ByRef stack() As tFindNode, ByRef index As Variant) As tFindNode
    Dim size As Long: size = UBound(stack) + 1
    If index < size / 3 And index > minStackSize Then
        ReDim Preserve stack(0 To CLng(size / 2))
    End If
    index = index - 1
    
    stackPopV.initialised = stack(index).initialised
    stackPopV.depth = stack(index).depth
    set stackPopV.element = stack(index).element
    stack(index).initialised = False
    stack(index).depth = 0
    Set stack(index).element = Nothing
End Function

'From an array of Key-Value pairs, create a dictionary of Key-->Value and Value-->Key lookups
'@param {Variant()} 1D array of Key,Value pairs
'@returns {Dictionary} Dictionary<{S2N: Dictionary<key,value>, N2S: Dictionary<value, key> }>
Private Function CreateLookupDict(arr as Variant) as Object
    Dim oRet as object
    set oRet = CreateObject("Scripting.Dictionary")
    oRet.CompareMode = 1
    Dim i as long

    'Create string --> number dictionary
    set oRet("S2N") = CreateObject("Scripting.Dictionary")
    oRet("S2N").CompareMode = 1
    For i = lbound(arr) to ubound(arr) step 2
        Call oRet("S2N").add(arr(i), arr(i+1))
    next

    'Create number --> string dictionary
    set oRet("N2S") = CreateObject("Scripting.Dictionary")
    oRet("N2S").CompareMode = 1
    For i = lbound(arr) to ubound(arr) step 2
      if isObject(arr(i)) then
        Set oRet("N2S")(arr(i+1)) = arr(i)
      else
        Let oRet("N2S")(arr(i+1)) = arr(i)
      end if
    next

    'Return dictionary
    set CreateLookupDict = oRet
End Function

'Create an atom class string representing an atom.
'@param {Long} Atom to convert to a string
'@returns {String} Classname representing atom
Private Function MakeIntAtom(ByVal iAtom as long) as string
  MakeIntAtom = "#" & iAtom
End Function


Private Function GetWndRect() as boolean





'Public Sub Requires()
'  if IsEmpty(stdEnumerator) then Call Msgbox("Requires `stdEnumerator`", vbInformation) else stdEnumerator.Requires
'
'End Sub
