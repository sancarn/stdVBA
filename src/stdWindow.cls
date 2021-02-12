VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
  WH_FOREGROUNDIDLE =  11
End Enum

'https://www.autohotkey.com/docs_1.0/misc/Styles.htm
Public Enum apiWindowStyles
  WS_BORDER = &H00800000
  WS_CAPTION = &H00C00000
  WS_CHILD = &H40000000
  WS_CHILDWINDOW = &H40000000
  WS_CLIPCHILDREN = &H02000000
  WS_CLIPSIBLINGS = &H04000000
  WS_DISABLED = &H08000000
  WS_DLGFRAME = &H00400000
  WS_GROUP = &H00020000
  WS_HSCROLL = &H00100000
  WS_ICONIC = &H20000000
  WS_MAXIMIZE = &H01000000
  WS_MAXIMIZEBOX = &H00010000
  WS_MINIMIZE = &H20000000
  WS_MINIMIZEBOX = &H00020000
  WS_OVERLAPPED = &H00000000
  WS_POPUP = &H80000000
  WS_SIZEBOX = &H00040000
  WS_SYSMENU = &H00080000
  WS_TABSTOP = &H00010000
  WS_THICKFRAME = &H00040000
  WS_TILED = &H00000000
  WS_VISIBLE = &H10000000
  WS_VSCROLL = &H00200000

  WS_OVERLAPPEDWINDOW = WS_OVERLAPPED OR WS_CAPTION OR WS_SYSMENU OR WS_THICKFRAME OR WS_MINIMIZEBOX OR WS_MAXIMIZEBOX
  WS_POPUPWINDOW = WS_POPUP OR WS_BORDER OR WS_SYSMENU
End Enum



Private Enum apiTagInputType
  INPUT_MOUSE = 0
  INPUT_KEYBOARD = 1
  INPUT_HARDWARE = 2
End Enum
Private Type apiKeyboardInput
  wVirtualKey as Integer   
  wScanCode as Integer
  dwFlags as long
  time as Long
  dwExtraInfo as LongPtr
End type
Private Type apiTagInput
  iType as apiTagInputType
  kbd as apiKeyboardInput
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
  wVirtualKey as Integer  'https://docs.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes
  wScanCode as Integer    '
  iKeyState as KeyState
  iModifiers() as KeyModifier
End Type

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

Private enum apiWindowShowStyles
  SW_HIDE = 0
  SW_SHOWNORMAL = 1      'Shows/Restores + Activates
  SW_SHOWMINIMIZED = 2   'Activates the window and displays it as a minimized window.
  SW_MAXIMIZE = 3        'Maximize
  SW_SHOWNOACTIVATE = 4  'Shows in most recent size + position but doesn't activate
  SW_SHOW = 5            'Activate
  SW_MINIMIZE = 6        'Minimize
  SW_SHOWMINNOACTIVE = 7 'Minimize no activate
  SW_SHOWNA = 8          'Show in current size and position, no activate
  SW_RESTORE = 9         'Restore
  SW_SHOWDEFAULT = 10    'Default window state at start of program
  SW_FORCEMINIMIZE = 11
End Enum

Private enum apiWindowLongType
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

Private enum apiWindowAncestorType
  GA_PARENT = 1
  GA_ROOT = 2
  GA_ROOTOWNER = 3
End enum


'Constructors
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function WindowFromPoint(ByVal x as long, ByVal y as long) as LongPtr

'Getting window data
Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hwnd as LongPtr) as Boolean
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As LongPtr) As Boolean
Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hwnd As LongPtr) As Boolean
Private Declare PtrSafe Function IsHungAppWindow Lib "user32" (ByVal hwnd As LongPtr) As Boolean
Private Declare PtrSafe Function IsZoomed Lib "user32" (ByVal hwnd As LongPtr) As Boolean
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hwnd As LongPtr, ByRef pRect As apiRect) As Boolean
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, ByRef pRect As apiRect) As Boolean
Private Declare PtrSafe Function GetWindowInfo Lib "user32" (ByVal hwnd as LongPtr, ByRef pInf as apiWindowInfo) as Boolean
Private Declare PtrSafe Function GetParent Lib "user32" (ByVal hwnd as LongPtr) as LongPtr
Private Declare PtrSafe Function SetParent Lib "user32" (ByVal hwnd as LongPtr, ByVal hwndParent as LongPtr) as LongPtr
Private Declare PtrSafe Function GetWindowDC Lib "user32" (ByVal hwnd as LongPtr) as LongPtr
Private Declare PtrSafe Function GetWindowLongA Lib "user32" (ByVal hwnd as LongPtr, ByVal nIndex as apiWindowLongType) as Long
Private Declare PtrSafe Function GetWindowLongPtrA Lib "user32" (ByVal hwnd as LongPtr, ByVal nIndex as apiWindowLongType) as LongPtr
Private Declare PtrSafe Function SetWindowLongA Lib "user32" (ByVal hwnd as LongPtr, ByVal nIndex as apiWindowLongType, ByVal dwNewLong as Long) as Long
Private Declare PtrSafe Function SetWindowLongPtrA Lib "user32" (ByVal hwnd as LongPtr, ByVal nIndex as apiWindowLongType, ByVal dwNewPtr as LongPtr) as Long
Private Declare PtrSafe Function GetAncestor Lib "user32" (ByVal hwnd as LongPtr, ByVal nIndex as apiWindowAncestorType) as LongPtr

Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd as LongPtr, ByVal hwndInsertAfter as LongPtr, ByVal x as Long, ByVal y as Long, ByVal width as long, ByVal height as long, ByVal flags as long) as Boolean
Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hwnd as LongPtr, ByVal x as long, ByVal y as long, ByVal width as long, ByVal height as long, ByVal bRepaint as Boolean) as Boolean

'Redrawing window, UpdateWindow can also be used but isn't as safe...
Private Declare PtrSafe Function RedrawWindow Lib "user32" (ByVal hwnd as LongPtr, ByVal lprcUpdate as LongPtr, ByVal hrgnUpdate as LongPtr, ByVal flags as long) as Boolean

'Get children / siblings / parent
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal wCmd As apiWindowRelationship) As LongPtr

'Get process related data
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As LongPtr, ByRef ldpwProcessId As Long) As Long
Private Declare PtrSafe Function GetWindowModuleFileName Lib "user32" Alias "GetWindowModuleFileNameA" (ByVal hwnd As LongPtr, ByVal pszFileName As String, ByVal cchFileNameMax As Long) As Long

'Setting window data
Private Declare PtrSafe Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd as LongPtr, ByVal lpString as string) as boolean

'Automating windows
Private Declare PtrSafe Function apiSendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare PtrSafe Function apiPostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Boolean
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As apiWindowShowStyles) As Long
Private Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hwnd As LongPtr) As Long

'SendKeys
Private Declare PtrSafe Function GetMessageExtraInfo Lib "user32" () as LongPtr

'--------------------------------------------------------------------------------
'Class Declares
'--------------------------------------------------------------------------------
Private pHandle as LongPtr
Private pInitialized as Boolean

'Create a window object from information passed in by this function
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
Public Function Create(ByVal sClassName as string, ByVal sCaption as string, ByVal dwStyle as long, ByVal x as long, ByVal y as long, ByVal Width as long, ByVal Height as Long, ByVal hWndParent as LongPtr, ByVal hMenu as LongPtr, ByVal hInstance as LongPtr, ByVal lpParam as long) as stdWindow
  'TODO:
  CriticalRaise "Create", "Not Implemented"
End Function

'Create a window object used mainly for highlighting areas
'@param {ByVal Long} The x coordinate of the window
'@param {ByVal Long} The y coordinate of the window
'@param {ByVal Long} The width of the window
'@param {ByVal Long} The height of the window
'@param {ByVal Long} The width of the colored border
'@param {ByVal Long} The color of the colored border
'@returns {stdWindow} The created highlighting box
'@remarks https://stackoverflow.com/questions/3970066/creating-a-transparent-window-in-c-win32
Public Function CreateStaticPopup(ByVal x as long, ByVal y as long, ByVal Width as long, ByVal Height as Long, ByVal BorderWidth as long, ByVal BorderColor as long) as stdWindow
  'TODO:
  CriticalRaise "Create", "Not Implemented"
End Function

'Create a window from the desktop window
'@returns {stdWindow} Desktop window
Public Function CreateFromDesktop() as stdWindow
  set CreateFromDesktop = new stdWindow
  Call CreateFromDesktop.init(GetDesktopWindow())
End Function

'Create a window object from a window handle
'@param {ByVal LongPtr} Handle to window (hwnd) to create window object for
'@returns {stdWindow} Specificed window.
Public Function CreateFromHwnd(ByVal hwnd as LongPtr) as stdWindow
  set CreateFromDesktop = new stdWindow
  Call CreateFromDesktop.init(hwnd)
End Function


'Find and Create a window object for a window displayed intersecting a point on the screen.
'@param {ByVal Long} X of Point to find window at
'@param {ByVal Long} Y of Point to find window at
'@returns {stdWindow} Window intersecting point.
Public Function CreateFromPoint(ByVal x as Long, ByVal y as Long) as stdWindow
  set CreateFromPoint = new stdWindow
  Call CreateFromPoint.init(WindowFromPoint(x,y))
End Function

''Create from Shell event
'Public Function CreateFromEvent() as stdWindow
'
'End Function

'Create a window object from an object which implements either IOleWindow, IInternetSecurityMgrSite or IShellView. Uses shell API's IUnknown_GetWindow internally.
'@param {ByVal IUnknown} Object which implements GetWindow() method
'@returns {stdWindow} Window specified by object
Public Function CreateFromIUnknown(ByVal obj as IUnknown) as stdWindow
  Dim hwnd as LongPtr, hResult as long
  hResult = IUnknown_GetWindow(obj, VarPtr(hwnd))
  if hResult = 0 then
    Set CreateFromIUnknown = new stdWindow
    Call CreateFromIUnknown.init(hwnd)
  elseif hResult = -2147467262 then
    CriticalRaise "CreateFromIUnknown", "This object does not implement IOleWindow, IInternetSecurityMgrSite or IShellView, and thus cannot retrieve the window assosciated with the object."
  else
    CriticalRaise "CreateFromIUnknown", "An unknown error has occurred.", hResult
  end if
End Function

'Initialize a window object
'@protected
'@param {ByVal LongPtr} Handle to window (hwnd) to create window object for
'@returns {stdWindow} Specificed window.
Friend Sub Init(ByVal hwnd as LongPtr)
  pHandle = hwnd
  pInitialized = true
End Sub



'--------------------------------------------------------------------------------
'Access window information
'--------------------------------------------------------------------------------

'Get the hWND / window ID of the window
Public Property Get handle() as LongPtr
  handle = pHandle
End Property

'Get the handle to the display context for the window
Public Property Get hDC() as LongPtr
  hDC = GetWindowDC(pHandle)
End Property

'Detect if the window exists
Public Property Get Exists as Boolean
  Exists = IsWindow(pHandle)
End Property

'Detect if the window is visible
Public Property Get IsVisible() as Boolean
  if Exists then
    IsVisible = IsWindowVisible(pHandle)
  else
    CriticalRaise "IsVisible", "Window does not exist."
  end if
End Property

'Detect if the window is minimised
Public Property Get IsMinimised() as Boolean
  If Exists then
    IsMinimised = IsIconic(pHandle)
  else
    CriticalRaise "IsMinimised", "Window does not exist."
  end if
End Property

'Detect if the window is maximised
Public Property Get IsMaximised() as Boolean
  If Exists then
    IsMaximised = IsZoomed(pHandle)
  else
    CriticalRaise "IsMaximised", "Window does not exist."
  end if
End Property

'Detect if the window is hanging/frozen
Public Property Get IsFrozen() as Boolean
  if Exists then
    IsFrozen = IsHungAppWindow(pHandle)
  else
    CriticalRaise "IsFrozen", "Window does not exist."
  end if
End Property

'Get/Set the window caption
Public Property Get Caption() as string
  if Exists then
    Dim sCharBuffer as string, iNumChars as long
    iNumChars = GetWindowText(hwnd, sCharBuffer, 256)
    Caption = Mid(sCharBuffer, 1, iNumChars)
  else
    CriticalRaise "Caption", "Window does not exist."
  end if
End Property
Public Property Let Caption(ByVal s as string)
  if Exists then
    if not SetWindowText(pHandle,s) then
      CriticalRaise "Caption [Let]", "Window text could not be set."
    end if
  else
    CriticalRaise "Caption [Let]", "Window does not exist."
  end if
End Property

'Get the window class
Public Property Get Class() as string
  if Exists then
    Dim sCharBuffer as string, iNumChars as long
    iNumChars = GetClassName(hwnd, sCharBuffer, 256)
    Class = Mid(sCharBuffer, 1, iNumChars)
  else
    CriticalRaise "Class", "Window does not exist."
  end if
End Property

'Get the Client rect - I.E. Position and Size of Window's Client area
Public Property Get RectClient() as Long()
  if Exists then
    Dim rect as apiRect
    if not GetClientRect(pHandle,rect) then
      Dim fRet(0 to 3) as Double
      With rect
        fRet(0) = .left
        fRet(1) = .top
        fRet(2) = .right - .left
        fRet(3) = .bottom - .top
        RectClient = fRet
      end with 
    else
      CriticalRaise "RectClient", "Cannot get client rect.", Err.LastDllError
    end if
  else
    CriticalRaise "RectClient", "Window does not exist."
  end if
End Property

'Get/Set the Window rect - I.E. Position and Size of window
Public Property Get RectWindow() as Long()
  if Exists then
    Dim rect as apiRect
    if not GetWindowRect(pHandle,rect) then
      Dim fRet(0 to 3) as Double
      With rect
        fRet(0) = .left
        fRet(1) = .top
        fRet(2) = .right - .left
        fRet(3) = .bottom - .top
        RectWindow = fRet
      end with 
    else
      CriticalRaise "RectWindow", "Cannot get window rect.", Err.LastDllError
    end if
  else
    CriticalRaise "RectWindow", "Window does not exist."
  end if
End Property
Public Property Let RectWindow(rect() as Long)
  if Exists then
    Call MoveWindow(pHandle, rect(0), rect(1), rect(2), rect(3), true)
  else
    CriticalRaise "RectWindow [Let]", "Window does not exist."
  end if
End Property

'Get/Set the X position of this window
Public Property Get X() as Long
  if Exists then
    X = Me.RectWindow(0)
  else
    CriticalRaise "X", "Window does not exist."
  end if
End Property
Public Property Let X(ByVal vX as Long)
  if Exists then
    Dim rect() as long: rect = RectWindow
    Call MoveWindow(pHandle, vX, rect(1), rect(2), rect(3), true)
  else
    CriticalRaise "X [Let]", "Window does not exist."
  end if
End Property

'Get/Set the Y position of this window
Public Property Get Y() as Long
  if Exists then
    Y = Me.RectWindow(1)
  else
    CriticalRaise "Y", "Window does not exist."
  end if
End Property
Public Property Let Y(ByVal vY as Long)
  if Exists then
    Dim rect() as long: rect = RectWindow
    Call MoveWindow(pHandle, rect(0), vY, rect(2), rect(3), true)
  else
    CriticalRaise "Y [Let]", "Window does not exist."
  end if
End Property

'Get/Set the width of this window
Public Property Get Width() as Long
  if Exists then
    Width =  = Me.RectWindow(2)
  else
    CriticalRaise "Width", "Window does not exist."
  end if
End Property
Public Property Let Width(ByVal vW as Long)
  if Exists then
    Dim rect() as long: rect = RectWindow
    Call MoveWindow(pHandle, rect(0), rect(1), vW, rect(3), true)
  else
    CriticalRaise "Width [Let]", "Window does not exist."
  end if
End Property

'Get/Set the height of this window
Public Property Get Height() as Long
  if Exists then
    Height = Me.RectWindow(3)
  else
    CriticalRaise "Height", "Window does not exist."
  end if
End Property
Public Property Let Height(ByVal vH as Long)
  if Exists then
    Dim rect() as long: rect = RectWindow
    Call MoveWindow(pHandle, rect(0), rect(1), rect(2), vH, true)
  else
    CriticalRaise "Height [Let]", "Window does not exist."
  end if
End Property

'Get the ID of the process running this window
Public Property Get ProcessID() as long
  if Exists then
    Call GetWindowThreadProcessId(pHandle, ProcessID)
  else
    CriticalRaise "ProcessID", "Window does not exist."
  end if
End Property

'Get the name of the process running this window
Public Property Get ProcessName() as string
  if Exists then
    Dim sCharBuffer as string, iNumChars as long
    iNumChars = GetWindowModuleFileName(pHandle, sCharBuffer, 256)
    ProcessName = Mid(sCharBuffer, 1, iNumChars)
  else
    CriticalRaise "ProcessName", "Window does not exist."
  end if
End Property

'Get/Set the window's parent window:
Public Property Get Parent() as stdWindow
  set Parent = stdWindow.CreateFromHwnd(GetParent(pHandle))
End Property
Public Property Set Parent(ByVal win as stdWindow)
  if not Exists then
    CriticalRaise "Parent [Set]", "Window does not exist."
  if not win.Exists then
    CriticalRaise "Parent [Set]", "New parent window no longer exists."
  else
    Dim hOldParent as LongPtr
    hOldParent = SetParent(pHandle, win.Handle)
  end if
End Property

'Get the root window of this Window/ChildWindow
Public Property Get AncestralRoot() as stdWindow
  if Exists then
    set AncestralRoot = stdWindow.CreateFromHwnd(GetAncestor(hwnd,apiWindowAncestorType.GA_ROOT))
  else
    CriticalRaise "AncestralRoot", "Window does not exist."
  end if
End Property

'Get/Set the style of the window
Public Property Get Style() as Long
  if Exists then
    Style = GetWindowLongA(pHandle, apiWindowLongType.GWL_STYLE)
  else
    CriticalRaise "Style", "Window does not exist."
  end if
End Property
Public Property Let Style(ByVal newStyle as Long)
  if Exists then
    'Clear Error. See return value at https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowlonga
    Err.Clear

    'Set window long
    Dim hResult as Long: hResult = SetWindowLongA(pHandle, apiWindowLongType.GWL_STYLE, newStyle)

    'Check for errors
    if hResult = 0 and Err.LastDllError <> 0 then CriticalRaise "Style [Let]", "Unexpected error in SetWindowLongA", Err.LastDllError
  else
    CriticalRaise "Style [Let]", "Window does not exist."
  end if
End Property

'Get/Set the extended style of the window
Public Property Get StyleEx() as Long
  if Exists then
    StyleEx = GetWindowLongA(pHandle, apiWindowLongType.GWL_EXSTYLE)
  else
    CriticalRaise "StyleEx", "Window does not exist."
  end if
End Property
Public Property Let StyleEx(ByVal newStyle as Long)
  if Exists then
    'Clear Error. See return value at https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowlonga
    Err.Clear

    'Set window long
    Dim hResult as Long: hResult = SetWindowLongA(pHandle, apiWindowLongType.GWL_EXSTYLE, newStyle)

    'Check for errors
    if hResult = 0 and Err.LastDllError <> 0 then CriticalRaise "StyleEx [Let]", "Unexpected error in SetWindowLongA", Err.LastDllError
  else
    CriticalRaise "StyleEx [Let]", "Window does not exist."
  end if
End Property

'Get/Set a pointer to userdata/metadata
Public Property Get UserData() as LongPtr
  if Exists then
    UserData = GetWindowLongPtrA(pHandle, apiWindowLongType.GWL_USERDATA)
  else
    CriticalRaise "UserData", "Window does not exist."
  end if
End Property
Public Property Let UserData(ByVal newUserData as LongPtr)
  if Exists then
    'Clear Error. See return value at https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowlonga
    Err.Clear

    'Set window long
    Dim hResult as LongPtr: hResult = SetWindowLongPtrA(pHandle, apiWindowLongType.GWL_USERDATA, newUserData)

    'Check for errors
    if hResult = 0 and Err.LastDllError <> 0 then CriticalRaise "UserData [Let]", "Unexpected error in SetWindowLongA", Err.LastDllError
  else
    CriticalRaise "UserData [Let]", "Window does not exist."
  end if
End Property

'Get/Set the WndProc of the window
Public Property Get WndProc() as LongPtr
  if Exists then
    WndProc = GetWindowLongPtrA(pHandle, apiWindowLongType.GWL_WNDPROC)
  else
    CriticalRaise "WndProc", "Window does not exist."
  end if
End Property
Public Property Let WndProc(ByVal newWndProc as LongPtr)
  if Exists then
    'Clear Error. See return value at https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowlonga
    Err.Clear

    'Set window long
    Dim hResult as LongPtr: hResult = SetWindowLongPtrA(pHandle, apiWindowLongType.GWL_WNDPROC, newWndProc)

    'Check for errors
    if hResult = 0 and Err.LastDllError <> 0 then CriticalRaise "WndProc [Let]", "Unexpected error in SetWindowLongA", Err.LastDllError
  else
    CriticalRaise "WndProc [Let]", "Window does not exist."
  end if
End Property

'Get/Set ability to resize
Public Property Get Resizable() as Boolean
  'THICK FRAME style is used to determine if a window is resizable
  Resizable = Style AND apiWindowStyles.WS_THICKFRAME
End Property
Public Property Let Resizable(ByVal v as Boolean)
  if v then 
    Style = Style OR apiWindowStyles.WS_THICKFRAME
  else
    Style = Style AND (Not apiWindowStyles.WS_THICKFRAME)
  end if
End Property

'Set hooks for a window
Public Function SetHook(ByVal idHook as apiWindowHookType, ByVal hook as LongPtr, hInstance as LongPtr, dwThreadID as long) as LongPtr
  'TODO:
End Function




'--------------------------------------------------------------------------------
'Automate the window
'--------------------------------------------------------------------------------

Public Sub Redraw()
  const RDW_INVALIDATE = &H1
  RedrawWindow(pHandle,0&,0&,RDW_INVALIDATE)
End Sub

Public Function SendMessage(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
  if Exists then
    SendMessage = apiSendMessage(pHandle, wMsg, wParam, lParam)
  else
    CriticalRaise "SendMessage", "Window does not exist."
  end if
End Function
Public Sub PostMessage(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
  if Exists then
    'If desktop then broadcasts
    Dim hwnd as long: hwnd = pHandle
    if pHandle = GetDesktopWindow then hwnd = &HFFFF
    
    if not apiPostMessage(hwnd, wMsg, wParam, lParam) then
      CriticalRaise "PostMessage", "An unexpected error occurred while posting the message.", Err.LastDllError
    end if
  else
    CriticalRaise "PostMessage", "Window does not exist."
  end if
End Sub

Public Function SendMessageTimeout(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal TimeoutMilliseconds as long)
  if Exists then
    'TODO:
  else
    CriticalRaise "SendMessage", "Window does not exist."
  end if
End Function



Public Sub ClickInput(optional ByVal x as long = &HFFFF, Optional ByVal y as long = &HFFFF, Optional ByVal Button as EMouseButton)
  'TODO:
  'If x or y are missing, then interpret them as center of window
  if x = &HFFFF or y = &HFFFF then
    Dim r() as Long: r = RectWindow
    if x = &HFFFF then x = CLng((r(2) - r(0))/2)
    if y = &HFFFF then y = CLng((r(3) - r(1))/2)
  end if

  select case Button
    case LButton
    case RButton
    case mButton

  end select

  'TODO: Use SendInput()
End Sub

Public Sub ClickEvent(optional ByVal x as long = &H10000, Optional ByVal y as long = &H10000, Optional ByVal Button as EMouseButton, Optional ByVal isDoubleClick as boolean = false ,Optional ByVal wParam as Long = 0)
  'If x or y are missing, then interpret them as center of window
  if x = &HFFFF or y = &HFFFF then
    Dim r() as Long: r = RectWindow
    if x > &HFFFF then x = CLng((r(2) - r(0))/2)
    if y > &HFFFF then y = CLng((r(3) - r(1))/2)
  end if

  const WM_LBUTTONDOWN = &H0201
  const WM_LBUTTONUP = &H0202
  const WM_LBUTTONDBLCLK = &H0203
  const WM_RBUTTONDOWN = &H0204
  const WM_RBUTTONUP = &H0205
  const WM_RBUTTONDBLCLK = &H0206
  const WM_MBUTTONDOWN = &H0207
  const WM_MBUTTONUP = &H0208
  const WM_MBUTTONDBLCLK = &H0209

  Dim lParam as long: lParam = MakeDWord(x,y)

  select case Button
    case LButton
      if isDoubleClick then
        Call PostMessage(WM_LBUTTONDBLCLK,wParam,lParam)
      else
        Call PostMessage(WM_LBUTTONDOWN,wParam,lParam)
        Call PostMessage(WM_LBUTTONUP,wParam,lParam)
      end if
    case RButton
      if isDoubleClick then
        Call PostMessage(WM_RBUTTONDBLCLK,wParam,lParam)
      else
        Call PostMessage(WM_RBUTTONDOWN,wParam,lParam)
        Call PostMessage(WM_RBUTTONUP,wParam,lParam)
      end if
    case MButton
      if isDoubleClick then
        Call PostMessage(WM_MBUTTONDBLCLK,wParam,lParam)
      else
        Call PostMessage(WM_MBUTTONDOWN,wParam,lParam)
        Call PostMessage(WM_MBUTTONUP,wParam,lParam)
      end if
  end select
End Sub

'Uses `SendInput` to send keystrokes to a window.
'@param {ByVal String} Keys to send to the window
'@param {Optional ByVal Boolean} Whether to ignore special chars or not e.g. `{Enter}`
'@param {Optional ByVal Long} Delay between each keystroke
Public Sub SendKeysInput(ByVal sKeys as string, optional ByVal bRaw as boolean = false, optional byVal keyDelay as long = 0)
  Dim Keys() as KeyToken: vKeys = TokeniseKeys(sKeys)
  if Exists then
    'TODO:
  else
    CriticalRaise "SendKeysInput", "Window does not exist."
  end if
End Sub

'Uses `SendMessage` to send keystrokes to a window.
'@param {ByVal String} Keys to send to the window
'@param {Optional ByVal Boolean} Whether to ignore special chars or not e.g. `{Enter}`
'@param {Optional ByVal Long} Delay between each keystroke
Public Sub SendKeysEvent(ByVal sKeys as string, optional ByVal bRaw as boolean = false, optional byVal keyDelay as long = 0)
  Dim Keys() as KeyToken: vKeys = TokeniseKeys(sKeys)
  if Exists then
    'TODO:
  else
    CriticalRaise "SendKeysEvent", "Window does not exist."
  end if
End Sub

'Parses a set of keys and converts them into a KeyToken array 
Private Function TokeniseKeys(ByVal sKeys as string) as KeyToken()
  'TODO:
End Function



Public Sub Show()
  if Exists then
    Call ShowWindow(pHandle,apiWindowShowStyles.SW_SHOWNOACTIVATE)
  else
    CriticalRaise "Show", "Window does not exist."
  end if
End Sub
Public Sub Hide()
  if Exists then
    Call ShowWindow(pHandle,apiWindowShowStyles.SW_HIDE)
  else
    CriticalRaise "Hide", "Window does not exist."
  end if
End Sub
Public Sub Maximize()
  if Exists then
    Call ShowWindow(pHandle,apiWindowShowStyles.SW_MAXIMIZE)
  else
    CriticalRaise "Maximize", "Window does not exist."
  end if
End Sub
Public Sub Minimize()
  if Exists then
    Call ShowWindow(pHandle,apiWindowShowStyles.SW_MINIMIZE)
  else
    CriticalRaise "Minimize", "Window does not exist."
  end if
End Sub

Public Sub Activate()
  if Exists then
    Call ShowWindow(pHandle,apiWindowShowStyles.SW_SHOW)
  else
    CriticalRaise "Minimize", "Window does not exist."
  end if
End Sub

Public Property Get Children() as stdEnumerator
  if Exists then
    Dim hwnd As LongPtr
    hwnd = GetWindow(pHandle, GW_CHILD)
    
    Dim hwnds() As LongPtr, i As Long: i = -1
    Do While (hwnd <> 0)
        i = i + 1
        ReDim Preserve hwnds(i)
        hwnds(i) = hwnd
        hwnd = GetWindow(hwnd, GW_HWNDNEXT)
    Loop
    
    Dim ret as Collection
    set ret = new Collection
    For i = lbound(hwnds) to ubound(hwnds)
      Call ret.add(stdWindow.CreateFromHwnd(hwnds(i)))
    Next
    
    set Children = stdEnumerator.CreateFromIEnumVARIANT(ret)
  else
    CriticalRaise "Children", "Window does not exist."
  end if
End Property

Public Property Get Descendents(Optional ByVal DFS as Boolean = true) as stdEnumerator
  'Don't find all objects at once as this will be too slow, instead load when needed:
  Dim stack as Collection
  set stack = new Collection
  stack.add Me
  set Descendents = stdEnumerator.CreateFromFunction( _ 
    stdCallback.CreateFromObjectMethod(Me, "zProtGetNextDescendent").bind(stack, DFS) _ 
  )
End Function

'Obtain the next window given a stack
Public Function zProtGetNextDescendent(ByVal stack as Collection, ByVal DFS as Boolean, ByVal Prev as stdWindow) as stdWindow
  if stack.count > 0 then
    'Get the next window, use popCol if we want to do Depth First Search, else use shiftCol
    Dim oNext as stdWindow
    if DFS then
      set oNext = popCol(stack)
    else
      set oNext = shiftCol(stack)
    end if

    'Add all children to stack
    Dim windows as Collection: set windows = oNext.Children.Object
    for each child in windows
      stack.add child
    next

    'Return oNext
    set zProtGetNextDescendent = oNext
  else
    zProtGetNextDescendent = null
  end if
End Function


'*******************
'* PRIVATE HELPERS *
'*******************
Private Function PopCol(ByRef col as collection) as stdWindow
  set PopCol = col(col.count)
  Call col.remove(col.count)
End Function
Private Function ShiftCol(ByRef Col as collection) as stdWindow
  set ShiftCol = col(1)
  Call col.remove(1)
End Function
Private Function MakeDWord(wHi as Integer, wLo as Integer) as Long
  If wHi And &H8000& then
    MakeDWord = (((wHi And &H7FFF&) * (&HFFFF& + 1)) Or (wLo And &HFFFF&)) Or &H80000000
  else
    MakeDWord = (wHi * &HFFFF) + wLo
  End If
End Function

Private Sub CriticalRaise(ByVal sMethod as string, ByVal sMessage As String, Optional ByVal ErrorNum as Long)
  'If stdError exists
    If VarType(stdError) Then
      Call stdError.Raise("stdWindow::" & sMethod & sMessage & iif(ErrorNum>0, " (" & ErrorNum & ")",""))
    Else
      Call Err.Raise(ErrorNum, "stdWindow::" & sMethod, sMessage)
      End
    End If
End Sub

'Public Sub Requires()
'  if IsEmpty(stdEnumerator) then Call Msgbox("Requires `stdEnumerator`", vbInformation) else stdEnumerator.Requires
'
'End Sub