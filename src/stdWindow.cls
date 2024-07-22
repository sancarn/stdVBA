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
'   [X] Get/Let isTopmost() as Boolean   'AlwaysOnTop
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
Private Enum apiWindowMessage
  WM_GETICON = &H7F
  WM_SETICON = &H80
End Enum
Private Enum apiWindowIconType
  ICON_SMALL = 0
  ICON_BIG = 1
  ICON_SMALL2 = 2 'Used by WM_GETICON only. Retrieves the small icon provided by the application. If the application does not provide one, the system uses the system-generated icon for that window.
End Enum

'https://www.autohotkey.com/docs_1.0/misc/Styles.htm
Public Enum EWndStyles
  WS_BORDER = &H800000
  WS_CAPTION = &HC00000
  WS_CHILD = &H40000000
  WS_CHILDWINDOW = &H40000000
  WS_CLIPCHILDREN = &H2000000
  WS_CLIPSIBLINGS = &H4000000
  WS_DISABLED = &H8000000
  WS_DLGFRAME = &H400000
  WS_GROUP = &H20000
  WS_HSCROLL = &H100000
  WS_ICONIC = &H20000000
  WS_MAXIMIZE = &H1000000
  WS_MAXIMIZEBOX = &H10000
  WS_MINIMIZE = &H20000000
  WS_MINIMIZEBOX = &H20000
  WS_OVERLAPPED = &H0
  WS_POPUP = &H80000000
  WS_SIZEBOX = &H40000
  WS_SYSMENU = &H80000
  WS_TABSTOP = &H10000
  WS_THICKFRAME = &H40000
  WS_TILED = &H0
  WS_VISIBLE = &H10000000
  WS_VSCROLL = &H200000

  WS_OVERLAPPEDWINDOW = WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX
  WS_POPUPWINDOW = WS_POPUP Or WS_BORDER Or WS_SYSMENU
End Enum
Public Enum EWndExStyles
  WS_EX_ACCEPTFILES = &H10
  WS_EX_APPWINDOW = &H40000
  WS_EX_CLIENTEDGE = &H200
  WS_EX_COMPOSITED = &H2000000
  WS_EX_CONTEXTHELP = &H400
  WS_EX_CONTROLPARENT = &H10000
  WS_EX_DLGMODALFRAME = &H1
  WS_EX_LAYERED = &H80000
  WS_EX_LAYOUTRTL = &H400000
  WS_EX_LEFT = &H0
  WS_EX_LEFTSCROLLBAR = &H4000
  WS_EX_LTRREADING = &H0
  WS_EX_MDICHILD = &H40
  WS_EX_NOACTIVATE = &H8000000
  WS_EX_NOINHERITLAYOUT = &H100000
  WS_EX_NOPARENTNOTIFY = &H4
  WS_EX_NOREDIRECTIONBITMAP = &H200000
  WS_EX_RIGHT = &H1000
  WS_EX_RIGHTSCROLLBAR = &H0
  WS_EX_RTLREADING = &H2000
  WS_EX_STATICEDGE = &H20000
  WS_EX_TOOLWINDOW = &H80
  WS_EX_TOPMOST = &H8
  WS_EX_TRANSPARENT = &H20
  WS_EX_WINDOWEDGE = &H100
  WS_EX_OVERLAPPEDWINDOW = WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE
  WS_EX_PALETTEWINDOW = WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST
End Enum
Public Enum LayeredWindowFlag
  lwfChromaKey = &H1
  lwfOpacity = &H2
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

Private Enum KeyState
  tap
  up
  down
End Enum
Private Enum EVirtualKey
  VK_LBUTTON = &H1:   VK_RBUTTON = &H2:     VK_CANCEL = &H3:     VK_MBUTTON = &H4:     VK_XBUTTON1 = &H5
  VK_XBUTTON2 = &H6:  VK_BACK = &H8:        VK_TAB = &H9:        VK_CLEAR = &HC:       VK_RETURN = &HD
  VK_SHIFT = &H10:    VK_CONTROL = &H11:    VK_ALT = &H12:       VK_PAUSE = &H13:      VK_CAPITAL = &H14
  VK_KANA = &H15:     VK_HANGUEL = &H15:    VK_HANGUL = &H15:    VK_IME_ON = &H16:     VK_JUNJA = &H17
  VK_FINAL = &H18:    VK_HANJA = &H19:      VK_KANJI = &H19:     VK_IME_OFF = &H1A:    VK_ESCAPE = &H1B
  VK_CONVERT = &H1C:  VK_NONCONVERT = &H1D: VK_ACCEPT = &H1E:    VK_MODECHANGE = &H1F: VK_SPACE = &H20
  VK_PRIOR = &H21:    VK_NEXT = &H22:       VK_END = &H23:       VK_HOME = &H24:       VK_LEFT = &H25
  VK_UP = &H26:       VK_RIGHT = &H27:      VK_DOWN = &H28:      VK_SELECT = &H29:     VK_PRINT = &H2A
  VK_EXECUTE = &H2B:  VK_SNAPSHOT = &H2C:   VK_INSERT = &H2D:    VK_DELETE = &H2E:     VK_HELP = &H2F
  
  'Numbers
  VK_0 = &H30:        VK_1 = &H31:          VK_2 = &H32:         VK_3 = &H33:          VK_4 = &H34
  VK_5 = &H35:        VK_6 = &H36:          VK_7 = &H37:         VK_8 = &H38:          VK_9 = &H39
  
  'Alphabet:
  VK_A = &H41:        VK_B = &H42:          VK_C = &H43:         VK_D = &H44:          VK_E = &H45
  VK_F = &H46:        VK_G = &H47:          VK_H = &H48:         VK_I = &H49:          VK_J = &H4A
  VK_K = &H4B:        VK_L = &H4C:          VK_M = &H4D:         VK_N = &H4E:          VK_O = &H4F
  VK_P = &H50:        VK_Q = &H51:          VK_R = &H52:         VK_S = &H53:          VK_T = &H54
  VK_U = &H55:        VK_V = &H56:          VK_W = &H57:         VK_X = &H58:          VK_Y = &H59
  VK_Z = &H5A:

  VK_LWIN = &H5B:     VK_RWIN = &H5C:       VK_APPS = &H5D:      VK_SLEEP = &H5F
  
  'Numpad
  VK_NUMPAD0 = &H60:  VK_NUMPAD1 = &H61:    VK_NUMPAD2 = &H62:   VK_NUMPAD3 = &H63:    VK_NUMPAD4 = &H64:
  VK_NUMPAD5 = &H65:  VK_NUMPAD6 = &H66:    VK_NUMPAD7 = &H67:   VK_NUMPAD8 = &H68:    VK_NUMPAD9 = &H69:
  VK_MULTIPLY = &H6A: VK_ADD = &H6B:        VK_SEPARATOR = &H6C: VK_SUBTRACT = &H6D:   VK_DECIMAL = &H6E:
  VK_DIVIDE = &H6F:
  
  'Function keys
  VK_F1 = &H70:       VK_F2 = &H71:         VK_F3 = &H72:        VK_F4 = &H73:         VK_F5 = &H74:
  VK_F6 = &H75:       VK_F7 = &H76:         VK_F8 = &H77:        VK_F9 = &H78:         VK_F10 = &H79:
  VK_F11 = &H7A:      VK_F12 = &H7B:        VK_F13 = &H7C:       VK_F14 = &H7D:        VK_F15 = &H7E:
  VK_F16 = &H7F:      VK_F17 = &H80:        VK_F18 = &H81:       VK_F19 = &H82:        VK_F20 = &H83:
  VK_F21 = &H84:      VK_F22 = &H85:        VK_F23 = &H86:       VK_F24 = &H87:
  
  'Modifiers
  VK_NUMLOCK = &H90:  VK_SCROLL = &H91:
  VK_LSHIFT = &HA0:   VK_RSHIFT = &HA1:
  VK_LCONTROL = &HA2: VK_RCONTROL = &HA3:
  VK_LALT = &HA4:     VK_RALT = &HA5:

  'Media keys
  VK_BROWSER_BACK = &HA6:    VK_BROWSER_FORWARD = &HA7:   VK_BROWSER_REFRESH = &HA8:   VK_BROWSER_STOP = &HA9:   VK_BROWSER_SEARCH = &HAA:   VK_BROWSER_FAVORITES = &HAB:   VK_BROWSER_HOME = &HAC:
  VK_VOLUME_MUTE = &HAD:   VK_VOLUME_DOWN = &HAE:   VK_VOLUME_UP = &HAF:
  VK_MEDIA_NEXT_TRACK = &HB0:   VK_MEDIA_PREV_TRACK = &HB1:   VK_MEDIA_STOP = &HB2:   VK_MEDIA_PLAY_PAUSE = &HB3:
  VK_LAUNCH_MAIL = &HB4:   VK_LAUNCH_MEDIA_SELECT = &HB5:   VK_LAUNCH_APP1 = &HB6:   VK_LAUNCH_APP2 = &HB7:
  VK_OEM_PLUS = &HBB:   VK_OEM_COMMA = &HBC:   VK_OEM_MINUS = &HBD:   VK_OEM_PERIOD = &HBE: VK_OEM_CLEAR = &HFE:
  VK_OEM_1 = &HBA:   VK_OEM_2 = &HBF:   VK_OEM_3 = &HC0:   VK_OEM_4 = &HDB:   VK_OEM_5 = &HDC:   VK_OEM_6 = &HDD:   VK_OEM_7 = &HDE:   VK_OEM_8 = &HDF:   VK_OEM_102 = &HE2:
  VK_PROCESSKEY = &HE5:   VK_PACKET = &HE7:   VK_ATTN = &HF6:   VK_CRSEL = &HF7:   VK_EXSEL = &HF8:   VK_EREOF = &HF9:   VK_PLAY = &HFA:   VK_ZOOM = &HFB:   VK_NONAME = &HFC:   VK_PA1 = &HFD
End Enum

'
Private Type KeyToken
  wVirtualKey As EVirtualKey  'https://docs.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes
  wScanCode As Integer    '
  iKeyState As KeyState
  iTimes As Long
End Type

'Used by SendInput() to send keys to window
Private Type KeyboardInputEx
  wVk As Integer            'WORD
  wScan As Integer          'WORD
  dwFlags As Long           'DWORD
  time As Long              'DWORD
  #If VBA7 Then             'ULONG_PTR
    dwExtraInfo As LongPtr
  #Else
    dwExtraInfo As Long
  #End If
  padding As Currency
End Type
Private Type KeyboardInput
  InputType As Long        'DWORD = INPUT_KEYBOARD
  ki As KeyboardInputEx
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
  SW_HIDE = 0
  SW_SHOWNORMAL = 1          'Shows/Restores + Activates
  SW_SHOWMINIMIZED = 2       'Activates the window and displays it as a minimized window.
  SW_MAXIMIZE = 3            'Maximize
  SW_SHOWNOACTIVATE = 4      'Shows in most recent size + position but doesn't activate
  SW_SHOW = 5                'Activate
  SW_MINIMIZE = 6            'Minimize
  SW_SHOWMINNOACTIVE = 7     'Minimize no activate
  SW_SHOWNA = 8              'Show in current size and position, no activate
  SW_RESTORE = 9             'Restore
  SW_SHOWDEFAULT = 10        'Default window state at start of program
  SW_FORCEMINIMIZE = 11
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
  matchFound = 1                   'Matched
  MatchFoundSearchDescendents = 4  'Same as `ESearchResult.MatchFound`
  NoMatchFound = 0                 'Not found, continue searching descendents
  NoMatchCancelSearch = 2          'Not found, cancel search
  NoMatchSkipDescendents = 3       'Not found, don't search descendents
End Enum

Private Type tFindNode
  initialised As Boolean
  depth As Long
  element As Object
End Type

Private Enum apiClassLongType
  GCL_MENUNAME = -8
  GCL_HBRBACKGROUND = -10
  GCL_HCURSOR = -12
  GCL_HICON = -14
  GCL_HMODULE = -16
  GCL_CBWNDEXTRA = -18
  GCL_CBCLSEXTRA = -20
  GCL_WNDPROC = -24
  GCL_STYLE = -26
  GCL_HICONSM = -34
End Enum

#If VBA7 Then
  'Constructors
  Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
  Private Declare PtrSafe Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByVal hwnd As LongPtr) As Long
  Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal x As Long, ByVal y As Long) As LongPtr
  Private Declare PtrSafe Function WindowFromAccessibleObject Lib "oleacc" (ByVal pacc As IAccessible, ByRef pHwnd As LongPtr) As Long

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
  Private Declare PtrSafe Function GetClassLongPtrA Lib "user32" (ByVal hwnd As LongPtr, ByVal nIndex As apiClassLongType) As LongPtr
  Private Declare PtrSafe Function GetAncestor Lib "user32" (ByVal hwnd As LongPtr, ByVal nIndex As apiWindowAncestorType) As LongPtr
  Private Declare PtrSafe Function FindWindowExA Lib "user32" (ByVal hwnd As LongPtr, ByVal hwndChildAfter As LongPtr, ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr
  Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hwndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal flags As Long) As Long
  Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal bRepaint As Boolean) As Long
  Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
  
  'Redrawing window, UpdateWindow can also be used but isn't as safe...
  Private Declare PtrSafe Function RedrawWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal lprcUpdate As LongPtr, ByVal hrgnUpdate As LongPtr, ByVal flags As Long) As Long

  'Get children / siblings / parent
  Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal wCmd As apiWindowRelationship) As LongPtr

  'Get process related data
  Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As LongPtr, ByRef ldpwProcessId As Long) As Long
  Private Declare PtrSafe Function GetWindowModuleFileName Lib "user32" Alias "GetWindowModuleFileNameA" (ByVal hwnd As LongPtr, ByVal pszFileName As String, ByVal cchFileNameMax As Long) As Long
  Private Declare PtrSafe Function GetCurrentThreadId Lib "Kernel32" () As Long
  Private Declare PtrSafe Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
  
  'Setting window data
  Private Declare PtrSafe Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String) As Long
  
  'Layered windows
  Private Declare PtrSafe Function GetLayeredWindowAttributes Lib "user32" (ByVal hwnd As LongPtr, ByRef crKey As Long, ByRef bAlpha As Byte, ByRef dwFlags As Long) As Long
  Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As LongPtr, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

  'Automating windows
  Private Declare PtrSafe Function SendMessageA Lib "user32" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As Any, ByVal lParam As Any) As Long
  Private Declare PtrSafe Function PostMessageA Lib "user32" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As Any, ByVal lParam As Any) As Long
  Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As EWndShowStyle) As Long
  Private Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hwnd As LongPtr) As Long
  Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hwnd As LongPtr) As Long

  'SendKeys
  Private Declare PtrSafe Function MapVirtualKeyA Lib "user32" (ByVal uCode As Long, ByVal uMapType As Long) As Long
  Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVK As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As LongPtr)
  Private Declare PtrSafe Function SendInput Lib "user32" (ByVal cInputs As Long, ByRef pInput As KeyboardInput, ByVal cbSize As Integer) As Long
  Private Declare PtrSafe Function GetMessageExtraInfo Lib "user32" () As LongPtr
  Private Declare PtrSafe Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

  'Create windows
  Private Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
  Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
  Private Declare PtrSafe Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As LongPtr
  Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
  Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As LongPtr
  Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As LongPtr
  Private Declare PtrSafe Function Rectangle Lib "gdi32" (ByVal hdc As LongPtr, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As LongPtr
  Private Declare PtrSafe Function UpdateWindow Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
#Else
  Private Enum LongPtr
    [_]
  End Enum
  'Constructors
  Private Declare Function GetDesktopWindow Lib "user32" () As LongPtr
  Private Declare Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByVal hwnd As LongPtr) As Long
  Private Declare Function WindowFromPoint Lib "user32" (ByVal x As Long, ByVal y As Long) As LongPtr
  Private Declare Function WindowFromAccessibleObject Lib "oleacc" (ByVal pacc As IAccessible, ByRef pHwnd As LongPtr) As Long

  'Getting window data
  Private Declare Function IsWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
  Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As LongPtr) As Long
  Private Declare Function IsIconic Lib "user32" (ByVal hwnd As LongPtr) As Long
  Private Declare Function IsHungAppWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
  Private Declare Function IsZoomed Lib "user32" (ByVal hwnd As LongPtr) As Long
  Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
  Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
  Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As LongPtr, ByRef pRect As apiRect) As Long
  Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, ByRef pRect As apiRect) As Long
  Private Declare Function GetWindowInfo Lib "user32" (ByVal hwnd As LongPtr, ByRef pInf As apiWindowInfo) As Long
  Private Declare Function GetParent Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
  Private Declare Function SetParent Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndParent As LongPtr) As LongPtr
  Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
  Private Declare Function GetWindowLongA Lib "user32" (ByVal hwnd As LongPtr, ByVal nIndex As apiWindowLongType) As Long
  Private Declare Function GetWindowLongPtrA Lib "user32" (ByVal hwnd As LongPtr, ByVal nIndex As apiWindowLongType) As LongPtr
  Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As LongPtr, ByVal nIndex As apiWindowLongType, ByVal dwNewLong As Long) As Long
  Private Declare Function SetWindowLongPtrA Lib "user32" (ByVal hwnd As LongPtr, ByVal nIndex As apiWindowLongType, ByVal dwNewPtr As LongPtr) As Long
  Private Declare Function GetClassLongPtrA Lib "user32" (ByVal hwnd As LongPtr, ByVal nIndex As apiClassLongType) As LongPtr
  Private Declare Function GetAncestor Lib "user32" (ByVal hwnd As LongPtr, ByVal nIndex As apiWindowAncestorType) As LongPtr
  Private Declare Function FindWindowExA Lib "user32" (ByVal hwnd As LongPtr, ByVal hwndChildAfter As LongPtr, ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr
  Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hwndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal flags As Long) As Long
  Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal bRepaint As Boolean) As Long
  Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
  
  'Redrawing window, UpdateWindow can also be used but isn't as safe...
  Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal lprcUpdate As LongPtr, ByVal hrgnUpdate As LongPtr, ByVal flags As Long) As Long

  'Get children / siblings / parent
  Private Declare Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal wCmd As apiWindowRelationship) As LongPtr

  'Get process related data
  Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As LongPtr, ByRef ldpwProcessId As Long) As Long
  Private Declare Function GetWindowModuleFileName Lib "user32" Alias "GetWindowModuleFileNameA" (ByVal hwnd As LongPtr, ByVal pszFileName As String, ByVal cchFileNameMax As Long) As Long
  Private Declare Function GetCurrentThreadId Lib "Kernel32" () As Long
  Private Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
  
  'Setting window data
  Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String) As Long
  
  'Layered windows
  Private Declare Function GetLayeredWindowAttributes Lib "user32" (ByVal hwnd As LongPtr, ByRef crKey As Long, ByRef bAlpha As Byte, ByRef dwFlags As Long) As Long
  Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As LongPtr, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

  'Automating windows
  Private Declare Function SendMessageA Lib "user32" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As Any, ByVal lParam As Any) As Long
  Private Declare Function PostMessageA Lib "user32" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As Any, ByVal lParam As Any) As Long
  Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As EWndShowStyle) As Long
  Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As LongPtr) As Long
  Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As LongPtr) As Long

  'SendKeys
  Private Declare Function MapVirtualKeyA Lib "user32" (ByVal uCode As Long, ByVal uMapType As Long) As Long
  Private Declare Sub keybd_event Lib "user32" (ByVal bVK As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As LongPtr)
  Private Declare Function SendInput Lib "user32" (ByVal cInputs As Long, ByRef pInput As KeyboardInput, ByVal cbSize As Integer) As Long
  Private Declare Function GetMessageExtraInfo Lib "user32" () As LongPtr
  Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

  'Create windows
  Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
  Private Declare Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
  Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As LongPtr
  Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
  Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As LongPtr
  Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As LongPtr
  Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As LongPtr, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As LongPtr
  Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
#End If


Private Type TLayeredWindowInfo
  crKey As Long
  bAlpha As Byte
End Type

Private Type TThis
  Handle As LongPtr
  Initialized As Boolean
  LayeredWindowInfo As TLayeredWindowInfo
  Lookups As Object
  isCreatedByVBA as Boolean
End Type
Private This as TThis


'================================================================================================
'=      PUBLIC CONSTRUCTORS
'================================================================================================
'Create a window and return a window object
'@constructor
'@param sClassName - The class name can be any name registered with RegisterClass or RegisterClassEx, provided that the module that registers the class is also the module that creates the window. The class name can also be any of the predefined system class names. For a list of system class names, see the Remarks section.
'@param sCaption   - The name/caption of the window
'@param dwStyle    - The window style for the window
'@param x          - The x coordinate of the window
'@param y          - The y coordinate of the window
'@param width      - The width of the window
'@param height     - The height of the window
'@param hWndParent - Parent window handle. Can be 0 for pop-up windows.
'@param hMenu      - Menu handle. Can be 0 for pop-up windows.
'@param hInstance  - Module Instance handle.
'@param lpParam    - Pointer to a location where extra information is stored. Or ClientCreateStruct (for MDI windows), or null if no extra data required
'@returns - The created window
'@remarks System Class Names: `BUTTON`, `COMBOBOX`, `EDIT`, `LISTBOX`, `MDICLIENT`, `RICHEDIT`, `RICHEDIT_CLASS`, `SCROLLBAR`, `STATIC`
'@example ```vb
'Dim wnd as stdWindow: set wnd = stdWindow.Create("STATIC", 0, 0, 100, 100, "Hello World")
'wnd.visible = true
'```
#If VBA7 Then
Public Function Create(ByVal sClassName As String, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, Optional ByVal sCaption As String = vbNullString, Optional ByVal dwStyle As Long = WS_POPUP, Optional ByVal dwStyleEx As Long = 0, Optional ByVal hWndParent As LongPtr = 0, Optional ByVal hMenu As LongPtr = 0, Optional ByVal hInstance As LongPtr = 0, Optional ByVal lpParam As Long = 0) As stdWindow
#Else
Public Function Create(ByVal sClassName As String, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, Optional ByVal sCaption As String = vbNullString, Optional ByVal dwStyle As Long = WS_POPUP, Optional ByVal dwStyleEx As Long = 0, Optional ByVal hWndParent As Long = 0, Optional ByVal hMenu As Long = 0, Optional ByVal hInstance As Long = 0, Optional ByVal lpParam As Long = 0) As stdWindow
#End If
  Dim hwnd as LongPtr: hwnd = CreateWindowEx(dwStyleEx, sClassName, sCaption, dwStyle, x, y, width, height, hWndParent, hMenu, hInstance, lpParam)
  If hwnd = 0 Then Err.Raise 1, "Create", "Failed to create window."
  
  Set Create = new stdWindow
  Call Create.protInit(hwnd, true)
End Function

'Create a static popup window and return a window object
'@constructor
'@param x           - The x coordinate of the window
'@param y           - The y coordinate of the window
'@param width       - The width of the window
'@param height      - The height of the window
'@param BorderWidth - The width of the colored border
'@param BorderColor - The color of the colored border
'@returns           - The created highlighting box
'@remarks See [Creating a transparent window](https://stackoverflow.com/questions/3970066/creating-a-transparent-window-in-c-win32)
Public Function CreateHighlightRect(ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, Optional ByVal BorderWidth As Long = 5, Optional ByVal BorderColor As Long = 255) As stdWindow
  set CreateHighlightRect = Create("STATIC", x, y, width, height)
  CreateHighlightRect.transparentColor = RGB(255,255,255)
  CreateHighlightRect.isClickThroughEnabled = true
  CreateHighlightRect.Visible = True
  Call UpdateWindow(CreateHighlightRect.Handle)
  
  'Create highlighted border
  Dim hDC as LongPtr: hDC = CreateHighlightRect.hDCClient
  Dim hPen as LongPtr: hPen = CreatePen(0, BorderWidth, BorderColor)
  SelectObject hDC, hPen
  Rectangle hDC, 0, 0, width, height
  SelectObject hDC, 0
  DeleteObject hPen
  Call ReleaseDC(CreateHighlightRect.Handle, hDC)
End Function

'Whenever object is destroyed, destroy window if it was created by VBA
Private Sub Class_Terminate()
  if This.isCreatedByVBA then
    if This.Handle <> 0 then Call DestroyWindow(This.Handle)
  end if
End Sub

'Create a window object from an existing window handle
'@constructor
'@param hwnd - Handle to window (hwnd) to create window object for
'@returns - Specificed window.
#If VBA7 Then
Public Function CreateFromHwnd(ByVal hwnd As LongPtr) As stdWindow
#Else
Public Function CreateFromHwnd(ByVal hwnd As Long) As stdWindow
#End If
  Set CreateFromHwnd = New stdWindow
  Call CreateFromHwnd.protInit(hwnd)
End Function

'Creates an `stdWindow` object from the current running application (e.g. Excel / Word / Powerpoint).
'@constructor
'@param oApp - Optional. Application object to create window from. If not specified, the current running application is used.
'@returns - Window object representing the current running application
'@remarks - This function is only implemented for Excel, Word and PowerPoint. For other applications, find the window with `stdWindow.CreateFromDesktop()` and then use `desktop.children` to find the required window object.
'Alternatively you might be able to use CommandBars("Status Bar") for other applications, similar to that of PowerPoint, but this is untested.
Public Function CreateFromApplication(Optional ByVal oApp as Object = nothing) As stdWindow
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
End Function

'Create a window from the desktop window
'@constructor
'@returns - Desktop window
Public Function CreateFromDesktop() As stdWindow
  Set CreateFromDesktop = CreateFromHwnd(GetDesktopWindow())
End Function


'Find and Create a window object for a window displayed intersecting a point on the screen.
'@constructor
'@param x - X of Point to find window at.
'@param y - Y of Point to find window at.
'@returns - Window intersecting point.
Public Function CreateFromPoint(ByVal x As Long, ByVal y As Long) As stdWindow
  Set CreateFromPoint = New stdWindow
  Call CreateFromPoint.protInit(WindowFromPoint(x, y))
End Function

''Create from Shell event
'Public Function CreateFromEvent() as stdWindow
'
'End Function

'Create a window object from an IUnknown object.
'@constructor
'@param obj - Object which implements GetWindow() method
'@returns - Window specified by the object.
'@remark - Uses IUserWindow::GetWindow() to get the window handle. This is implemented by `IOleWindow`, `IInternetSecurityMgrSite` and `IShellView`.
'@remark - A `UserForm` is an example of an object which implements `IOleWindow`.
Public Function CreateFromIUnknown(ByVal obj As IUnknown) As stdWindow
  Dim hwnd As LongPtr
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

'Create a window object from an IAccessible object
'@constructor
'@param pAcc as IAccessible - Object which implements IAccessible
'@returns - Window specified by IAccessible object
Public Function CreateFromIAccessible(ByVal pAcc As IUnknown) As stdWindow
  Dim hwnd As LongPtr
  If WindowFromAccessibleObject(pAcc, hwnd) = 0 Then
    Set CreateFromIAccessible = CreateFromHwnd(hwnd)
  Else
    Err.Raise 1, "CreateFromIAccessible", "Could not find window from IAccessible."
  End If
End Function

'Create a window object from the context menu. There should only ever be 1 context menu on the system at one time.
'@constructor
'@returns - Window object representing the ContextMenu
Public Function CreateFromContextMenu() As stdWindow
  Dim hwnd As LongPtr
  hwnd = FindWindowExA(Null, Null, MakeIntAtom(&H8000&), vbNullString)
  If hwnd <> 0 Then
    Set CreateFromContextMenu = CreateFromHwnd(hwnd)
  Else
    Set CreateFromContextMenu = Nothing
  End If
End Function

'Create window objects for all desktop windows which have a specified process ID
'@constructor
'@param processID - Process ID to find windows for
'@returns - Collection of visible windows for the process
'@remarks - Functionally equivalent of `stdWindow.CreateFromDesktop().FindAll(stdLambda.Create("$2.exists and $2.visible and $2.ProcessID = $1").bind(processID), iStaticDepthLimit:=1)`
Public Function CreateManyFromProcessId(ByVal processID As Long) As Collection
  Dim ret As Collection: Set ret = New Collection
  Dim win As stdWindow
  For Each win In stdWindow.CreateFromDesktop().children
    If win.exists and win.visible then
      If win.processID = processID Then
        ret.add win
      End If
    end if
  Next
  Set CreateManyFromProcessId = ret
End Function

'Create window objects for all shell windows
'@constructor
'@returns Collection<stdWindow> - Collection of shell windows
Public Function CreateManyFromShellWindows() as Collection
  set CreateManyFromShellWindows = new Collection
  Dim shell: set shell = CreateObject("Shell.Application")
  Dim win as object
  For each win in shell.windows
    CreateManyFromShellWindows.add CreateFromHwnd(win.hwnd)
  next
End Function

'Converts a stdWindow object to a stdAcc object
'@constructor
'@returns Object<stdAcc>|stdAcc - stdAcc object representing the window
'@remark - This function requires stdAcc dependency.
'@remark - For full intellisense, add `stdAccPresent = 1` to the compiler constants
#if stdAccPresent then
Public Function AsAcc() as stdAcc
#else
Public Function AsAcc() as Object
#end if
  if isObject(stdAcc) then
    set AsAcc = stdAcc.CreateFromHwnd(This.Handle)
  else
    Err.Raise 1, "", "This function requires stdAcc dependency."
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
'@param hwnd - Handle to window (hwnd) to create window object for
'@returns - Specificed window.
#If VBA7 Then
Friend Sub protInit(ByVal hwnd As LongPtr, Optional ByVal isCreatedByVBA as Boolean = false)
#Else
Friend Sub protInit(ByVal hwnd As Long, Optional ByVal isCreatedByVBA as Boolean = false)
#End If
  This.Handle = hwnd
  This.Initialized = True
  This.isCreatedByVBA = isCreatedByVBA
End Sub

'Whenever we initialise ensure to get lookups from `stdWindow`
Private Sub Class_Initialize()
    Set This.Lookups = stdWindow.protGetLookups()
End Sub


'--------------------------------------------------------------------------------
'Access window information
'--------------------------------------------------------------------------------

'Get the hWND / window ID of the window
'@returns - Window handle
#If VBA7 Then
Public Property Get handle() As LongPtr
#Else
Public Property Get handle() As Long
#End If
  handle = This.Handle
End Property

'Get the handle to the display context for the window
'@returns - Window display context
#If VBA7 Then
Public Property Get hDC() As LongPtr
#Else
Public Property Get hDC() As Long
#End If
  hDC = GetWindowDC(This.Handle)
End Property

'Get the handle to the display context for the window client
'@returns - Window display context
#If VBA7 Then
Public Property Get hDCClient() As LongPtr
#Else
Public Property Get hDCClient() As Long
#End If
  hDCClient = GetDC(This.Handle)
End Property

'Detect if the window exists
'@returns - `true` if window exists, `false` otherwise
Public Property Get Exists() As Boolean
  Exists = IsWindow(This.Handle)
End Property

'Detect if the window is hanging/frozen
'@returns - `true` if window is hanging/frozen, `false` otherwise
Public Property Get IsFrozen() As Boolean
  If Exists Then
    IsFrozen = IsHungAppWindow(This.Handle)
  Else
    Err.Raise 1, "IsFrozen", "Window does not exist."
  End If
End Property

'Get/Set the window caption
'@returns - Window caption
Public Property Get Caption() As String
  If Exists Then
    Dim sCharBuffer As String, iNumChars As Long
    sCharBuffer = space(256)
    iNumChars = GetWindowText(This.Handle, sCharBuffer, 256)
    Caption = Mid(sCharBuffer, 1, iNumChars)
  Else
    Err.Raise 1, "Caption", "Window does not exist."
  End If
End Property
Public Property Let Caption(ByVal s As String)
  If Exists Then
    If Not CBool(SetWindowText(This.Handle, s)) Then
      Err.Raise 1, "Caption [Let]", "Window text could not be set."
    End If
  Else
    Err.Raise 1, "Caption [Let]", "Window does not exist."
  End If
End Property

'Get the window class
'@returns - Window's win32 class
Public Property Get Class() As String
  If Exists Then
    Dim sCharBuffer As String, iNumChars As Long
    sCharBuffer = space(256)
    iNumChars = GetClassName(This.Handle, sCharBuffer, 256)
    Class = Mid(sCharBuffer, 1, iNumChars)
  Else
    Err.Raise 1, "Class", "Window does not exist."
  End If
End Property

'Get/Let visibility of window
'@returns - `true` if window is visible, `false` otherwise
Public Property Get Visible() As Boolean
  If Exists Then
    Visible = IsWindowVisible(This.Handle)
  Else
    Err.Raise 1, "Visible", "Window does not exist."
  End If
End Property
Public Property Let Visible(ByVal x As Boolean)
  If Exists Then
    If x Then
      Call ShowWindow(This.Handle, EWndShowStyle.SW_SHOWNOACTIVATE)
    Else
      Call ShowWindow(This.Handle, EWndShowStyle.SW_HIDE)
    End If
  Else
    Err.Raise 1, "Visible", "Window does not exist."
  End If
End Property

'Get/Let windowState of window
'@returns - Window state
Public Property Get State() As EWndState
  If Exists Then
    If IsZoomed(This.Handle) Then
      State = EWndState.Maximised
    ElseIf IsIconic(This.Handle) Then
      State = EWndState.Minimised
    Else
      State = EWndState.Normal
    End If
  Else
    Err.Raise 1, "State", "Window does not exist."
  End If
End Property
Public Property Let State(ByVal x As EWndState)
  If Exists Then
    If Visible Then
      Select Case x
        Case EWndState.Normal
          Call ShowWindow(This.Handle, EWndShowStyle.SW_RESTORE)
        Case EWndState.Maximised
          Call ShowWindow(This.Handle, EWndShowStyle.SW_MAXIMIZE)
        Case EWndState.Minimised
          Call ShowWindow(This.Handle, EWndShowStyle.SW_FORCEMINIMIZE)
      End Select
    Else
      Err.Raise 1, "State", "Cannot set window state of a hidden window."
    End If
  Else
    Err.Raise 1, "State", "Window does not exist."
  End If
End Property

'Get the child windows of this window
'@returns Collection<stdWindow> - Collection of child windows
Public Property Get children() As collection
  If Exists Then
    'Define collection to return
    Dim ret As collection
    Set ret = New collection

    'Attempt to get a child window
    Dim childHandle As LongPtr
    
    childHandle = GetWindow(This.Handle, GW_CHILD)

    'If a child window exists, find all children
    If childHandle <> 0 Then
      'Quickly add all hwnds to an array
      'This is a fast operation, object creation is anticipated to be a slower operation, thus is left till afterwards
      Dim childHandles() As LongPtr

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
    End If

    'Return child objects
    Set children = ret
  Else
    Err.Raise 1, "Children", "Window does not exist."
  End If
End Property

'Get the Client rect - I.E. Position and Size of Window's Client area
'@returns Array<Left: double, Top: double, Width: double, Height: double> - Client rect
Public Property Get RectClient() As Double()
  If Exists Then
    Dim rect As apiRect
    If Not GetClientRect(This.Handle, rect) Then
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
'@returns Variant<Array<Left: double, Top: double, Width: double, Height: double>> - Window rect
'@remark - Type is wrapped in Variant, because VBA does not support byval Double array parameters. This is a workaround.
Public Property Get RectWindow() As Variant
  If Exists Then
    Dim rect As apiRect
    If Not GetWindowRect(This.Handle, rect) Then
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
  If isArray(rect) Then
    If Exists Then
      Call MoveWindow(This.Handle, rect(0), rect(1), rect(2), rect(3), True)
    Else
      Err.Raise 1, "RectWindow [Let]", "Window does not exist."
    End If
  Else
    Err.Raise 1, "RectWindow [Let]", "Value must be an array of 4 doubles - left, top, width, height."
  End If
End Property

'Get/Set the Window rect with passed param
'@param iClWnd - Type of rect to get/set. Can be `RectTypeClient` or `RectTypeWindow`
'@returns Variant<Array<Left: double, Top: double, Width: double, Height: double>> - Window rect
Public Property Get RectByType(Optional iClWnd As EWndRectType) As Variant
  Select Case iClWnd
    Case EWndRectType.RectTypeClient
      RectByType = RectClient
    Case EWndRectType.RectTypeWindow
      RectByType = RectWindow
  End Select
End Property
Public Property Let RectByType(Optional iClWnd As EWndRectType, ByVal rect As Variant)
  Select Case iClWnd
    Case EWndRectType.RectTypeClient
      Err.Raise 1, "RectByType [Let]", "Cannot set client rect"
    Case EWndRectType.RectTypeWindow
      RectWindow = rect
  End Select
End Property

'Get/Set the X position of this window
'@param iClWnd - Type of rect to get/set. Can be `RectTypeClient` or `RectTypeWindow`
'@returns - X position of window
Public Property Get x(Optional ByVal iClWnd As EWndRectType) As Long
  If Exists Then
    x = RectByType(iClWnd)(0)
  Else
    Err.Raise 1, "stdWindow#x", "Window does not exist."
  End If
End Property
Public Property Let x(Optional ByVal iClWnd As EWndRectType, ByVal vX As Long)
  If Exists Then
    Dim rect As Variant: rect = RectByType(iClWnd)
    Call MoveWindow(This.Handle, vX, rect(1), rect(2), rect(3), True)
  Else
    Err.Raise 1, "stdWindow#x [Let]", "Window does not exist."
  End If
End Property

'Get/Set the Y position of this window
'@param iClWnd - Type of rect to get/set. Can be `RectTypeClient` or `RectTypeWindow`
'@returns - Y position of window
Public Property Get y(Optional ByVal iClWnd As EWndRectType) As Long
  If Exists Then
    y = RectByType(iClWnd)(1)
  Else
    Err.Raise 1, "stdWindow#y", "Window does not exist."
  End If
End Property
Public Property Let y(Optional ByVal iClWnd As EWndRectType, ByVal vY As Long)
  If Exists Then
    Dim rect As Variant: rect = RectByType(iClWnd)
    Call MoveWindow(This.Handle, rect(0), vY, rect(2), rect(3), True)
  Else
    Err.Raise 1, "stdWindow#y [Let]", "Window does not exist."
  End If
End Property

'Get/Set the width of this window
'@param iClWnd - Type of rect to get/set. Can be `RectTypeClient` or `RectTypeWindow`
'@returns - Width of window
Public Property Get width(Optional ByVal iClWnd As EWndRectType) As Long
  If Exists Then
    width = RectByType(iClWnd)(2)
  Else
    Err.Raise 1, "stdWindow#width", "Window does not exist."
  End If
End Property
Public Property Let width(Optional ByVal iClWnd As EWndRectType, ByVal vW As Long)
  If Exists Then
    Dim rect As Variant:: rect = RectByType(iClWnd)
    Call MoveWindow(This.Handle, rect(0), rect(1), vW, rect(3), True)
  Else
    Err.Raise 1, "stdWindow#width [Let]", "Window does not exist."
  End If
End Property

'Get/Set the height of this window
'@param iClWnd - Type of rect to get/set. Can be `RectTypeClient` or `RectTypeWindow`
'@returns - Height of window
Public Property Get height(Optional ByVal iClWnd As EWndRectType) As Long
  If Exists Then
    height = RectByType(iClWnd)(3)
  Else
    Err.Raise 1, "stdWindow#height", "Window does not exist."
  End If
End Property
Public Property Let height(Optional ByVal iClWnd As EWndRectType, ByVal vH As Long)
  If Exists Then
    Dim rect As Variant:: rect = RectByType(iClWnd)
    Call MoveWindow(This.Handle, rect(0), rect(1), rect(2), vH, True)
  Else
    Err.Raise 1, "stdWindow#height [Let]", "Window does not exist."
  End If
End Property

'Get the ID of the process running this window
'@returns - Process ID
Public Property Get ProcessID() As Long
  If Exists Then
    Call GetWindowThreadProcessId(This.Handle, ProcessID)
  Else
    Err.Raise 1, "ProcessID", "Window does not exist."
  End If
End Property

'Get the name of the process running this window
'@returns - Process name
Public Property Get ProcessName() As String
  If Exists Then
    Dim sCharBuffer As String, iNumChars As Long
    sCharBuffer = space(256)
    iNumChars = GetWindowModuleFileName(This.Handle, sCharBuffer, 256)
    ProcessName = Mid(sCharBuffer, 1, iNumChars)
  Else
    Err.Raise 1, "ProcessName", "Window does not exist."
  End If
End Property

'Get/Set the window's parent window:
'@returns - Parent window
Public Property Get parent() As stdWindow
  Dim pHwnd As LongPtr: pHwnd = GetParent(This.Handle)
  If pHwnd <> 0 Then
    Set parent = stdWindow.CreateFromHwnd(GetParent(This.Handle))
  ElseIf Class <> "#32769" Then
    Set parent = stdWindow.CreateFromDesktop()
  End If
End Property
Public Property Set parent(ByVal win As stdWindow)
  If Not Exists Then
    Err.Raise 1, "Parent [Set]", "Window does not exist."
  ElseIf Not win.Exists Then
    Err.Raise 1, "Parent [Set]", "New parent window no longer exists."
  Else
    Dim hOldParent As LongPtr: hOldParent = SetParent(This.Handle, win.handle)
  End If
End Property

'Get/Set the owner of the window.
'@remark - The owner is not the same as the parent. See: Raymond Chen's article: https://devblogs.microsoft.com/oldnewthing/20100315-00/?p=14613
Public Property Get owner() as stdWindow
  Dim tOwner as LongPtr: tOwner = GetWindowLongPtrA(This.Handle, apiWindowLongType.GWL_HWNDPARENT)
  set owner = stdWindow.CreateFromHwnd(tOwner)
end property
Public Property Set owner(ByVal win as stdWindow)
  If Not Exists Then
    Err.Raise 1, "Owner [Set]", "Window does not exist."
  ElseIf Not win.Exists Then
    Err.Raise 1, "Owner [Set]", "New owner window no longer exists."
  Else
    Call SetWindowLongPtrA(This.Handle, apiWindowLongType.GWL_HWNDPARENT, win.handle)
  End If
End Property

'Set the owner HWND of the window. Useful if you want to set the owner to 0 (no owner).
'@param hwnd - Handle to window to set as owner
'@remark - The owner is not the same as the parent. See: Raymond Chen's article: https://devblogs.microsoft.com/oldnewthing/20100315-00/?p=14613.
'Calling `stdWindow.CreateFromIUnknown(Me).setOwnerHandle(0)` is especially useful, as it raises the Excel Userform to a top level window.
Public Sub setOwnerHandle(ByVal hwnd as LongPtr)
  Call SetWindowLongPtrA(This.Handle, apiWindowLongType.GWL_HWNDPARENT, hwnd)
End Sub

'Get/Set a handle to the icon of the window
'@returns - Icon handle
'@example - 
'```vb
'stdWindow.CreateFromHwnd(hwnd).hIcon = Image1.Picture.Handle`
'stdWindow.CreateFromHwnd(hwnd).hIcon = stdImage.CreateFromFile(myFile).hicon 'Future
'```
Public Property Get HICON() As LongPtr
  If Exists Then
    HICON = SendMessage(apiWindowMessage.WM_GETICON, apiWindowIconType.ICON_SMALL2, 0)
    if HICON = NULL_PTR then HICON = GetClassLongPtrA(This.Handle, apiClassLongType.GCL_HICONSM)
    If HICON = NULL_PTR then Err.Raise 1, "HICON", "No icon found."
  Else
    Err.Raise 1, "HICON", "Window does not exist."
  End If
End Property
Public Property Let HICON(ByVal v As LongPtr)
  If Exists Then
    Call SendMessage(apiWindowMessage.WM_SETICON, apiWindowIconType.ICON_BIG, CLNG(v))
    Call SendMessage(apiWindowMessage.WM_SETICON, apiWindowIconType.ICON_SMALL, CLNG(v))
  Else
    Err.Raise 1, "HICON [Let]", "Window does not exist."
  End If
End Property

'Get the root window of this Window/ChildWindow
'@returns - Searches up the window tree to find the top level window which is a direct child of the desktop window.
Public Property Get AncestralRoot() As stdWindow
  If Exists Then
    Set AncestralRoot = stdWindow.CreateFromHwnd(GetAncestor(This.Handle, apiWindowAncestorType.GA_ROOT))
  Else
    Err.Raise 1, "AncestralRoot", "Window does not exist."
  End If
End Property

'Get/Set the style of the window
'@returns - Window style
Public Property Get Style() As Long
  If Exists Then
    Style = GetWindowLongA(This.Handle, apiWindowLongType.GWL_STYLE)
  Else
    Err.Raise 1, "Style", "Window does not exist."
  End If
End Property
Public Property Let Style(ByVal newStyle As Long)
  If Exists Then
    'Clear Error. See return value at https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowlonga
    Err.Clear

    'Set window long
    Dim hResult As Long: hResult = SetWindowLongA(This.Handle, apiWindowLongType.GWL_STYLE, newStyle)

    'Check for errors
    If hResult = 0 And Err.LastDllError <> 0 Then Err.Raise 1, "Style [Let]", "Unexpected error in SetWindowLongA", Err.LastDllError
  Else
    Err.Raise 1, "Style [Let]", "Window does not exist."
  End If
End Property

'Obtain a semi-colon delimited list of styles
'@returns - semi-colon delimited list of style names e.g. "WS_BORDER;WS_CAPTION"
Public Property Get StyleTexts() as String
  Dim iStyle as Long: iStyle = Style
  static vKeys   as variant: if isEmpty(vKeys)   then vKeys   = This.Lookups("EWndStyles").keys()
  static vStyles as variant: if isEmpty(vStyles) then vStyles = This.Lookups("EWndStyles").items()
  
  Dim s as string
  Dim i as long: For i = 0 to ubound(vKeys)
    if iStyle and vStyles(i) then s = s & vKeys(i) & ";"
  next
  StyleTexts = s
End Property

'Get/Set the extended style of the window
'@returns - Window extended style
Public Property Get StyleEx() As Long
  If Exists Then
    StyleEx = GetWindowLongA(This.Handle, apiWindowLongType.GWL_EXSTYLE)
  Else
    Err.Raise 1, "StyleEx", "Window does not exist."
  End If
End Property
Public Property Let StyleEx(ByVal newStyle As Long)
  If Exists Then
    'Clear Error. See return value at https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowlonga
    Err.Clear

    'Set window long
    Dim hResult As Long: hResult = SetWindowLongA(This.Handle, apiWindowLongType.GWL_EXSTYLE, newStyle)

    'Check for errors
    If hResult = 0 And Err.LastDllError <> 0 Then Err.Raise 1, "StyleEx [Let]", "Unexpected error in SetWindowLongA", Err.LastDllError
  Else
    Err.Raise 1, "StyleEx [Let]", "Window does not exist."
  End If
End Property

'Obtain a semi-colon delimited list of styles
'@returns - semi-colon delimited list of styleex names e.g. "WS_EX_APPWINDOW;WS_EX_TOPMOST"
Public Property Get StyleExTexts() as String
  Dim iStyleEx as Long: iStyleEx = StyleEx
  static vKeys   as variant: if isEmpty(vKeys)   then vKeys   = This.Lookups("EWndExStyles").keys()
  static vStyles as variant: if isEmpty(vStyles) then vStyles = This.Lookups("EWndExStyles").items()
  
  Dim s as string
  Dim i as long: For i = 0 to ubound(vKeys)
    if iStyleEx and vStyles(i) then s = s & vKeys(i) & ";"
  next
  StyleExTexts = s
End Property

'Get/Set a pointer to userdata/metadata
'@returns - This is a pointer to a location where extra information is stored. Or ClientCreateStruct (for MDI windows), or null if no extra data required.
'@remark - This can be set to a VBA object pointer. However, this is not recommended as VBA will still garbage collect the object unless the ref counter is incremented, and if it is the ref counter won't decrement upon window destroy.
'i.e. You have to manually manage the lifecycle of this object.
#If VBA7 Then
  Public Property Get UserData() As LongPtr
#Else
  Public Property Get UserData() As Long
#end if
  If Exists Then
    UserData = GetWindowLongPtrA(This.Handle, apiWindowLongType.GWL_USERDATA)
  Else
    Err.Raise 1, "UserData", "Window does not exist."
  End If
End Property
#If VBA7 then
  Public Property Let UserData(ByVal newUserData As LongPtr)
#else
  Public Property Let UserData(ByVal newUserData As Long)
#end if
  If Exists Then
    'Clear Error. See return value at https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowlonga
    Err.Clear

    'Set window long
    Dim hResult As LongPtr: hResult = SetWindowLongPtrA(This.Handle, apiWindowLongType.GWL_USERDATA, newUserData)

    'Check for errors
    If hResult = 0 And Err.LastDllError <> 0 Then Err.Raise 1, "UserData [Let]", "Unexpected error in SetWindowLongA", Err.LastDllError
  Else
    Err.Raise 1, "UserData [Let]", "Window does not exist."
  End If
End Property

'Get/Set the WndProc of the window
'@returns - Window's WndProc
'@remark - This is a pointer to the window's WndProc function. This can be set to a VBA module function pointer if the Window is owned by Excel.
#If VBA7 then
  Public Property Get WndProc() As LongPtr
#else
  Public Property Get WndProc() As Long
#end if
  If Exists Then
    WndProc = GetWindowLongPtrA(This.Handle, apiWindowLongType.GWL_WNDPROC)
  Else
    Err.Raise 1, "WndProc", "Window does not exist."
  End If
End Property
#If VBA7 then
  Public Property Let WndProc(ByVal newWndProc As LongPtr)
#else
  Public Property Let WndProc(ByVal newWndProc As Long)
#end if
  If Exists Then
    'Clear Error. See return value at https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowlonga
    Err.Clear

    'Set window long
    Dim hResult As LongPtr: hResult = SetWindowLongPtrA(This.Handle, apiWindowLongType.GWL_WNDPROC, newWndProc)

    'Check for errors
    If hResult = 0 And Err.LastDllError <> 0 Then Err.Raise 1, "WndProc [Let]", "Unexpected error in SetWindowLongA", Err.LastDllError
  Else
    Err.Raise 1, "WndProc [Let]", "Window does not exist."
  End If
End Property

  



'Get/Set ability to resize
'@returns - `true` if window is resizable, `false` otherwise
Public Property Get isResizable() As Boolean
  'THICK FRAME style is used to determine if a window is resizable
  isResizable = bitFlagCheck(Style, EWndStyles.WS_THICKFRAME)
End Property
Public Property Let isResizable(ByVal v As Boolean)
  Style = bitFlagSet(Style, EWndStyles.WS_THICKFRAME, v)
End Property

'Get/Set caption visibility
'@returns - `true` if window caption is visible, `false` otherwise
Public Property Get isCaptionVisible() As Boolean
  isCaptionVisible = bitFlagCheck(Style, EWndStyles.WS_CAPTION)
End Property
Public Property Let isCaptionVisible(ByVal v As Boolean)
  Style = bitFlagSet(Style, EWndStyles.WS_CAPTION, v)
  Call Redraw
End Property

'Get/Set border visibility
Public Property Get isBorderVisible() As Boolean
  isBorderVisible = bitFlagCheck(Style, EWndStyles.WS_BORDER)
End Property
Public Property Let isBorderVisible(ByVal v As Boolean)
  Style = bitFlagSet(Style, EWndStyles.WS_BORDER, v)
  Call Redraw
End Property

'Get/Set maximise button visibility
Public Property Get isMaximiseButtonVisible() As Boolean
  isMaximiseButtonVisible = bitFlagCheck(Style, EWndStyles.WS_MAXIMIZEBOX)
End Property
Public Property Let isMaximiseButtonVisible(ByVal v As Boolean)
  Style = bitFlagSet(Style, EWndStyles.WS_MAXIMIZEBOX, v)
End Property

'Get/Set minimise button visibility
Public Property Get isMinimiseButtonVisible() As Boolean
  isMinimiseButtonVisible = bitFlagCheck(Style, EWndStyles.WS_MINIMIZEBOX)
End Property
Public Property Let isMinimiseButtonVisible(ByVal v As Boolean)
  Style = bitFlagSet(Style, EWndStyles.WS_MINIMIZEBOX, v)
End Property

'Get/Set system menu visibility
'@remark - The system menu is the menu that appears when you click the icon in the top left of the window.
'          It contains options such as "Restore", "Move", "Size", "Minimize", "Maximize", "Close".
'          Toggling the visibility will also toggle the visibility of the close button, as this is part of the system menu.
Public Property Get isSystemMenuVisible() As Boolean
  isSystemMenuVisible = bitFlagCheck(Style, EWndStyles.WS_SYSMENU)
End Property
Public Property Let isSystemMenuVisible(ByVal v As Boolean)
  Style = bitFlagSet(Style, EWndStyles.WS_SYSMENU, v)
End Property

'Get/Set vertical scrollbar visibility
'@remark - Scrollbar visibility doesn't do anything on VBA userform windows.
Public Property Get isVerticalScrollbarVisible() As Boolean
  isVerticalScrollbarVisible = bitFlagCheck(Style, EWndStyles.WS_VSCROLL)
End Property
Public Property Let isVerticalScrollbarVisible(ByVal v As Boolean)
  Style = bitFlagSet(Style, EWndStyles.WS_VSCROLL, v)
End Property

'Get/Set horizontal scrollbar visibility
'@remark - Scrollbar visibility doesn't do anything on VBA userform windows.
Public Property Get isHorizontalScrollbarVisible() As Boolean
  isHorizontalScrollbarVisible = bitFlagCheck(Style, EWndStyles.WS_HSCROLL)
End Property
Public Property Let isHorizontalScrollbarVisible(ByVal v As Boolean)
  Style = bitFlagSet(Style, EWndStyles.WS_HSCROLL, v)
End Property

'Get/Set whether the window is a child window
'@returns - `true` if window is a child window, `false` otherwise
Public Property Get isChildWindow() As Boolean
  isChildWindow = bitFlagCheck(Style, EWndStyles.WS_CHILD)
End Property
Public Property Let isChildWindow(ByVal v As Boolean)
  Style = bitFlagSet(Style, EWndStyles.WS_CHILD, v)
End Property

'Get/Set whether the window is a popup window
'@returns - `true` if window is a popup window, `false` otherwise
Public Property Get isPopupWindow() As Boolean
  isPopupWindow = bitFlagCheck(Style, EWndStyles.WS_POPUP)
End Property
Public Property Let isPopupWindow(ByVal v As Boolean)
  Style = bitFlagSet(Style, EWndStyles.WS_POPUP, v)
End Property

'Get/Set topmost / Always on top
'@returns - `true` if window is topmost, `false` otherwise
Public Property Get isAlwaysOnTop() As Boolean
  isAlwaysOnTop = bitFlagCheck(StyleEx, EWndExStyles.WS_EX_TOPMOST)
End Property
Public Property Let isAlwaysOnTop(ByVal v As Boolean)
  Const HWND_TOPMOST = -1
  Const HWND_NOTOPMOST = -2
  Const SWP_NOSIZE = &H1
  Const SWP_NOMOVE = &H2
  Call SetWindowPos(This.Handle, iif(v, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Property

'Get/Set whether the window is an app window
'@returns - `true` if window is an app window, `false` otherwise
'@remark - An app window is a window that is intended to be used as a main window for an application.
'Typically an app window will have an icon in the taskbar, and will be shown in the Alt+Tab menu.
Public Property Get isAppWindow() As Boolean
  isAppWindow = bitFlagCheck(StyleEx, EWndExStyles.WS_EX_APPWINDOW)
End Property
Public Property Let isAppWindow(ByVal v As Boolean)
  StyleEx = bitFlagSet(StyleEx, EWndExStyles.WS_EX_APPWINDOW, v)
End Property

'Get/Set whether the window is a tool window
'@returns - `true` if window is a tool window, `false` otherwise
Public Property Get isToolWindow() As Boolean
  isToolWindow = bitFlagCheck(StyleEx, EWndExStyles.WS_EX_TOOLWINDOW)
End Property
Public Property Let isToolWindow(ByVal v As Boolean)
  StyleEx = bitFlagSet(StyleEx, EWndExStyles.WS_EX_TOOLWINDOW, v)
End Property

'Get/Set whether clicks are passed through the window. I.E. Clicks will be passed to the window behind this one.
'@remark - This is useful in combination with layered windows to create click-through windows.
Public Property Get isClickThroughEnabled() As Boolean
  isClickThroughEnabled = bitFlagCheck(StyleEx, EWndExStyles.WS_EX_TRANSPARENT)
End Property
Public Property Let isClickThroughEnabled(ByVal v As Boolean)
  StyleEx = bitFlagSet(StyleEx, EWndExStyles.WS_EX_TRANSPARENT, v)
End Property

'Get/Set whether the border has a sunken edge
Public Property Get isBorderEdgeSunken() As Boolean
  isBorderEdgeSunken = bitFlagCheck(StyleEx, EWndExStyles.WS_EX_CLIENTEDGE)
End Property
Public Property Let isBorderEdgeSunken(ByVal v As Boolean)
  StyleEx = bitFlagSet(StyleEx, EWndExStyles.WS_EX_CLIENTEDGE, v)
End Property

'Get/Set whether the window is double bordered
Public Property Get isBorderDoubled() As Boolean
  isBorderDoubled = bitFlagCheck(StyleEx, EWndExStyles.WS_EX_DLGMODALFRAME)
End Property
Public Property Let isBorderDoubled(ByVal v As Boolean)
  StyleEx = bitFlagSet(StyleEx, EWndExStyles.WS_EX_DLGMODALFRAME, v)
End Property

'Get/Set whether the window is a layered window
Public Property Get isLayeredWindow() As Boolean
  isLayeredWindow = bitFlagCheck(StyleEx, EWndExStyles.WS_EX_LAYERED)
End Property
Public Property Let isLayeredWindow(ByVal v As Boolean)
  StyleEx = bitFlagSet(StyleEx, EWndExStyles.WS_EX_LAYERED, v)
End Property

'Get/Set the existing layered window flags
Public Property Get LayeredWindowFlags() As LayeredWindowFlag
  If isLayeredWindow Then
    Call GetLayeredWindowAttributes(This.Handle, 0, 0, LayeredWindowFlags)
  Else
    Err.Raise 1, "LayeredWindowFlags", "Window is not a layered window."
  End If
End Property
Public Property Let LayeredWindowFlags(ByVal v As LayeredWindowFlag)
  If isLayeredWindow Then
    Call SetLayeredWindowAttributes(This.Handle, 0, 0, v)
  Else
    Err.Raise 1, "LayeredWindowFlags [Let]", "Window is not a layered window."
  End If
End Property

'Get/Set the transparent color for a window.  for a layered window
'@returns - The transparent color for the window. If isLayeredWindow is false, or the window's LayeredWindowFlags does not have the ChromaKey flag set, this will return -1.
'@remark - Setting this value will also set the window to be a layered window if it is not already.
Public Property Get TransparentColor() As Long
  if isLayeredWindow then 
    if bitFlagCheck(LayeredWindowFlags, LayeredWindowFlag.lwfChromaKey) then
      Call GetLayeredWindowAttributes(This.Handle, TransparentColor, 0, 0)
    else
      TransparentColor = -1
    end if
  else
    TransparentColor = -1
  end if
End Property
Public Property Let TransparentColor(ByVal v As Long)
  isLayeredWindow = true
  This.LayeredWindowInfo.crKey = v
  Dim newFlags as Long: newFlags = bitFlagSet(LayeredWindowFlags, LayeredWindowFlag.lwfChromaKey, True)
  Call SetLayeredWindowAttributes(This.Handle, This.LayeredWindowInfo.crKey, This.LayeredWindowInfo.bAlpha, newFlags)
End Property

'Get/Set the opacity of a layered window
'@remark - Setting this value will also set the window to be a layered window if it is not already.
Public Property Get Opacity() As Byte
  If isLayeredWindow Then
    Call GetLayeredWindowAttributes(This.Handle, 0, lwfOpacity, 0)
  Else
    Err.Raise 1, "Opacity", "Window is not a layered window."
  End If
End Property
Public Property Let Opacity(ByVal v As Byte)
  isLayeredWindow = true
  This.LayeredWindowInfo.bAlpha = v
  Dim newFlags as Long: newFlags = bitFlagSet(LayeredWindowFlags, LayeredWindowFlag.lwfOpacity, True)
  Call SetLayeredWindowAttributes(This.Handle, This.LayeredWindowInfo.crKey, This.LayeredWindowInfo.bAlpha, newFlags)
End Property

'Gets the Shell Window COM object for this window if it exists
'@returns Object<Shell.IWebBrowser2|Nothing> - Shell Window COM object
'@example ```
'  Dim location as string: location = stdWindow.CreateFromHwnd(hwnd).ShellWindow.LocationURL
'```
'@TODO: is there a better way?
Public Property Get ShellWindow() as Object
  if Class = "CabinetWClass" then
    With stdWindow.protGetShellWindowCache(This.Handle)
      if .exists(CStr(This.Handle)) then
        set ShellWindow = .item(CStr(This.Handle))
      end if
    end with
  end if
End Property

'================================================================================================
'=      PUBLIC INSTANCE METHODS
'================================================================================================


'Set hooks for a window
'@param idHook - Type of hook to set
'@param hook - Hook function e.g `AddressOf MyHookFunction`
'@param hInstance - Instance handle of the hook function
'@param dwThreadID - Thread ID of the hook function
'@returns - Handle to the hook
#If VBA7 Then
Public Function SetHook(ByVal idHook As apiWindowHookType, ByVal hook As LongPtr, hInstance As LongPtr, dwThreadID As Long) As LongPtr
#Else
Public Function SetHook(ByVal idHook As apiWindowHookType, ByVal hook As Long, hInstance As Long, dwThreadID As Long) As Long
#End If
  'TODO:
  Err.Raise 1, "SetHook", "Error: Not implemented."
End Function

'Search the Window tree for elements which match a certain criteria. Return the first element found.
'@param query as stdICallable<(stdWindow,depth)=>EWndFindResult> - Callback returning `EWndFindResult` options:
'
' * `EWndFindResult.NoMatchFound`/`0`/`False`        - Not found, countinue walking
' * `EWndFindResult.MatchFound`/`1`/`-1`/`True`      - Found, return this element
' * `EWndFindResult.NoMatchCancelSearch`/`2`         - Not found, cancel search
' * `EWndFindResult.NoMatchSkipDescendents`/`3`,     - Not found, don't search descendents
' * `EWndFindResult.MatchFoundSearchDescendents`/`4` - Same as `EWndFindResult.MatchFound` in this case.
'@param searchType - The type of search, 0 for Breadth First Search (BFS) and 1 for Depth First Search (DFS).
' To understand the difference between BFS and DFS take this tree:
'```
'        A
'       / \
'      B   C
'     /   / \
'    D   E   F
'```
' * A BFS will walk this tree in the following order: A, B, C, D, E, F
' * A DFS will walk this tree in a different order:   A, C, F, E, B, D
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
    Dim stack() As tFindNode
    ReDim stack(0 To 0)
    stack(0).initialised = True
    stack(0).depth = 0
    Set stack(0).element = Me
    
    Dim Length As Long: Length = 1
    Dim index As Long: index = -1

    'Bind globals to query
    Call BindGlobals(query)

    'Loop over the stack/array
    While Length > 0 And index < Length
        Dim part As tFindNode
        Select Case searchType
            Case EWndFindType.DepthFirst
                'Depth first search, so pop the item out of the stack
                part = stackPopV(stack, Length)
            Case EWndFindType.BreadthFirst
                'Breadth first search, get item directly out of array, no need to change array size
                index = index + 1
                part = stack(index)
            Case Else
                Err.Raise 1, "stdWindow#FindFirst", "Invalid search type given. Please use EWndFindType"
        End Select
        
        With part
            If Not .initialised Then Exit Function
            
            'Run query and test options
            Select Case query.Run(.element, .depth)
                Case EWndFindResult.NoMatchFound
                    'Check static depth limit
                    If .depth + 1 <= iStaticDepthLimit Or iStaticDepthLimit = -1 Then
                        'Nothing found, search descendents
                        Dim child As stdWindow
                        For Each child In part.element.children
                            Call stackPushV(stack, Length, CreateFindNode(.depth + 1, child))
                        Next
                    End If
                Case EWndFindResult.matchFound, True, EWndFindResult.MatchFoundSearchDescendents
                    'Found, return element
                    Set FindFirst = .element
                    Exit Function
                Case EWndFindResult.NoMatchCancelSearch
                    'Nothing found, cancel function
                    Set FindFirst = Nothing
                    Exit Function
                Case EWndFindResult.NoMatchSkipDescendents
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
'@param query as stdICallable<(stdWindow,depth)=>EWndFindResult> - Callback returning `EWndFindResult` options:
'
' * `EWndFindResult.NoMatchFound`/`0`/`False`        - Not found, countinue walking
' * `EWndFindResult.MatchFound`/`1`/`-1`/`True`      - Found, return this element, won't search descendents of elements found
' * `EWndFindResult.NoMatchCancelSearch`/`2`         - Not found, cancel search
' * `EWndFindResult.NoMatchSkipDescendents`/`3`,     - Not found, don't search descendents
' * `EWndFindResult.MatchFoundSearchDescendents`/`4` - Found, return this element, but continue searching descendents
'@param searchType - The type of search, 0 for Breadth First Search (BFS) and 1 for Depth First Search (DFS).
' To understand the difference between BFS and DFS take this tree:
'```
'        A
'       / \
'      B   C
'     /   / \
'    D   E   F
'```
' * A BFS will walk this tree in the following order: A, B, C, D, E, F
' * A DFS will walk this tree in a different order:   A, C, F, E, B, D
'@param iStaticDepthLimit - Static depth limit. E.G. if you want to search children only, set this value to 1
'@returns Collection<stdWindow> - Collection of elements found, or `Nothing` if no element found.
'@examples
' ```vb
' 'Find where name is "hello" and class is "world":
' el.FindAll(stdLambda.Create("$1.name=""hello"" and $1.class=""world"""))
' 'Find all elements with depth <= 4:
' el.FindAll(stdLambda.Create("if $2 < 4 then 4 else if $2 = 4 then 1 else 3"))
' ```
Public Function FindAll(ByVal query As stdICallable, Optional ByVal searchType As EWndFindType = EWndFindType.BreadthFirst, Optional ByVal iStaticDepthLimit As Long = -1) As collection
    Dim stack() As tFindNode
    ReDim stack(0 To 0)
    stack(0).initialised = True
    stack(0).depth = 0
    Set stack(0).element = Me
    
    Dim Length As Long: Length = 1
    Dim index As Long: index = -1
    
    'Bind globals to query
    Call BindGlobals(query)

    'Initialise collection
    Set FindAll = New collection

    'Loop over the stack/array
    While Length > 0 And index < Length
        Dim part As tFindNode
        Select Case searchType
            Case EWndFindType.DepthFirst
                'Depth first search, so pop the item out of the stack
                part = stackPopV(stack, Length)
            Case EWndFindType.BreadthFirst
                'Breadth first search, get item directly out of array, no need to change array size
                index = index + 1
                part = stack(index)
            Case Else
                Err.Raise 1, "stdWindow#FindAll", "Invalid search type given. Please use EWndFindType"
        End Select
        
        With part
            'When hitting the edge of the stack quit
            If Not .initialised Then Exit Function

            'Run query and test options
            Dim child As stdWindow
            Select Case query.Run(.element, .depth)
                Case EWndFindResult.NoMatchFound
                    'Check static depth limit
                    If .depth + 1 <= iStaticDepthLimit Or iStaticDepthLimit = -1 Then
                        'Nothing found, search descendents
                        For Each child In .element.children
                            Call stackPushV(stack, Length, CreateFindNode(.depth + 1, child))
                        Next
                    End If
                Case EWndFindResult.matchFound, True
                    'Found, add element
                    Call FindAll.Add(.element)
                Case EWndFindResult.NoMatchCancelSearch
                    'Nothing found, cancel function
                    Exit Function
                Case EWndFindResult.NoMatchSkipDescendents
                    'Nothing found, don't search descendents
                Case EWndFindResult.MatchFoundSearchDescendents
                  'Check static depth limit
                  If .depth + 1 <= iStaticDepthLimit Or iStaticDepthLimit = -1 Then
                      Call FindAll.Add(.element)
                      For Each child In .element.children
                          Call stackPushV(stack, Length, CreateFindNode(.depth + 1, child))
                      Next
                  End If
            End Select
        End With
        'Just make sure no freezing occurs
        DoEvents
    Wend
End Function

'Obtain all text from a window
'@returns - All captions from a window concatenated together (DFS-ordering)
Public Function getAllText() As String
  Dim s As String: s = Me.Caption
  Dim wnd As stdWindow
  For Each wnd In Me.GetDescendents()
    If wnd.Visible Then s = s & vbCrLf & wnd.Caption
  Next
  getAllText = s
End Function


'--------------------------------------------------------------------------------
'Automate the window
'--------------------------------------------------------------------------------

'Close the window
Public Sub Quit()
  If Exists Then
    Call PostMessage(&H10&,0,0)
  Else
    Err.Raise 1, "Close", "Window does not exist."
  End If
End Sub

'Forcefully close the window
Public Sub ForceQuit()
  If Exists Then
    If DestroyWindow(This.Handle) = 0 Then
      Err.Raise Err.LastDllError, "Close", "Failed to close window"
    End If
  Else
    Err.Raise 1, "Close", "Window does not exist."
  End If
End Sub

'Activate the window
Public Sub Activate()
  If Exists Then
    Call setThreadInput(True)
      Call SetForegroundWindow(This.Handle)
    Call setThreadInput(False)
  Else
    Err.Raise 1, "Minimize", "Window does not exist."
  End If
End Sub

'Get all descendents of the stdWindow
'@returns Collection<stdWindow> - Collection of descendents
Public Function GetDescendents() As collection
  'Create collection which will be returned
  Dim c As collection
  Set c = New collection
  
  Dim child As stdWindow, desc As stdWindow
  For Each child In children
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

'Redraw the window
Public Sub Redraw()
  Const RDW_INVALIDATE = &H1
  Call RedrawWindow(This.Handle, 0&, 0&, RDW_INVALIDATE)
End Sub

'Sends a message to the window
'@param wMsg - Message to send
'@param wParam - wParam to send
'@param lParam - lParam to send
'@returns - Result of SendMessage
Public Function SendMessage(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) as Long
  If Exists Then
    SendMessage = SendMessageA(This.Handle, wMsg, wParam, lParam)
  Else
    Err.Raise 1, "SendMessage", "Window does not exist."
  End If
End Function

'Posts a message to the window
'@param wMsg - Message to post
'@param wParam - wParam to post
'@param lParam - lParam to post
Public Sub PostMessage(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
  If Exists Then
    'If desktop then broadcasts
    Dim hwnd As LongPtr: hwnd = This.Handle
    If This.Handle = GetDesktopWindow Then hwnd = &HFFFF
    
    If PostMessageA(hwnd, wMsg, wParam, lParam) =  0 Then
      Err.Raise 1, "PostMessage", "An unexpected error occurred while posting the message.", Err.LastDllError
    End If
  Else
    Err.Raise 1, "PostMessage", "Window does not exist."
  End If
End Sub

'Sends a message to the window and waits for a response within a timeout period
'@param wMsg - Message to send
'@param wParam - wParam to send
'@param lParam - lParam to send
'@param TimeoutMilliseconds - Timeout in milliseconds
'@returns - Result of SendMessageTimeout
'@TODO: Complete this function
Private Function SendMessageTimeout(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal TimeoutMilliseconds As Long) as Long
  If Exists Then
    Err.Raise 1, "SendMessageTimeout", "Error: Not implemented."
  Else
    Err.Raise 1, "SendMessageTimeout", "Window does not exist."
  End If
End Function

'Clicks the window
'@param x - X position to click
'@param y - Y position to click
'@param Button - Button of mouse to click
'@TODO: Complete this function
Private Sub ClickInput(Optional ByVal x As Long = &HFFFF, Optional ByVal y As Long = &HFFFF, Optional ByVal Button As EWndMouseButton)
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

'Clicks the window using `keybd_event`
'@param x - X position to click
'@param y - Y position to click
'@param Button - Button of mouse to click
'@TODO: Complete this function
Private Sub ClickEvent(Optional ByVal x As Long = &HFFFF, Optional ByVal y As Long = &HFFFF, Optional ByVal Button As EWndMouseButton)

End Sub

'Clicks the window using PostMessage
'@param x - X position to click
'@param y - Y position to click
'@param Button - Button of mouse to click
'@param isDoubleClick - Whether to double click or not
'@param wParam - wParam to send
Public Sub ClickMessage(Optional ByVal x As Long = &H10000, Optional ByVal y As Long = &H10000, Optional ByVal Button As EWndMouseButton, Optional ByVal isDoubleClick As Boolean = False, Optional ByVal wParam As Long = 0)
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
'@param sKeys - Keys to send to the window
'@param bRaw - Whether to ignore special chars or not e.g. `{Enter}`
'@param keyDelay - Delay between each keystroke
'@param bAutoRelease - Whether keys pressed down should be auto-released
'@example ```
'Call stdWindow.CreateFromHwnd(Application.VBE.mainwindow.hwnd).SendKeysInput("^a")
'```
Public Sub SendKeysInput(ByVal sKeys As String, Optional ByVal bRaw As Boolean = False, Optional ByVal keyDelay As Long = 0, Optional bAutoRelease As Boolean = True)
  Const INPUT_KEYBOARD As Long = 1
  Const KEYEVENTF_KEYUP = &H2
  Dim keys() As KeyToken: keys = TokeniseKeys(sKeys, bAutoRelease)
  If Exists Then
    Call Activate
    
    'Loop over all keys
    Dim iKey As Long
    For iKey = 1 To UBound(keys)
        Dim key As KeyToken: key = keys(iKey)
        
        'Create generic key signal
        Dim inputKey As KeyboardInput
        inputKey.InputType = INPUT_KEYBOARD
        inputKey.ki.wVk = key.wVirtualKey
        inputKey.ki.wScan = key.wScanCode
        inputKey.ki.time = 0
        
        'Key down input
        If key.iKeyState = tap Or key.iKeyState = down Then
          inputKey.ki.dwFlags = 0
          If SendInput(1, inputKey, LenB(inputKey)) = 0 Then
              Err.Raise Err.LastDllError, "SendKeysInput", "Input might be blocked by another thread (DLL Error: " & Err.LastDllError & ")"
          End If
          Call Sleep(keyDelay)
        End If
        
        'Key up input
        If key.iKeyState = tap Or key.iKeyState = up Then
          inputKey.ki.dwFlags = KEYEVENTF_KEYUP
          If SendInput(1, inputKey, LenB(inputKey)) = 0 Then
              Err.Raise Err.LastDllError, "SendKeysInput", "Input might be blocked by another thread (DLL Error: " & Err.LastDllError & ")"
          End If
          Call Sleep(keyDelay)
        End If
    Next
  Else
    Err.Raise 1, "SendKeysInput", "Window does not exist."
  End If
End Sub

'Uses `kybd_event` to send keystrokes to a window.
'@param sKeys - Keys to send to the window
'@param keyDelay - Delay between each keystroke
'@param bAutoRelease - Whether keys pressed down should be auto-released
'@example `notepadWindow.sendKeysEvent("^a")`
Public Sub SendKeysEvent(ByVal sKeys As String, Optional ByVal keyDelay As Long = 10, Optional ByVal bAutoRelease As Boolean = True)
  Const KEYEVENTF_KEYUP = &H2
  Const WM_KEYDOWN = &H100
  Const WM_KEYUP = &H101
        
  Dim keys() As KeyToken: keys = TokeniseKeys(sKeys, bAutoRelease)
  If Exists Then
    Call setThreadInput(True)
      'Activate window
      Call Activate()
      Call Sleep(keyDelay)
      
      'Use keybd event
      Dim iKey as Long
      For iKey = 1 To UBound(keys)
        Dim key As KeyToken: key = keys(iKey)
        
        'Loop as many times as required times
        For n = 1 To key.iTimes
          'Key down event
          If key.iKeyState = tap Or key.iKeyState = down Then
            Call keybd_event(key.wVirtualKey, key.wScanCode, 0, 0)
            Call Sleep(keyDelay)
          End If
          
          'Key up event
          If key.iKeyState = tap Or key.iKeyState = up Then
            Call keybd_event(key.wVirtualKey, key.wScanCode, KEYEVENTF_KEYUP, 0)
            Call Sleep(keyDelay)
          End If
        Next
      Next
    Call setThreadInput(False)
  Else
    Err.Raise 1, "SendKeysEvent", "Window does not exist."
  End If
End Sub

'Uses `PostMessage` to send keystrokes to a window.
'@param sKeys - Keys to send to the window
'@param keyDelay - Delay between each keystroke
'@param bAutoRelease - Whether keys pressed down should be auto-released
'@remark It should be noted that this method is extremely unstable and is unlikely to work in the majority of cases.
'@example `notepadWindow.sendKeysMessage("^a")`
Public Sub SendKeysMessage(ByVal sKeys As String, Optional ByVal keyDelay As Long = 30, Optional ByVal bAutoRelease As Boolean = True)
  Const WM_KEYDOWN = &H100
  Const WM_KEYUP = &H101
        
  Dim keys() As KeyToken: keys = TokeniseKeys(sKeys, bAutoRelease)
  If Exists Then
    Call setThreadInput(True)
      Dim iKey As Long
      For iKey = 1 To UBound(keys)
        Dim key As KeyToken: key = keys(iKey)
        
        'Loop as many times as required times
        For n = 1 To key.iTimes
          'Key down event
          If key.iKeyState = tap Or key.iKeyState = down Then
            'Assume LParam = 0. This is rarely the case but it's not usually required anyway
            Call PostMessageA(This.Handle, WM_KEYDOWN, key.wVirtualKey, 0&)
            Call Sleep(keyDelay)
          End If
          
          'Key up event
          If key.iKeyState = tap Or key.iKeyState = up Then
            'Assume LParam = 0. This is rarely the case but it's not usually required anyway
            Call PostMessageA(This.Handle, WM_KEYUP, key.wVirtualKey, 0&)
            Call Sleep(keyDelay)
          End If
        Next
      Next
    Call setThreadInput(False)
  Else
    Err.Raise 1, "SendKeysMessage", "Window does not exist."
  End If
End Sub

'Await for the existence of an window which meets a specified condition.
'@param awaitFor as stdICallable<(stdWindow)=>Boolean> - If this callback returns true, the function will immediately set the result of the function to the awaited window, otherwise will continue awaiting.
'@param error as stdICallable<(stdWindow)=>Boolean> - If this callback returns true, the function will immediately set the result of the function to `nothing`, otherwise will continue awaiting. If none is provided no error is checked.
'@param timeout - Timeout in seconds. Defaults to -1 = NoTimeout. When timeout is reached the return value of the function is `nothing`
'@returns - Window awaited for
'@remark calls `Me.FindFirst(awaitFor)` to wait for element chosen.
Public Function AwaitForWindow(awaitFor As stdICallable, Optional error As stdICallable = Nothing, Optional timeout As Long = -1) As stdWindow
  Dim bCond As Boolean: bCond = True
  Dim el As Object: Set el = Nothing
  Dim oldTime As Date: oldTime = Now()
  While bCond
      Set el = FindFirst(awaitFor)
      If Not el Is Nothing Then bCond = False
      If Not error Is Nothing Then If error.Run(Me) Then bCond = False
      If timeout <> -1 And Second(Now() - oldTime) > timeout Then bCond = False
      DoEvents
  Wend
  Set AwaitForWindow = el
End Function

'================================================================================================
'=      PROTECTED METHODS
'================================================================================================

'Obtain a shell window cache
'@protected
'@param hwnd - Handle of the window to test for in the cache. If this window is not in the cache, the cache will be updated.
'@returns - Shell window cache
#if VBA7 then
  Friend Function protGetShellWindowCache(ByVal hwnd as LongPtr) as object
#else
  Friend Function protGetShellWindowCache(ByVal hwnd as Long) as object
#end if
  Static cache As Object: If cache Is Nothing Then Set cache = CreateObject("Scripting.Dictionary")
  If Not cache.Exists(CStr(hwnd)) Then
    Static shell As Object: If shell Is Nothing Then Set shell = CreateObject("Shell.Application")

    Dim win
    For Each win In shell.windows
      If Not cache.Exists(CStr(win.hwnd)) Then
        cache.Add CStr(win.hwnd), win
      End If
    Next
  End If
  Set protGetShellWindowCache = cache
End Function

'Obtain the next window given a stack
'@protected
'@param stack - Stack of windows
'@param DFS - Whether to use Depth First Search or not
'@param Prev - Previous window
'@returns - Next window
Friend Function protGetNextDescendent(ByVal stack As collection, ByVal DFS As Boolean, ByVal Prev As stdWindow) As stdWindow
  If stack.Count > 0 Then
    'Get the next window, use popCol if we want to do Depth First Search, else use shiftCol
    Dim oNext As stdWindow
    If DFS Then
      Set oNext = PopCol(stack)
    Else
      Set oNext = ShiftCol(stack)
    End If

    'Add all children to stack
    Dim windows As collection: Set windows = oNext.children.Object
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
'@protected
'@returns Object<Dictionary<Dictionary<Dictionary<string|long>>>> - Lookups object
Friend Function protGetLookups() as Object
    If This.Lookups Is Nothing Then
        'Notes:
        'EWndShowStyle is not needed, use isVisible, isMinimized and isMaximised

        Set This.Lookups = CreateObject("Scripting.Dictionary")
        Set This.Lookups("EWndStyles") = CreateLookupDict(Array( _
          "WS_BORDER", &H800000, "WS_CAPTION", &HC00000, _
          "WS_CHILD", &H40000000, "WS_CHILDWINDOW", &H40000000, _
          "WS_CLIPCHILDREN", &H2000000, "WS_CLIPSIBLINGS", &H4000000, _
          "WS_DISABLED", &H8000000, "WS_DLGFRAME", &H400000, _
          "WS_GROUP", &H20000, "WS_HSCROLL", &H100000, _
          "WS_ICONIC", &H20000000, "WS_MAXIMIZE", &H1000000, _
          "WS_MAXIMIZEBOX", &H10000, "WS_MINIMIZE", &H20000000, _
          "WS_MINIMIZEBOX", &H20000, "WS_OVERLAPPED", &H0, _
          "WS_POPUP", &H80000000, "WS_SIZEBOX", &H40000, _
          "WS_SYSMENU", &H80000, "WS_TABSTOP", &H10000, _
          "WS_THICKFRAME", &H40000, "WS_TILED", &H0, _
          "WS_VISIBLE", &H10000000, "WS_VSCROLL", &H200000, _
          "WS_OVERLAPPEDWINDOW", WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX, _
          "WS_POPUPWINDOW", WS_POPUP Or WS_BORDER Or WS_SYSMENU _
        ))
        Set This.Lookups("EWndExStyles") = CreateLookupDict(Array( _
          "WS_EX_ACCEPTFILES", &H10, "WS_EX_APPWINDOW", &H40000, _
          "WS_EX_CLIENTEDGE", &H200, "WS_EX_COMPOSITED", &H2000000, _
          "WS_EX_CONTEXTHELP", &H400, "WS_EX_CONTROLPARENT", &H10000, _
          "WS_EX_DLGMODALFRAME", &H1, "WS_EX_LAYERED", &H80000, _
          "WS_EX_LAYOUTRTL", &H400000, "WS_EX_LEFT", &H0, _
          "WS_EX_LEFTSCROLLBAR", &H4000, "WS_EX_LTRREADING", &H0, _
          "WS_EX_MDICHILD", &H40, "WS_EX_NOACTIVATE", &H8000000, _
          "WS_EX_NOINHERITLAYOUT", &H100000, "WS_EX_NOPARENTNOTIFY", &H4, _
          "WS_EX_NOREDIRECTIONBITMAP", &H200000, "WS_EX_RIGHT", &H1000, _
          "WS_EX_RIGHTSCROLLBAR", &H0, "WS_EX_RTLREADING", &H2000, _
          "WS_EX_STATICEDGE", &H20000, "WS_EX_TOOLWINDOW", &H80, _
          "WS_EX_TOPMOST", &H8, "WS_EX_TRANSPARENT", &H20, _
          "WS_EX_WINDOWEDGE", &H100, _
          "WS_EX_OVERLAPPEDWINDOW", WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE, _
          "WS_EX_PALETTEWINDOW", WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST _
        ))
        Set This.Lookups("EWndFindResult") = CreateLookupDict(Array( _
          "MatchFound", EWndFindResult.matchFound, _
          "MatchFoundSearchDescendents", EWndFindResult.MatchFoundSearchDescendents, _
          "NoMatchFound", EWndFindResult.NoMatchFound, _
          "NoMatchCancelSearch", EWndFindResult.NoMatchCancelSearch, _
          "NoMatchSkipDescendents", EWndFindResult.NoMatchSkipDescendents _
        ))
        Set This.Lookups("EWndState") = CreateLookupDict(Array( _
          "Normal", EWndState.Normal, _
          "Maximised", EWndState.Maximised, _
          "Minimised", EWndState.Minimised _
        ))
    End If
    Set protGetLookups = This.Lookups
End Function


'================================================================================================
'=      PRIVATE HELPERS
'================================================================================================
'Pop a value out of the end of a collection
'@param col as Collection<stdWindow> - Collection to pop value out of
'@returns - Window popped out of collection
Private Function PopCol(ByRef col As collection) As stdWindow
  Set PopCol = col(col.Count)
  Call col.remove(col.Count)
End Function

'Shift a value out of a collection
'@param col as Collection<stdWindow> - Collection to shift value out of
'@returns - Window shifted out of collection
Private Function ShiftCol(ByRef col As collection) As stdWindow
  Set ShiftCol = col(1)
  Call col.remove(1)
End Function

'Obtain a DWord from the high and low parts
'@param wHi - High word
'@param wLo - Low word
'@returns - Long dword returned as the combination of high and low word
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
'@param query as stdICallable - A callable with which standard globals will be bound to
Private Sub BindGlobals(ByRef query As stdICallable)
    Dim bSuccess As Boolean
    Dim keys: keys = This.Lookups.keys()
    For Each key In keys
      Call query.SendMessage("bindGlobal", bSuccess, Array(key, This.Lookups(key)("S2N")))
    Next
End Sub

'Constructor for a tFindNode type
'@param depth - depth of window
'@param element - window or element to store
'@returns {tFindNode} FindNode struct
Private Function CreateFindNode(ByVal depth As Long, ByVal element As stdWindow) As tFindNode
    CreateFindNode.initialised = True
    CreateFindNode.depth = depth
    Set CreateFindNode.element = element
End Function

'Given an array `stack`, and length `index`, append `item` onto the stack
'@param stack as Array<tFindNode> - stack of windows
'@param index - current index in stack
'@param item - Item to insert into stack
Private Sub stackPushV(ByRef stack() As tFindNode, ByRef index As Long, ByRef Item As tFindNode)
    Dim ub As Long: ub = UBound(stack)
    Dim size As Long: size = ub + 1
    If index > ub Then
        ReDim Preserve stack(0 To size * 2)
    End If
    stack(index).initialised = Item.initialised
    stack(index).depth = Item.depth
    Set stack(index).element = Item.element
    index = index + 1
End Sub

'Given an array `stack`, and `index`=>`n`, pop the nth item off the stack. Used in DFS
'@param stack as Array<tFindNode> - stack of windows
'@param index - current index in stack
'@returns - FindNode found at index
Private Function stackPopV(ByRef stack() As tFindNode, ByRef index As Variant) As tFindNode
    Dim size As Long: size = UBound(stack) + 1
    If index < size / 3 And index > minStackSize Then
        ReDim Preserve stack(0 To CLng(size / 2))
    End If
    index = index - 1
    
    stackPopV.initialised = stack(index).initialised
    stackPopV.depth = stack(index).depth
    Set stackPopV.element = stack(index).element
    stack(index).initialised = False
    stack(index).depth = 0
    Set stack(index).element = Nothing
End Function

'From an array of Key-Value pairs, create a dictionary of Key-->Value and Value-->Key lookups
'@param arr as Array<Variant> - 1D array of Key,Value pairs
'@returns Object<Dictionary<S2N: Dictionary<string,variant>, N2S: Dictionary<variant,string>>> - Created dictionary
Private Function CreateLookupDict(arr As Variant) As Object
    Dim oRet As Object
    Set oRet = CreateObject("Scripting.Dictionary")
    oRet.CompareMode = 1
    Dim i As Long

    'Create string --> number dictionary
    Set oRet("S2N") = CreateObject("Scripting.Dictionary")
    oRet("S2N").CompareMode = 1
    For i = LBound(arr) To UBound(arr) Step 2
        Call oRet("S2N").Add(arr(i), arr(i + 1))
    Next

    'Create number --> string dictionary
    Set oRet("N2S") = CreateObject("Scripting.Dictionary")
    oRet("N2S").CompareMode = 1
    For i = LBound(arr) To UBound(arr) Step 2
      If isObject(arr(i)) Then
        Set oRet("N2S")(arr(i + 1)) = arr(i)
      Else
        Let oRet("N2S")(arr(i + 1)) = arr(i)
      End If
    Next

    'Return dictionary
    Set CreateLookupDict = oRet
End Function

'Create an atom class string representing an atom.
'@param iAtom - Atom to convert to a string
'@returns - Classname representing atom
Private Function MakeIntAtom(ByVal iAtom As Long) As String
  MakeIntAtom = "#" & iAtom
End Function

'Sets the thread input state of the window.
'@protected
'@param bAttach - Whether to attach or detach the thread
'@remark This is used to synchronise the thread of the window with the thread of VBA. Important for `SendInput`, Hook procedures as well as usage in drag/drop and DDE.
Public Sub setThreadInput(ByVal bAttach As Boolean)
    Dim dwExcelTID As Long: dwExcelTID = GetCurrentThreadId()
    Dim dwThisTID As Long: dwThisTID = GetWindowThreadProcessId(This.Handle, 0)
    
    'Synchonise Excel's thread with this window's thread
    if Exists then
      if dwExcelTID <> dwThisTID then
        If AttachThreadInput(dwExcelTID, dwThisTID, iif(bAttach, 1, 0)) = 0 Then
            Err.Raise Err.LastDllError
        End If
      end if
    end if
End Sub

'Parses a set of keys and converts them into a KeyToken array
'@param sKeys - Key string to tokenise
'@returns Array<KeyToken> - Array of key tokens containing all data required for SendKeysMessage / SendKeysEvent / SendKeysInput
'@example `stdWindow.TokeniseKeys("^a{home up 3}{LControl up}")`
Private Function TokeniseKeys(ByVal sKeys As String, Optional ByVal bAutoRelease As Boolean = True) As KeyToken()
  Const MAPVK_VK_TO_VSC As Long = 0
  Dim iKeyCount As Long: iKeyCount = 0
  Dim ks() As KeyToken
  Dim sKeyExpr As String: sKeyExpr = sKeys
  Dim bModifiers(1 To 4) As Boolean
  Const MODIFIER_CONTROL = 1
  Const MODIFIER_ALT = 2
  Const MODIFIER_WIN = 3
  Const MODIFIER_SHIFT = 4
  
  'Simple keyboard expression parser
  While sKeyExpr <> ""
    'Catch infinite loop
    Dim sPrevExpr As String
    If sPrevExpr = sKeyExpr Then Err.Raise 1, "TokeniseKeys", "Expression same as previous, catching infinite loop."
    sPrevExpr = sKeyExpr
    
    'Parse token
    Dim sChar As String: sChar = Mid(sKeyExpr, 1, 1)
    Dim iVKey As EVirtualKey, iTimes As Long, iState As KeyState, iNextChar As Long, bIsUppercase As Boolean
    Select Case sChar
      Case "^"
        iVKey = VK_LCONTROL
        iTimes = 1
        iState = down
        iNextChar = 2
        bModifiers(MODIFIER_CONTROL) = True
        bIsUppercase = False
      Case "+"
        iVKey = VK_LSHIFT
        iTimes = 1
        iState = down
        iNextChar = 2
        bModifiers(MODIFIER_SHIFT) = True
        bIsUppercase = False
      Case "#"
        iVKey = VK_LWIN
        iTimes = 1
        iState = down
        iNextChar = 2
        bModifiers(MODIFIER_WIN) = True
        bIsUppercase = False
      Case "%", "!"
        iVKey = VK_LALT
        iTimes = 1
        iState = down
        iNextChar = 2
        bModifiers(MODIFIER_ALT) = True
        bIsUppercase = False
      Case " "
        iVKey = VK_SPACE
        iTimes = 1
        iState = tap
        iNextChar = 2
        bIsUppercase = False
      Case "\"
        iVKey = VKeyFromString(Mid(sKeyExpr, 2, 1))
        iTimes = 1
        iState = tap
        iNextChar = 3
        bIsUppercase = False
        
      Case "{"
        'Obtain expression to turn to key code
        Dim iExpressionEnd As Long: iExpressionEnd = InStr(1, sKeyExpr, "}")
        Dim sExpression As String: sExpression = Mid(sKeyExpr, 2, iExpressionEnd - 2)
        While InStr(1, sExpression, "  ") > 0: sExpression = Replace(sExpression, "  ", " "): Wend
        Dim vExpression: vExpression = Split(sExpression, " ")
        Dim iExpressionLen As Long: iExpressionLen = UBound(vExpression) - LBound(vExpression) + 1
        
        '1st arg is always the vkey
        Dim sSpecialKey As String: sSpecialKey = vExpression(0)
        iVKey = VKeyFromString(sSpecialKey)
        iState = tap 'Default
        iTimes = 1   'Default
        
        'Override defaults based on other arguments parsed
        'Arguments can either be UP, DOWN, TAP (key state), or a number (times key is pressed)
        Dim iArg As Long
        For iArg = 1 To UBound(vExpression)
          vArg = UCase(vExpression(iArg))
          Select Case vArg
            Case "UP"
              iState = up
            Case "DOWN"
              iState = down
            Case "TAP"
              iState = tap
            Case Else
            If IsNumeric(vArg) Then
              iTimes = CLng(vArg)
            Else
              Err.Raise 1, "TokeniseKeys", "No arg of type """ & vArg & """ allowed."
            End If
          End Select
        Next
        
        iNextChar = iExpressionEnd + 1
        bIsUppercase = False
      Case Else
        iVKey = VKeyFromString(sChar)
        iTimes = 1
        iState = tap
        iNextChar = 2
        'Note: if only havine one of the checks, characters like `5` are signalled as upper case when they aren't.
        bIsUppercase = sChar = UCase(sChar) And sChar <> LCase(sChar)
    End Select
    
    'Include shift keys for upper case letters
    If bIsUppercase Then
        ReDim Preserve ks(1 To iKeyCount + 3)
        
        With ks(iKeyCount + 1)
            .wVirtualKey = VK_LSHIFT
            .wScanCode = MapVirtualKeyA(VK_LSHIFT, MAPVK_VK_TO_VSC)
            .iKeyState = down
            .iTimes = 1
        End With
        With ks(iKeyCount + 2)
          .wVirtualKey = iVKey
          .wScanCode = MapVirtualKeyA(iVKey, MAPVK_VK_TO_VSC)
          .iKeyState = iState
          .iTimes = iTimes
        End With
        With ks(iKeyCount + 3)
            .wVirtualKey = VK_LSHIFT
            .wScanCode = MapVirtualKeyA(VK_LSHIFT, MAPVK_VK_TO_VSC)
            .iKeyState = up
            .iTimes = 1
        End With
        
        iKeyCount = iKeyCount + 3
    Else
        'Add key directly
        iKeyCount = iKeyCount + 1
        ReDim Preserve ks(1 To iKeyCount)
        With ks(iKeyCount)
          .wVirtualKey = iVKey
          .wScanCode = MapVirtualKeyA(iVKey, MAPVK_VK_TO_VSC)
          .iKeyState = iState
          .iTimes = iTimes
        End With
    End If
    
    'Trim expression
    sKeyExpr = Mid(sKeyExpr, iNextChar)
  Wend
  
  'Make sure all key down events are released
  If bAutoRelease Then
    Dim iStartKeyCount As Long: iStartKeyCount = iKeyCount
    Dim iModifier As Long, jModifier As Long
    For iModifier = 1 To iStartKeyCount
      If ks(iModifier).iKeyState = down Then
        'Check whether the key is down
        Dim isDown As Boolean: isDown = True
        For jModifier = iKeyCount To iModifier Step -1
          If ks(jModifier).wVirtualKey = ks(iModifier).wVirtualKey Then
            isDown = ks(jModifier).iKeyState = down
            Exit For
          End If
        Next
        
        'If the key is down, then add an additional key onto the end of the array releasing it
        If isDown Then
          iKeyCount = iKeyCount + 1
          ReDim Preserve ks(1 To iKeyCount)
          ks(iKeyCount).wVirtualKey = ks(iModifier).wVirtualKey
          ks(iKeyCount).wScanCode = ks(iModifier).wScanCode
          ks(iKeyCount).iTimes = 1
          ks(iKeyCount).iKeyState = up
        End If
      End If
    Next
  End If
  
  'return token list
  TokeniseKeys = ks
End Function

'Obtain VKey code from string
'@param s - KeyName as string
'@returns - VKey code
Private Function VKeyFromString(ByRef s As String) As EVirtualKey
  Select Case UCase(s)
    Case "LBUTTON": VKeyFromString = EVirtualKey.VK_LBUTTON
    Case "RBUTTON": VKeyFromString = EVirtualKey.VK_RBUTTON
    Case "CANCEL": VKeyFromString = EVirtualKey.VK_CANCEL
    Case "MBUTTON": VKeyFromString = EVirtualKey.VK_MBUTTON
    Case "XBUTTON1": VKeyFromString = EVirtualKey.VK_XBUTTON1
    Case "XBUTTON2": VKeyFromString = EVirtualKey.VK_XBUTTON2
    Case "BACK": VKeyFromString = EVirtualKey.VK_BACK
    Case "TAB": VKeyFromString = EVirtualKey.VK_TAB
    Case "CLEAR": VKeyFromString = EVirtualKey.VK_CLEAR
    Case "RETURN": VKeyFromString = EVirtualKey.VK_RETURN
    Case "ENTER": VKeyFromString = EVirtualKey.VK_RETURN
    Case "SHIFT": VKeyFromString = EVirtualKey.VK_SHIFT
    Case "CONTROL": VKeyFromString = EVirtualKey.VK_CONTROL
    Case "ALT": VKeyFromString = EVirtualKey.VK_ALT
    Case "PAUSE": VKeyFromString = EVirtualKey.VK_PAUSE
    Case "CAPITAL": VKeyFromString = EVirtualKey.VK_CAPITAL
    Case "KANA": VKeyFromString = EVirtualKey.VK_KANA
    Case "HANGUEL": VKeyFromString = EVirtualKey.VK_HANGUEL
    Case "HANGUL": VKeyFromString = EVirtualKey.VK_HANGUL
    Case "IME_ON": VKeyFromString = EVirtualKey.VK_IME_ON
    Case "JUNJA": VKeyFromString = EVirtualKey.VK_JUNJA
    Case "FINAL": VKeyFromString = EVirtualKey.VK_FINAL
    Case "HANJA": VKeyFromString = EVirtualKey.VK_HANJA
    Case "KANJI": VKeyFromString = EVirtualKey.VK_KANJI
    Case "IME_OFF": VKeyFromString = EVirtualKey.VK_IME_OFF
    Case "ESCAPE": VKeyFromString = EVirtualKey.VK_ESCAPE
    Case "CONVERT": VKeyFromString = EVirtualKey.VK_CONVERT
    Case "NONCONVERT": VKeyFromString = EVirtualKey.VK_NONCONVERT
    Case "ACCEPT": VKeyFromString = EVirtualKey.VK_ACCEPT
    Case "MODECHANGE": VKeyFromString = EVirtualKey.VK_MODECHANGE
    Case "SPACE": VKeyFromString = EVirtualKey.VK_SPACE
    Case "PRIOR": VKeyFromString = EVirtualKey.VK_PRIOR
    Case "NEXT": VKeyFromString = EVirtualKey.VK_NEXT
    Case "END": VKeyFromString = EVirtualKey.VK_END
    Case "HOME": VKeyFromString = EVirtualKey.VK_HOME
    Case "LEFT": VKeyFromString = EVirtualKey.VK_LEFT
    Case "UP": VKeyFromString = EVirtualKey.VK_UP
    Case "RIGHT": VKeyFromString = EVirtualKey.VK_RIGHT
    Case "DOWN": VKeyFromString = EVirtualKey.VK_DOWN
    Case "SELECT": VKeyFromString = EVirtualKey.VK_SELECT
    Case "PRINT": VKeyFromString = EVirtualKey.VK_PRINT
    Case "EXECUTE": VKeyFromString = EVirtualKey.VK_EXECUTE
    Case "SNAPSHOT": VKeyFromString = EVirtualKey.VK_SNAPSHOT
    Case "INSERT": VKeyFromString = EVirtualKey.VK_INSERT
    Case "DELETE": VKeyFromString = EVirtualKey.VK_DELETE
    Case "HELP": VKeyFromString = EVirtualKey.VK_HELP
    Case "0": VKeyFromString = EVirtualKey.VK_0
    Case "1": VKeyFromString = EVirtualKey.VK_1
    Case "2": VKeyFromString = EVirtualKey.VK_2
    Case "3": VKeyFromString = EVirtualKey.VK_3
    Case "4": VKeyFromString = EVirtualKey.VK_4
    Case "5": VKeyFromString = EVirtualKey.VK_5
    Case "6": VKeyFromString = EVirtualKey.VK_6
    Case "7": VKeyFromString = EVirtualKey.VK_7
    Case "8": VKeyFromString = EVirtualKey.VK_8
    Case "9": VKeyFromString = EVirtualKey.VK_9
    Case "A": VKeyFromString = EVirtualKey.VK_A
    Case "B": VKeyFromString = EVirtualKey.VK_B
    Case "C": VKeyFromString = EVirtualKey.VK_C
    Case "D": VKeyFromString = EVirtualKey.VK_D
    Case "E": VKeyFromString = EVirtualKey.VK_E
    Case "F": VKeyFromString = EVirtualKey.VK_F
    Case "G": VKeyFromString = EVirtualKey.VK_G
    Case "H": VKeyFromString = EVirtualKey.VK_H
    Case "I": VKeyFromString = EVirtualKey.VK_I
    Case "J": VKeyFromString = EVirtualKey.VK_J
    Case "K": VKeyFromString = EVirtualKey.VK_K
    Case "L": VKeyFromString = EVirtualKey.VK_L
    Case "M": VKeyFromString = EVirtualKey.VK_M
    Case "N": VKeyFromString = EVirtualKey.VK_N
    Case "O": VKeyFromString = EVirtualKey.VK_O
    Case "P": VKeyFromString = EVirtualKey.VK_P
    Case "Q": VKeyFromString = EVirtualKey.VK_Q
    Case "R": VKeyFromString = EVirtualKey.VK_R
    Case "S": VKeyFromString = EVirtualKey.VK_S
    Case "T": VKeyFromString = EVirtualKey.VK_T
    Case "U": VKeyFromString = EVirtualKey.VK_U
    Case "V": VKeyFromString = EVirtualKey.VK_V
    Case "W": VKeyFromString = EVirtualKey.VK_W
    Case "X": VKeyFromString = EVirtualKey.VK_X
    Case "Y": VKeyFromString = EVirtualKey.VK_Y
    Case "Z": VKeyFromString = EVirtualKey.VK_Z
    Case "LWIN": VKeyFromString = EVirtualKey.VK_LWIN
    Case "RWIN": VKeyFromString = EVirtualKey.VK_RWIN
    Case "APPS": VKeyFromString = EVirtualKey.VK_APPS
    Case "SLEEP": VKeyFromString = EVirtualKey.VK_SLEEP
    Case "NUMPAD0": VKeyFromString = EVirtualKey.VK_NUMPAD0
    Case "NUMPAD1": VKeyFromString = EVirtualKey.VK_NUMPAD1
    Case "NUMPAD2": VKeyFromString = EVirtualKey.VK_NUMPAD2
    Case "NUMPAD3": VKeyFromString = EVirtualKey.VK_NUMPAD3
    Case "NUMPAD4": VKeyFromString = EVirtualKey.VK_NUMPAD4
    Case "NUMPAD5": VKeyFromString = EVirtualKey.VK_NUMPAD5
    Case "NUMPAD6": VKeyFromString = EVirtualKey.VK_NUMPAD6
    Case "NUMPAD7": VKeyFromString = EVirtualKey.VK_NUMPAD7
    Case "NUMPAD8": VKeyFromString = EVirtualKey.VK_NUMPAD8
    Case "NUMPAD9": VKeyFromString = EVirtualKey.VK_NUMPAD9
    Case "MULTIPLY": VKeyFromString = EVirtualKey.VK_MULTIPLY
    Case "ADD": VKeyFromString = EVirtualKey.VK_ADD
    Case "SEPARATOR": VKeyFromString = EVirtualKey.VK_SEPARATOR
    Case "SUBTRACT": VKeyFromString = EVirtualKey.VK_SUBTRACT
    Case "DECIMAL": VKeyFromString = EVirtualKey.VK_DECIMAL
    Case "DIVIDE": VKeyFromString = EVirtualKey.VK_DIVIDE
    Case "F1": VKeyFromString = EVirtualKey.VK_F1
    Case "F2": VKeyFromString = EVirtualKey.VK_F2
    Case "F3": VKeyFromString = EVirtualKey.VK_F3
    Case "F4": VKeyFromString = EVirtualKey.VK_F4
    Case "F5": VKeyFromString = EVirtualKey.VK_F5
    Case "F6": VKeyFromString = EVirtualKey.VK_F6
    Case "F7": VKeyFromString = EVirtualKey.VK_F7
    Case "F8": VKeyFromString = EVirtualKey.VK_F8
    Case "F9": VKeyFromString = EVirtualKey.VK_F9
    Case "F10": VKeyFromString = EVirtualKey.VK_F10
    Case "F11": VKeyFromString = EVirtualKey.VK_F11
    Case "F12": VKeyFromString = EVirtualKey.VK_F12
    Case "F13": VKeyFromString = EVirtualKey.VK_F13
    Case "F14": VKeyFromString = EVirtualKey.VK_F14
    Case "F15": VKeyFromString = EVirtualKey.VK_F15
    Case "F16": VKeyFromString = EVirtualKey.VK_F16
    Case "F17": VKeyFromString = EVirtualKey.VK_F17
    Case "F18": VKeyFromString = EVirtualKey.VK_F18
    Case "F19": VKeyFromString = EVirtualKey.VK_F19
    Case "F20": VKeyFromString = EVirtualKey.VK_F20
    Case "F21": VKeyFromString = EVirtualKey.VK_F21
    Case "F22": VKeyFromString = EVirtualKey.VK_F22
    Case "F23": VKeyFromString = EVirtualKey.VK_F23
    Case "F24": VKeyFromString = EVirtualKey.VK_F24
    Case "NUMLOCK": VKeyFromString = EVirtualKey.VK_NUMLOCK
    Case "SCROLL": VKeyFromString = EVirtualKey.VK_SCROLL
    Case "LSHIFT": VKeyFromString = EVirtualKey.VK_LSHIFT
    Case "RSHIFT": VKeyFromString = EVirtualKey.VK_RSHIFT
    Case "LCONTROL": VKeyFromString = EVirtualKey.VK_LCONTROL
    Case "RCONTROL": VKeyFromString = EVirtualKey.VK_RCONTROL
    Case "LALT": VKeyFromString = EVirtualKey.VK_LALT
    Case "RALT": VKeyFromString = EVirtualKey.VK_RALT
    Case "/": VKeyFromString = EVirtualKey.VK_OEM_2
    Case ";": VKeyFromString = EVirtualKey.VK_OEM_1
    Case ",": VKeyFromString = EVirtualKey.VK_OEM_COMMA
    Case Else
      Err.Raise 1, "VKeyFromString()", "No key """ & s & """"
  End Select
End Function

'Create an lParam key from params
'@param repeatCount - Number of times to repeat key
'@param scanCode - Scan code of key
'@param extended - Whether the key is extended
'@param bDown - Whether the key is down
'@returns - key as an lParam
Private Function createKeyLParam(ByVal repeatCount As Long, ByVal scanCode As Long, ByVal extended As Boolean, ByVal bDown As Boolean) As Long
  createKeyLParam = repeatCount Or lshift(scanCode, 16) Or lshift(extended, 24) Or lshift(iif(bDown, 0, 1), 29) Or lshift(iif(bDown, 0, 1), 30) Or lshift(iif(bDown, 0, 1), 31)
End Function

'Left shift a value by n times
'@param val - Value to shift
'@param nTimes - Number of times to shift
'@returns - Shifted value
Private Function lshift(ByVal val As Long, ByVal nTimes As Integer) As Long
  lshift = val * 2 ^ nTimes
End Function

'Check whether a bit flag is set in a bit field
'@param field - The field to check whether the flag is set in
'@param flag - The flag to check for
'@returns - True if flag is present, false otherwise.
Private Function bitFlagCheck(ByVal field As Long, ByVal flag As Long) As Boolean
  bitFlagCheck = (field And flag) = flag
End Function

'Set a bit flag in a bit field.
'@param field - The field to set the flag within
'@param flag - The flag to set
'@param toSet - True to ensure the flag is set. False to ensure the flag is not set.
'@returns - field with added / ensured bit flag.
Private Function bitFlagSet(ByVal field As Long, ByVal flag As Long, ByVal toSet As Boolean) As Long
  If toSet Then
    bitFlagSet = field Or flag
  Else
    bitFlagSet = field And (Not flag)
  End If
End Function

'Public Sub Requires()
'  if IsEmpty(stdEnumerator) then Call Msgbox("Requires `stdEnumerator`", vbInformation) else stdEnumerator.Requires
'
'End Sub

