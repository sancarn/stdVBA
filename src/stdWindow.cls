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
Private Type KeyboardInput
  iType As Long    'DWORD = INPUT_KEYBOARD
  wVk As Integer   'WORD
  wScan As Integer 'WORD
  dwFlags As Long  'DWORD
  time As Long     'DWORD
  #If VBA7 Then    'ULONG_PTR
    dwExtraInfo As LongPtr
    bPadding(1 To 12) As Byte '   12 extra bytes, because mouses take more.
  #Else
    dwExtraInfo As Long
    bPadding(1 To 8) As Byte '   8 extra bytes, because mouses take more.
  #End If
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

#If VBA7 Then
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

  'Automating windows
  Private Declare PtrSafe Function SendMessageA Lib "user32" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As Any, ByVal lParam As Any) As Long
  Private Declare PtrSafe Function PostMessageA Lib "user32" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As Any, ByVal lParam As Any) As Long
  Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As EWndShowStyle) As Long
  Private Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hwnd As LongPtr) As Long
  Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hwnd As LongPtr) As Long

  'SendKeys
  Private Declare PtrSafe Function MapVirtualKeyA Lib "user32" (ByVal uCode As Long, ByVal uMapType As Long) As Long
  Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVK As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As LongPtr)
  Private Declare PtrSafe Function SendInput Lib "user32" (ByVal cInputs As Long, ByRef pInput As KeyboardInput, ByVal cbSize As Long) As Long
  Private Declare PtrSafe Function GetMessageExtraInfo Lib "user32" () As LongPtr
  Private Declare PtrSafe Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
#Else
  'Constructors
  Private Declare Function GetDesktopWindow Lib "user32" () As Long
  Private Declare Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByVal hwnd As Long) As Long
  Private Declare Function WindowFromPoint Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

  'Getting window data
  Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare Function IsHungAppWindow Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
  Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
  Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, ByRef pRect As apiRect) As Long
  Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, ByRef pRect As apiRect) As Long
  Private Declare Function GetWindowInfo Lib "user32" (ByVal hwnd As Long, ByRef pInf As apiWindowInfo) As Long
  Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare Function SetParent Lib "user32" (ByVal hwnd As Long, ByVal hWndParent As Long) As Long
  Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare Function GetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As apiWindowLongType) As Long
  Private Declare Function GetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As apiWindowLongType) As Long
  Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As apiWindowLongType, ByVal dwNewLong As Long) As Long
  Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As apiWindowLongType, ByVal dwNewPtr As Long) As Long
  Private Declare Function GetAncestor Lib "user32" (ByVal hwnd As Long, ByVal nIndex As apiWindowAncestorType) As Long
  Private Declare Function FindWindowExA Lib "user32" (ByVal hwnd As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
  Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hwndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal flags As Long) As Long
  Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal bRepaint As Boolean) As Long
  Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
  
  'Redrawing window, UpdateWindow can also be used but isn't as safe...
  Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal flags As Long) As Long

  'Get children / siblings / parent
  Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As apiWindowRelationship) As Long

  'Get process related data
  Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, ByRef ldpwProcessId As Long) As Long
  Private Declare Function GetWindowModuleFileName Lib "user32" Alias "GetWindowModuleFileNameA" (ByVal hwnd As Long, ByVal pszFileName As String, ByVal cchFileNameMax As Long) As Long
  Private Declare Function GetCurrentThreadId Lib "Kernel32" Alias "GetCurrentThreadID" () As Long
  Private Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long

  'Setting window data
  Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long

  'Automating windows
  Private Declare Function SendMessageA Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Any, ByVal lParam As Any) As Long
  Private Declare Function PostMessageA Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Any, ByVal lParam As Any) As Long
  Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As EWndShowStyle) As Long
  Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

  'SendKeys
  Private Declare Function MapVirtualKeyA Lib "user32" (ByVal uCode As Long, ByVal uMapType As Long) As Long
  Private Declare Sub keybd_event Lib "user32" (ByVal bVK As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
  Private Declare Function SendInput Lib "user32" (ByVal cInputs As Long, ByRef pInputs As KeyboardInput, ByVal cbSize As Long) As Long
  Private Declare Function GetMessageExtraInfo Lib "user32" () As Long
  Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
#End If


#If VBA7 Then
  Private pHandle As LongPtr
#Else
  Private pHandle As Long
#End If

Private pInitialized As Boolean
Private Lookups As Object


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
#If VBA7 Then
Private Function Create(ByVal sClassName As String, ByVal sCaption As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, ByVal lpParam As Long) As stdWindow
#Else
Private Function Create(ByVal sClassName As String, ByVal sCaption As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByVal lpParam As Long) As stdWindow
#End If
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

'Create a window object from a window handle
'@constructor
'@param {ByVal LongPtr} Handle to window (hwnd) to create window object for
'@returns {stdWindow} Specificed window.
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
'@returns {stdWindow}
Public Function CreateFromApplication() As stdWindow
    select case Application.Name
        case "Microsoft Excel"
            Set CreateFromApplication = CreateFromHwnd(Application.hwnd)
        case "Microsoft Word"
            Set CreateFromApplication = CreateFromHwnd(Application.ActiveWindow.Hwnd)
        case "Microsoft PowerPoint"
            Err.Raise 1, "stdAcc::CreateFromApplication()", "No built in way of finding powerpoint hwnd. We suggest using stdWindow to get hwnd, and create direct from hwnd instead " & Application.name
        case else
            Err.Raise 1, "stdAcc::CreateFromApplication()", "No implementation for getting application window of " & Application.name
    end select
End Function

'Create a window from the desktop window
'@constructor
'@returns {stdWindow} Desktop window
Public Function CreateFromDesktop() As stdWindow
  Set CreateFromDesktop = CreateFromHwnd(GetDesktopWindow())
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
  #If VBA7 Then
    Dim hwnd As LongPtr
  #Else
    Dim hwnd As Long
  #End If
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
Public Function CreateFromContextMenu() As stdWindow
  #If VBA7 Then
    Dim hwnd As LongPtr
  #Else
    Dim hwnd As Long
  #End If
  hwnd = FindWindowExA(Null, Null, MakeIntAtom(&H8000&), vbNullString)
  If hwnd <> 0 Then
    Set CreateFromContextMenu = CreateFromHwnd(hwnd)
  Else
    Set CreateFromContextMenu = Nothing
  End If
End Function

#if FULL_INTELLISENSE then
Public Function AsAcc() as stdAcc
#else
Public Function AsAcc() as Object
#end if
  if isObject(stdAcc) then
    set AsAcc = stdAcc.CreateFromHwnd(pHandle)
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
'@param {ByVal LongPtr} Handle to window (hwnd) to create window object for
'@returns {stdWindow} Specificed window.
#If VBA7 Then
Friend Sub protInit(ByVal hwnd As LongPtr)
#Else
Friend Sub protInit(ByVal hwnd As Long)
#End If
  pHandle = hwnd
  pInitialized = True
End Sub

'Whenever we initialise ensure to get lookups from `stdWindow`
Private Sub Class_Initialize()
    Set Lookups = stdWindow.protGetLookups()
End Sub


'--------------------------------------------------------------------------------
'Access window information
'--------------------------------------------------------------------------------

'Get the hWND / window ID of the window
#If VBA7 Then
Public Property Get handle() As LongPtr
#Else
Public Property Get handle() As Long
#End If
  handle = pHandle
End Property

'Get the handle to the display context for the window
#If VBA7 Then
Public Property Get hDC() As LongPtr
#Else
Public Property Get hDC() As Long
#End If
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
    sCharBuffer = space(256)
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
    sCharBuffer = space(256)
    iNumChars = GetClassName(pHandle, sCharBuffer, 256)
    Class = Mid(sCharBuffer, 1, iNumChars)
  Else
    Err.Raise 1, "Class", "Window does not exist."
  End If
End Property

'Get/Let visibility of window
Public Property Get Visible() As Boolean
  If Exists Then
    Visible = IsWindowVisible(pHandle)
  Else
    Err.Raise 1, "Visible", "Window does not exist."
  End If
End Property
Public Property Let Visible(ByVal x As Boolean)
  If Exists Then
    If x Then
      Call ShowWindow(pHandle, EWndShowStyle.SW_SHOW)
    Else
      Call ShowWindow(pHandle, EWndShowStyle.SW_HIDE)
    End If
  Else
    Err.Raise 1, "Visible", "Window does not exist."
  End If
End Property

'Get/Let windowState of window
Public Property Get State() As EWndState
  If Exists Then
    If IsZoomed(pHandle) Then
      State = EWndState.Maximised
    ElseIf IsIconic(pHandle) Then
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
          Call ShowWindow(pHandle, EWndShowStyle.SW_RESTORE)
        Case EWndState.Maximised
          Call ShowWindow(pHandle, EWndShowStyle.SW_MAXIMIZE)
        Case EWndState.Minimised
          Call ShowWindow(pHandle, EWndShowStyle.SW_FORCEMINIMIZE)
      End Select
    Else
      Err.Raise 1, "State", "Cannot set window state of a hidden window."
    End If
  Else
    Err.Raise 1, "State", "Window does not exist."
  End If
End Property

'Get the child windows of this window
Public Property Get children() As collection
  If Exists Then
    'Define collection to return
    Dim ret As collection
    Set ret = New collection

    'Attempt to get a child window
    #If VBA7 Then
      Dim childHandle As LongPtr
    #Else
      Dim childHandle As Long
    #End If
    childHandle = GetWindow(pHandle, GW_CHILD)

    'If a child window exists, find all children
    If childHandle <> 0 Then
      'Quickly add all hwnds to an array
      'This is a fast operation, object creation is anticipated to be a slower operation, thus is left till afterwards
      #If VBA7 Then
        Dim childHandles() As LongPtr
      #Else
        Dim childHandles() As Long
      #End If

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
  If isArray(rect) Then
    If Exists Then
      Call MoveWindow(pHandle, rect(0), rect(1), rect(2), rect(3), True)
    Else
      Err.Raise 1, "RectWindow [Let]", "Window does not exist."
    End If
  Else
    Err.Raise 1, "RectWindow [Let]", "Value must be an array of 4 doubles - left, top, width, height."
  End If
End Property
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
    Call MoveWindow(pHandle, vX, rect(1), rect(2), rect(3), True)
  Else
    Err.Raise 1, "stdWindow#x [Let]", "Window does not exist."
  End If
End Property

'Get/Set the Y position of this window
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
    Call MoveWindow(pHandle, rect(0), vY, rect(2), rect(3), True)
  Else
    Err.Raise 1, "stdWindow#y [Let]", "Window does not exist."
  End If
End Property

'Get/Set the width of this window
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
    Call MoveWindow(pHandle, rect(0), rect(1), vW, rect(3), True)
  Else
    Err.Raise 1, "stdWindow#width [Let]", "Window does not exist."
  End If
End Property

'Get/Set the height of this window
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
Public Property Get parent() As stdWindow
  #If VBA7 Then
    Dim pHwnd As LongPtr
  #Else
    Dim pHwnd As Long
  #End If
  pHwnd = GetParent(pHandle)
  If pHwnd <> 0 Then
    Set parent = stdWindow.CreateFromHwnd(GetParent(pHandle))
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
    #If VBA7 Then
      Dim hOldParent As LongPtr
    #Else
      Dim hOldParent As Long
    #End If
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

#If VBA7 Then
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
#Else
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
#End If
  



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

'Get/Set topmost
Public Property Get isTopmost() as Boolean
  const WS_EX_TOPMOST = &H8
  isTopMost = (StyleEx And WS_EX_TOPMOST) = WS_EX_TOPMOST
End Property
Public Property Let isTopmost(ByVal v as boolean)
  Const HWND_TOPMOST   = -1
  Const HWND_NOTOPMOST = -2
  Const SWP_NOSIZE     = &H1
  Const SWP_NOMOVE     = &H2
  Call SetWindowPos(pHandle, iif(v, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE or SWP_NOSIZE)
End Property

'================================================================================================
'=      PUBLIC INSTANCE METHODS
'================================================================================================


'Set hooks for a window
#If VBA7 Then
Public Function SetHook(ByVal idHook As apiWindowHookType, ByVal hook As LongPtr, hInstance As LongPtr, dwThreadID As Long) As LongPtr
#Else
Public Function SetHook(ByVal idHook As apiWindowHookType, ByVal hook As Long, hInstance As Long, dwThreadID As Long) As Long
#End If
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
'@param {Long=-1} Static depth limit. E.G. if you want to search children only, set this value to 1
'@examples
' ```
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
'@param {Long=-1} Static depth limit. E.G. if you want to search children only, set this value to 1
'@examples
' ```
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



'--------------------------------------------------------------------------------
'Automate the window
'--------------------------------------------------------------------------------

Public Sub Quit()
  If Exists Then
    Call PostMessage(&H10&,0,0)
  Else
    Err.Raise 1, "Close", "Window does not exist."
  End If
End Sub

Public Sub ForceQuit()
  If Exists Then
    If DestroyWindow(pHandle) = 0 Then
      Err.Raise Err.LastDllError, "Close", "Failed to close window"
    End If
  Else
    Err.Raise 1, "Close", "Window does not exist."
  End If
End Sub

Public Sub Activate()
  If Exists Then
    Call setThreadInput(True)
      Call SetForegroundWindow(pHandle)
    Call setThreadInput(False)
  Else
    Err.Raise 1, "Minimize", "Window does not exist."
  End If
End Sub

'Get all descendents of the stdWindow
'@returns {Collection<stdWindow>} Collection of descendents
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

Public Sub Redraw()
  Const RDW_INVALIDATE = &H1
  Call RedrawWindow(pHandle, 0&, 0&, RDW_INVALIDATE)
End Sub

Public Function SendMessage(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
  If Exists Then
    SendMessage = SendMessageA(pHandle, wMsg, wParam, lParam)
  Else
    Err.Raise 1, "SendMessage", "Window does not exist."
  End If
End Function
Public Sub PostMessage(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
  If Exists Then
    'If desktop then broadcasts
    #If VBA7 Then
      Dim hwnd As LongPtr
    #Else
      Dim hwnd As Long
    #End If
    hwnd = pHandle
    If pHandle = GetDesktopWindow Then hwnd = &HFFFF
    
    If PostMessageA(hwnd, wMsg, wParam, lParam) =  0 Then
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
'@remark Internally uses SendInput
Public Sub SendKeysInput(ByVal sKeys As String, Optional ByVal bRaw As Boolean = False, Optional ByVal keyDelay As Long = 0, Optional bAutoRelease As Boolean = True)
  Const INPUT_KEYBOARD As Long = 1
  Const KEYEVENTF_KEYUP = &H2
  Dim keys() As KeyToken: keys = TokeniseKeys(sKeys, bAutoRelease)
  If Exists Then
    Activate()
    
    'Loop over all keys
    Dim iKey As Long
    For iKey = 1 To UBound(keys)
        Dim key As KeyToken: key = keys(iKey)
        
        'Create generic key signal
        Dim inputKey As KeyboardInput
        inputKey.iType = INPUT_KEYBOARD
        inputKey.wVk = key.wVirtualKey
        inputKey.wScan = key.wScanCode
        inputKey.time = 0
        
        'Key down input
        If key.iKeyState = tap Or key.iKeyState = down Then
          inputKey.dwFlags = 0
          If SendInput(1, inputKey, LenB(inputKey)) = 0 Then
              Err.Raise Err.LastDllError, "SendKeysInput", "Input might be blocked by another thread (DLL Error: " & Err.LastDllError & ")"
          End If
          Call Sleep(keyDelay)
        End If
        
        'Key up input
        If key.iKeyState = tap Or key.iKeyState = up Then
          inputKey.dwFlags = KEYEVENTF_KEYUP
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
'@param {ByVal String} Keys to send to the window
'@param {Optional ByVal Long} Delay between each keystroke
'@param {Optional ByVal Boolean} Whether keys pressed down should be auto-released
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
'@param {ByVal String} Keys to send to the window
'@param {Optional ByVal Long} Delay between each keystroke
'@param {Optional ByVal Boolean} Whether keys pressed down should be auto-released
'@remark It should be noted that this method is extremely unstable and is unlikely to work in the majority of cases.
'@example `notepadWindow.sendKeysEvent("^a")`
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
            Call PostMessageA(pHandle, WM_KEYDOWN, key.wVirtualKey, 0&)
            Call Sleep(keyDelay)
          End If
          
          'Key up event
          If key.iKeyState = tap Or key.iKeyState = up Then
            'Assume LParam = 0. This is rarely the case but it's not usually required anyway
            Call PostMessageA(pHandle, WM_KEYUP, key.wVirtualKey, 0&)
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
'@param {stdICallable<stdWindow=>Boolean>} If this callback returns true, the function will immediately set the result of the function to the awaited window, otherwise will continue awaiting.
'@param {stdICallable<stdWindow=>Boolean> = Nothing} If this callback returns true, the function will immediately set the result of the function to `nothing`, oterwise will continue awaiting. If none is provided no error is checked.
'@param {Long=-1} Timeout in seconds. Defaults to -1 = NoTimeout. When timeout is reached the return value of the function is `nothing`
'@returns {stdWindow} Window awaited for
'@remark calls `Me.FindFirst(awaitFor)` to wait for element chosen
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

'Obtain the next window given a stack
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
'@returns {Dictionary<Dictionary<Dictionary<string|long>>>}
Friend Function protGetLookups()
    If Lookups Is Nothing Then
        'Notes:
        'EWndShowStyle is not needed, use isVisible, isMinimized and isMaximised

        Set Lookups = CreateObject("Scripting.Dictionary")
        Set Lookups("EWndStyles") = CreateLookupDict(Array( _
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
        Set Lookups("EWndExStyles") = CreateLookupDict(Array( _
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
        Set Lookups("EWndFindResult") = CreateLookupDict(Array( _
          "MatchFound", EWndFindResult.matchFound, _
          "MatchFoundSearchDescendents", EWndFindResult.MatchFoundSearchDescendents, _
          "NoMatchFound", EWndFindResult.NoMatchFound, _
          "NoMatchCancelSearch", EWndFindResult.NoMatchCancelSearch, _
          "NoMatchSkipDescendents", EWndFindResult.NoMatchSkipDescendents _
        ))
        Set Lookups("EWndState") = CreateLookupDict(Array( _
          "Normal", EWndState.Normal, _
          "Maximised", EWndState.Maximised, _
          "Minimised", EWndState.Minimised _
        ))
    End If
    Set protGetLookups = Lookups
End Function


'================================================================================================
'=      PRIVATE HELPERS
'================================================================================================
'Pop a value out of the end of a collection
'@param {ByRef Collection<stdWindow>} Collection to pop value out of
'@returns {stdWindow} Window popped out of collection
Private Function PopCol(ByRef col As collection) As stdWindow
  Set PopCol = col(col.Count)
  Call col.remove(col.Count)
End Function

'Shift a value out of a collection
'@param {ByRef Collection<stdWindow>} Collection to shift value out of
'@returns {stdWindow} Window shifted out of collection
Private Function ShiftCol(ByRef col As collection) As stdWindow
  Set ShiftCol = col(1)
  Call col.remove(1)
End Function

'Obtain a DWord from the high and low parts
'@param {ByVal Integer} High word
'@param {ByVal Integer} Low word
'@returns {Long} Long dword returned as the combination of high and low word
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
Private Sub BindGlobals(ByRef query As stdICallable)
    Dim bSuccess As Boolean
    Dim keys: keys = Lookups.keys()
    For Each key In keys
      Call query.SendMessage("bindGlobal", bSuccess, Array(key, Lookups(key)("S2N")))
    Next
End Sub

'Constructor for a tFindNode type
'@param {Long} depth of element
'@param {stdWindow} element
'@returns {tFindNode} FindNode struct
Private Function CreateFindNode(ByVal depth As Long, ByVal element As stdWindow) As tFindNode
    CreateFindNode.initialised = True
    CreateFindNode.depth = depth
    Set CreateFindNode.element = element
End Function

'Given an array `stack`, and length `index`, append `item` onto the stack
'@param {tFindNode()} stack of elements
'@param {Long} current index in stack
'@param {tFindNode} Item to insert into stack
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
    Set stackPopV.element = stack(index).element
    stack(index).initialised = False
    stack(index).depth = 0
    Set stack(index).element = Nothing
End Function

'From an array of Key-Value pairs, create a dictionary of Key-->Value and Value-->Key lookups
'@param {Variant()} 1D array of Key,Value pairs
'@returns {Dictionary} Dictionary<{S2N: Dictionary<key,value>, N2S: Dictionary<value, key> }>
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
'@param {Long} Atom to convert to a string
'@returns {String} Classname representing atom
Private Function MakeIntAtom(ByVal iAtom As Long) As String
  MakeIntAtom = "#" & iAtom
End Function

'TODO: Docs
Private Sub setThreadInput(ByVal bAttach As Boolean)
    Dim dwExcelTID As Long: dwExcelTID = GetCurrentThreadId()
    Dim dwThisTID As Long: dwThisTID = GetWindowThreadProcessId(pHandle, 0)
    
    'Synchonise Excel's thread with this window's thread
    If AttachThreadInput(dwExcelTID, dwThisTID, iif(bAttach, 1, 0)) = 0 Then
        Err.Raise Err.LastDllError
    End If
End Sub

'Parses a set of keys and converts them into a KeyToken array
'@param {String} Key string to tokenise
'@returns {KeyToken()} Array of key tokens containing all data required for PostMessage / SendInput
'@example   `stdWindow.TokeniseKeys("^a{home up 3}{LControl up}")`
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
'@param {String} KeyName as string
'@returns {EVirtualKey}
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
    Case Else
      Err.Raise 1, "VKeyFromString()", "No key """ & s & """"
  End Select
End Function

'Create an lParam key from params
'TODO: Docs
Private Function createKeyLParam(ByVal repeatCount As Long, ByVal scanCode As Long, ByVal extended As Boolean, ByVal bDown As Boolean) As Long
  createKeyLParam = repeatCount Or lshift(scanCode, 16) Or lshift(extended, 24) Or lshift(iif(bDown, 0, 1), 29) Or lshift(iif(bDown, 0, 1), 30) Or lshift(iif(bDown, 0, 1), 31)
End Function

'TODO: Docs
Private Function lshift(ByVal val As Long, ByVal nTimes As Integer) As Long
  lshift = val * 2 ^ nTimes
End Function

'Public Sub Requires()
'  if IsEmpty(stdEnumerator) then Call Msgbox("Requires `stdEnumerator`", vbInformation) else stdEnumerator.Requires
'
'End Sub
