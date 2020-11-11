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

Public Property Get handle() as LongPtr
  handle = pHandle
End Property

Public Property Get hDC() as LongPtr
  hDC = GetWindowDC(pHandle)
End Property

Public Property Get Exists as Boolean
  Exists = IsWindow(pHandle)
End Property

Public Property Get IsVisible() as Boolean
  if Exists then
    IsVisible = IsWindowVisible(pHandle)
  else
    CriticalRaise "IsVisible", "Window does not exist."
  end if
End Property
Public Property Get IsMinimised() as Boolean
  If Exists then
    IsMinimised = IsIconic(pHandle)
  else
    CriticalRaise "IsMinimised", "Window does not exist."
  end if
End Property
Public Property Get IsMaximised() as Boolean
  If Exists then
    IsMaximised = IsZoomed(pHandle)
  else
    CriticalRaise "IsMaximised", "Window does not exist."
  end if
End Property
Public Property Get IsHung() as Boolean
  if Exists then
    IsHung = IsHungAppWindow(pHandle)
  else
    CriticalRaise "IsHung", "Window does not exist."
  end if
End Property

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

Public Property Get Class() as string
  if Exists then
    Dim sCharBuffer as string, iNumChars as long
    iNumChars = GetClassName(hwnd, sCharBuffer, 256)
    Class = Mid(sCharBuffer, 1, iNumChars)
  else
    CriticalRaise "Class", "Window does not exist."
  end if
End Property

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
Public Property Get ProcessID() as long
  if Exists then
    Call GetWindowThreadProcessId(pHandle, ProcessID)
  else
    CriticalRaise "ProcessID", "Window does not exist."
  end if
End Property
Public Property Get ProcessName() as string
  if Exists then
    Dim sCharBuffer as string, iNumChars as long
    iNumChars = GetWindowModuleFileName(pHandle, sCharBuffer, 256)
    ProcessName = Mid(sCharBuffer, 1, iNumChars)
  else
    CriticalRaise "ProcessName", "Window does not exist."
  end if
End Property

Public Property Get Parent() as stdWindow
  set Parent = stdWindow.CreateFromHwnd(GetParent(pHandle))
End Property
Public Property Set Parent(ByVal win as stdWindow)
  Dim hOldParent as LongPtr
  hOldParent = SetParent(Me.Handle, win.Handle)
End Property

Public Property Get WindowLong() as Long
  'TODO:
End Property
Public Property Let WindowLong() as Long
  'TODO:
End Property

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

'https://github.com/AutoHotkey/AutoHotkey/blob/4b8aa248906771a3e25744683dd161ef53c7555a/Source/keyboard_mouse.cpp#L95
Public Sub SendKeysPlay()
  if Exists then
    'TODO:
  else
    CriticalRaise "SendKeysPlay", "Window does not exist."
  end if
End Sub
Public Sub SendKeysInput()
  if Exists then
    'TODO:
  else
    CriticalRaise "SendKeysInput", "Window does not exist."
  end if
End Sub
Public Sub SendKeysEvent()
  if Exists then
    'TODO:
  else
    CriticalRaise "SendKeysEvent", "Window does not exist."
  end if
End Sub


Public Sub Show()
  if Exists then
    'TODO:
  else
    CriticalRaise "Show", "Window does not exist."
  end if
End Sub
Public Sub Hide()
  if Exists then
    'TODO:
  else
    CriticalRaise "Hide", "Window does not exist."
  end if
End Sub
Public Sub Maximize()
  if Exists then
    'TODO:
  else
    CriticalRaise "Maximize", "Window does not exist."
  end if
End Sub
Public Sub Minimize()
  if Exists then
    'TODO:
  else
    CriticalRaise "Minimize", "Window does not exist."
  end if
End Sub

Public Sub Activate()
  if Exists then
    'If the window isn't visible, show it
    if not Visible then Call Show()

    'Bring the window to the front
    'TODO:
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
Public Function Find(ByVal callable as stdICallable, Optional ByVal searchType as ESearchType = ESearchType.BredthFirst) as stdWindow
  if Exists then
    'TODO:
  else
    CriticalRaise "Find", "Window does not exist."
  end if
End Property
Public Function FindAll(ByVal callable as stdICallable) as stdEnumerator
  if Exists then
    'TODO:
  else
    CriticalRaise "FindAll", "Window does not exist."
  end if
End Property