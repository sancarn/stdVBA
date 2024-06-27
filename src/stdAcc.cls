VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Spec:
'A class used to automate  objects which implement IAccessible
'
'
'CONSTRUCTORS
'    [X] CreateFromPoint
'    [X] CreateFromHwnd
'    [X] CreateFromApplication
'    [X] CreateFromDesktop
'    [X] CreateFromIAccessible
'    [X] CreateFromMouse
'PUBLIC INSTANCE METHODS:
'    [X] CreateFromPath
'    [X] GetDescendents
'    [X] FindFirst(stdICallable)
'    [X] FindAll(stdICallable)
'    [X] DoDefaultAction()
'    [X] getPath
'    [X] PrintChildTexts()
'    [X] PrintDescTexts(sFile?)
'    [X] SendMessage             TODO: Consider removing when we have stdWindow?
'    [ ] Highlight(seconds?)     TODO: Requires implementation
'    [ ] SendKeys() method       TODO: Consider whether this would be useful or not
'    [ ] SendKeysRaw() method    TODO: Consider whether this would be useful or not
'PROPERTIES:
'    [X] R   Parent
'    [X] R   Children
'    [X] R   hwnd
'    [X] R   location
'    [X] R   HitTest             TODO: Reconsider if this is useful
'    [ ] R   Selection           TODO: Returns Collection<stdAcc>. Need to figure out how to do VT_I4 proxy children, without access to parent - does this matter?
'    [X] R/W Value
'    [X] R   Name
'    [X] R   DefaultAction
'    [X] R   Role
'    [X] R   State
'    [X] R   Description
'    [X] R   KeyboardShortcut
'    [X] R/W Focus
'    [X] R   Help
'    [X] R   HelpTopic
'    [X] R   Text
'PRIVATE INSTANCE METHODS
'    getChildrenAsIAcc           TODO: Consider moving this to `Children`
'GENERIC
'TODO: VB6 complient, 32-Bit complient, 64-Bit complient

'TODO: Error in hwnd - AccessibleObjectFromWindow returns an E_FAIL code. Potentially related to https://stackoverflow.com/a/8617584/6302131
'                      Need to test if coinitialise has any other affects on excel. Might also be related to Excel trying to access it's own object
'                      model? Not sure...


'For CreateFromMouse()
#If VBA7 Then
  Private Declare PtrSafe Function GetCursorPos Lib "user32.dll" (lpPoint As tPOINT) As Long
  'FIX: stdAcc::CreateFromPoint() crashes on 64-bit.
  'The reason is that in reality the first arg of `AccessibleObjectFromPoint` is a ByVal `tPOINT` structure. Unfortunately in VBA you can't pass
  'structs by value. a ByRef struct won't do however, so instead we have to fake the data by picking args with the same size. In 32-bit you used to be able
  'to use 2 longs (the 2 parts of the struct), however 64-bit doesn't appear to like this notation. Instead we have to use a LongLong type to get the job done.
  #If Win64 Then
    Private Declare PtrSafe Function AccessibleObjectFromPoint Lib "oleacc" (ByVal arg1 As LongLong, ppacc As IAccessible, pVarChild As Variant) As Long
  #Else
    Private Declare PtrSafe Function AccessibleObjectFromPoint Lib "oleacc" (ByVal lX As Long, ByVal lY As Long, ppacc As IAccessible, pVarChild As Variant) As Long
  #End If
#Else
  Private Declare Function GetCursorPos Lib "user32.dll" (lpPoint As tPOINT) As Long
  Private Declare Function AccessibleObjectFromPoint Lib "Oleacc" (ByVal lX As Long, ByVal lY As Long, ppacc As IAccessible, pvarChild As Variant) As Long
#End If

'
#If VBA7 Then
  Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hwnd As LongPtr, ByVal dwId As Long, riid As tGUID, ppvObject As Object) As Long
  Private Declare PtrSafe Function AccessibleObjectFromEvent Lib "oleacc.dll" (hwnd As LongPtr, dwObjectId As Integer, dwChildID As Integer, ppacc As IAccessible, pVarChild As Object) As Long
  Private Declare PtrSafe Function AccessibleChildren Lib "oleacc" (ByVal paccContainer As IAccessible, ByVal iChildStart As Long, ByVal cChildren As Long, rgvarChildren As Variant, pcObtained As Long) As Long
  Private Declare PtrSafe Function WindowFromAccessibleObject Lib "oleacc" (ByVal pacc As IAccessible, ByRef pHwnd As LongPtr) As Long
  Private Declare PtrSafe Function APISendMessage Lib "user32" Alias "SendMessage" (ByVal hwnd As LongPtr, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare PtrSafe Function CoTaskMemFree Lib "ole32" (ByVal pv As LongPtr) As Long
  Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)
  Private Declare PtrSafe Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As LongPtr, ByVal offsetinVft As LongPtr, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByVal paTypes As LongPtr, ByVal paValues As LongPtr, ByRef retVAR As Variant) As Long
  Private Declare PtrSafe Function IIDFromString Lib "ole32.dll" (ByVal pstCLS As LongPtr, ByRef iid As GUID) As Long
#Else
  Private Declare Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hwnd As Long, ByVal dwId As Long, riid As tGUID, ppvObject As Object) As Long
  Private Declare Function AccessibleObjectFromEvent Lib "oleacc.dll" (hwnd As Long, dwObjectId As Integer, dwChildID As Integer, ppacc As IAccessible, pVarChild As Object) As Long
  Private Declare Function AccessibleChildren Lib "oleacc" (ByVal paccContainer As IAccessible, ByVal iChildStart As Long, ByVal cChildren As Long, rgvarChildren As Variant, pcObtained As Long) As Long
  Private Declare Function WindowFromAccessibleObject Lib "oleacc" (ByVal pacc As IAccessible, ByRef pHwnd As Long) As Long
  Private Declare Function APISendMessage Lib "user32" Alias "SendMessage" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function CoTaskMemFree Lib "ole32" (ByVal pv As Long) As Long
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As LongPtr)
  Private Declare Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByVal paTypes As Long, ByVal paValues As Long, ByRef retVAR As Variant) As Long
  Private Declare Function IIDFromString Lib "ole32.dll" (ByVal pstCLS As Long, ByRef iid As GUID) As Long
#End If

'GUID struct for QueryInterface
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

'Used while walking the Accessibility tree. Can be used to toggle between a Breadth first search and a depth first search.
Public Enum EAccFindType
    BreadthFirst = 0
    DepthFirst = 1
End Enum

'Used while walking the Accessibility tree. Can be used to discard entire trees of elements, to increase speed of walk algorithms.
Public Enum EAccFindResult
    MatchFound = 1                 'Matched                                    
    MatchFoundSearchDescendents=4  'Same as `ESearchResult.MatchFound`         
    NoMatchFound = 0               'Not found, continue searching descendents  
    NoMatchCancelSearch= 2         'Not found, cancel search                   
    NoMatchSkipDescendents= 3      'Not found, don't search descendents        
End Enum

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


'Not sure if this is the correct order, not sure if it's even needed given IAccessible apparently implements IDispatch...
Private Enum IAccessible_Methods
    accDoDefaultAction
    accHitTest
    accLocation
    accNavigate
    accSelect
    get_accChild
    get_accChildCount
    get_accDefaultAction
    get_accDescription
    get_accFocus
    get_accHelp
    get_accHelpTopic
    get_accKeyboardShortcut
    get_accName
    get_accParent
    get_accRole
    get_accSelection
    get_accState
    get_accValue
    put_accName
    put_accValue
End Enum

'GUID type for ObjectFromWindow
Private Type tGUID
    lData1            As Long
    nData2            As Integer
    nData3            As Integer
    abytData4(0 To 7) As Byte
End Type
Private Type tPOINT
    x As Long
    y As Long
End Type
Private Type tFindNode
    initialised as boolean
    depth as long
    element as object
End Type


Private Const CHILDID_SELF = 0&

Private Type TSingleton
  Lookups As Object
End Type
Private Type TProxy
    Parent As IAccessible
    Index As Long
End Type
Private Type TThis
    Singleton as TSingleton

    isProxy As Boolean
    Proxy as TProxy
    IAccessible As Object
End Type
Private This as TThis

'****************
'* Constructors *
'****************

'Creates an `stdAcc` object from an `X` and `Y` point location on the screen.
'@constructor
'@param x - X Coordinate
'@param y - Y Coordinate
'@returns stdAcc - Object at the specified point
Public Function CreateFromPoint(ByVal x As Long, ByVal y As Long) As stdAcc
    Set CreateFromPoint = New stdAcc
    set CreateFromPoint.protAccessible = AccessibilityObjectFromPointEx(x,y)
End Function

'Creates an `stdAcc` object from a window handle.
'@constructor
'@param hwnd - Window handle to wrap
'@returns stdAcc - stdAcc object for the window 
Public Function CreateFromHwnd(ByVal hwnd As LongPtr) As stdAcc
    Dim acc as stdAcc: set acc = new stdAcc
    Dim obj As IAccessible
    Set obj = IAccessibleFromHwnd(hwnd)
    set acc.protAccessible = obj
    Set CreateFromHwnd = acc
End Function

'Creates an `stdAcc` object from the current running application (e.g. Excel / Word / Powerpoint).
'@constructor
'@param oApp - Application to create from. Defaults to `Application` if unspecified.
'@returns - IAccessible wrapper around application window.
Public Function CreateFromApplication(Optional ByVal oApp as Object = Nothing) As stdAcc
  if oApp is nothing then set oApp = Application
  select case oApp.Name
    case "Microsoft Excel"
      Set CreateFromApplication = CreateFromHwnd(oApp.hwnd)
    case "Microsoft Word"
      Set CreateFromApplication = CreateFromHwnd(oApp.ActiveWindow.Hwnd)
    case "Microsoft Access"
      set CreateFromApplication = CreateFromHwnd(oApp.hWndAccessApp)
    case "Microsoft PowerPoint"
      set CreateFromApplication = CreateFromIAccessible(oApp.CommandBars("Status Bar")).AncestralRoot
    Case "Outlook"
      Set CreateFromApplication = CreateFromIAccessible(oApp.ActiveWindow.CommandBars("Status Bar")).AncestralRoot
    Case "Microsoft Publisher"
      Set CreateFromApplication = CreateFromIAccessible(oApp.CommandBars("Status Bar")).AncestralRoot
    case else
      Err.Raise 1, "stdAcc::CreateFromApplication()", "No implementation for getting application window of " & Application.name
  end select
End Function

'Creates an `stdAcc` object from the desktop.
'@constructor
'@returns - IAccessible wrapper around desktop window.
Public Function CreateFromDesktop() As stdAcc
    'Get this application's accessibility object
    Dim accThis As IAccessible
    Set accThis = IAccessibleFromHwnd(Application.hwnd)
    
    'Set desktop
    Dim accDesktop As IAccessible
    Set accDesktop = accThis.accParent
    
    Dim acc as stdAcc: set acc = new stdAcc
    set acc.protAccessible = accDesktop
    Set CreateFromDesktop = acc
End Function

'Creates an `stdAcc` object from an object which implements `IAccessible`.
'@constructor
'@param obj - Object implementing `IAccessible` interface.
'@returns - IAccessible wrapper around an IAccessible object.
Public Function CreateFromIAccessible(ByVal obj As IAccessible) As stdAcc
    Set CreateFromIAccessible = New stdAcc
    set CreateFromIAccessible.protAccessible = obj
End Function

'Creates an `stdAcc` object for the element the mouse currently hovers over.
'@constructor
'@returns - IAccessible wrapper around element under mouse.
Public Function CreateFromMouse() As stdAcc
    Dim pT As tPOINT
    Dim success As Long
    success = GetCursorPos(pT)
    Set CreateFromMouse = CreateFromPoint(pt.x,pt.y)
End Function

'Creates an `stdAcc` object for the element at a given path from the current element.
'@constructor
'@param sPath - Path to element, e.g. "P.1.2.3" or "1.2.3". If starting with "P." or ".", these are ignored as they are taken as the current element.
'@returns - IAccessible wrapper from a supplied path.
'@example `Debug.Print stdAcc.CreateFromApplication().CreateFromPath("3.1").name`
Public Function CreateFromPath(ByVal sPath As String) As stdAcc
    'If starting with "." remove it
    If left(sPath, 1) = "P" Then sPath = Mid(sPath, 2)
    If left(sPath, 1) = "." Then sPath = Mid(sPath, 2)
    
    'Get descendants list
    Dim descendants As Variant
    descendants = Split(sPath, ".")
    
    'Initiate acc (used for tracing through descendants)
    Dim acc As stdAcc
    Set acc = Me
    
    'Loop over descendants
    On Error GoTo ExitFunction 'theoretically this should never error, but still sometimes does...
      Dim i As Integer
      For i = 0 To UBound(descendants)
          If CLng(descendants(i)) > acc.children.Count Then Exit Function
          Set acc = acc.children(CLng(descendants(i)))
      Next i
    On Error GoTo 0
    
    'Return descendant
    Set CreateFromPath = acc
ExitFunction:
End Function

'Initialises an stdAcc object as a `Proxy` object, who's methods are implemented on the parent instead of on the element itself
'@constructor
'@protected
'@param oParent - Parent object to proxy
'@param index - Index of element in parent
Public Sub protInitWithProxy(ByRef oParent as stdAcc, ByVal index as long)
    this.isProxy = true
    set this.Proxy.Parent = oParent.protAccessible
    this.Proxy.Index = index
End Sub

'The object which implements `IAccessible`, with which this class wraps. Please don't modify this object, unless you know what you are doing.
'@protected 
'@returns Object<IAccessible> - The object which implements `IAccessible`
Public Property Get protAccessible() As Object
    Set protAccessible = This.IAccessible
End Property
Public Property Set protAccessible(ByVal obj As Object)
    Set This.IAccessible = obj
End Property


'Get all descendents of the stdAcc control
'@returns Collection<stdAcc> - Collection of descendents
Public Function GetDescendents() As Collection
    'Create collection which will be returned
    Dim c As Collection
    Set c = New Collection
    
    'Loop over all children...
    Dim accChild As stdAcc, accDesc As stdAcc
    For Each accChild In children
        'Add children to collection
        c.Add accChild
        
        'Loop over descendents and add these to collection also (recurse)
        For Each accDesc In accChild.GetDescendents
            c.Add accDesc
        Next
    Next
    
    'Return descendents
    Set GetDescendents = c
End Function

'Search the IAccessible tree for elements which match a certain criteria. Return the first element found.
'@param query as stdICallable<(stdAcc,depth)=>EAccFindResult> - Callback returning `EAccFindResult` options:
'
'* `EAccFindResult.NoMatchFound`/`0`/`False`        - Not found, countinue walking
'* `EAccFindResult.MatchFound`/`1`/`-1`/`True`      - Found, return this element
'* `EAccFindResult.NoMatchCancelSearch`/`2`         - Not found, cancel search
'* `EAccFindResult.NoMatchSkipDescendents`/`3`      - Not found, don't search descendents
'* `EAccFindResult.MatchFoundSearchDescendents`/`4` - Same as `EAccFindResult.MatchFound` in this case.
'@param searchType - The type of search to perform, either Breadth First Search (BFS) or Depth First Search (DFS).
' To understand the difference between BFS and DFS take this tree:
'```
'     A
'    / \
'   B   C
'  /   / \
' D   E   F
'```
'
'* A BFS will walk this tree in the following order: A, B, C, D, E, F
'* A DFS will walk this tree in a different order:   A, C, F, E, B, D
'@returns {stdAcc} - Element found.
'@examples ```
''Find where name is "hello" and class is "world":
'el.FindFirst(stdLambda.Create("$1.name=""hello"" and $1.class=""world"""))
''Find first element named "hello" at depth > 4:
'el.FindFirst(stdLambda.Create("$1.name = ""hello"" AND $2 > 4"))
'```
Public Function FindFirst(ByVal query As stdICallable, optional byval searchType as EAccFindType=EAccFindType.DepthFirst) As stdAcc
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
            case EAccFindType.DepthFirst
                'Depth first search, so pop the item out of the stack
                part = stackPopV(stack,length)
            case EAccFindType.BreadthFirst
                'Breadth first search, get item directly out of array, no need to change array size
                index = index + 1
                part = stack(index)
            case else
                Err.Raise 1, "stdAcc#FindFirst", "Invalid search type given. Please use EAccFindType"
        end select
        
        With part
            If not .initialised Then Exit Function
            
            'Run query and test options
            Select Case query.Run(.element, .depth)
                Case EAccFindResult.NoMatchFound
                    'Nothing found, search descendents
                    Dim child As stdAcc
                    For Each child In part.element.children
                        Call stackPushV(stack, length, CreateFindNode(.depth + 1, child))
                    Next
                Case EAccFindResult.MatchFound, True, EAccFindResult.MatchFoundSearchDescendents
                    'Found, return element
                    Set FindFirst = .element
                    Exit Function
                Case EAccFindResult.NoMatchCancelSearch
                    'Nothing found, cancel function
                    Set FindFirst = Nothing
                    Exit Function
                case EAccFindResult.NoMatchSkipDescendents
                '    Nothing found, don't search descendents
            End Select
        End With
        'Just make sure no freezing occurs
        DoEvents
    Wend
    
    'Else set to nothing
    Set FindFirst = Nothing
End Function

'Search the IAccessible tree for elements which match a certain criteria. Return all elements found.
'@param query as stdICallable<(stdAcc,depth)=>EAccFindResult> - Callback returning `EAccFindResult` options:
'
'* `EAccFindResult.NoMatchFound`/`0`/`False`           - Not found, countinue walking.
'* `EAccFindResult.MatchFound`/`1`/`-1`/`True`         - Found, return this element, won't search descendents of elements found.
'* `EAccFindResult.NoMatchCancelSearch`/`2`            - Not found, cancel search.
'* `EAccFindResult.NoMatchSkipDescendents`/`3`         - Not found, don't search descendents.
'* `EAccFindResult.MatchFoundSearchDescendents`/`4`    - Found, return this element, but continue searching descendents.
'@param searchType - The type of search, either Breadth First Search (BFS) or Depth First Search (DFS).
' To understand the difference between BFS and DFS take this tree:
'```
'     A
'    / \
'   B   C
'  /   / \
' D   E   F
'```
'
'* A BFS will walk this tree in the following order: A, B, C, D, E, F.
'* A DFS will walk this tree in a different order:   A, C, F, E, B, D.
'@returns Collection<stdAcc> - Collection of elements found.
'@examples ```
''Find where name is "hello" and class is "world":
'el.FindAll(stdLambda.Create("$1.name=""hello"" and $1.class=""world"""))
''Find all elements with depth <= 4:
'el.FindAll(stdLambda.Create("if $2 < 4 then 4 else if $2 = 4 then 1 else 3"))
'```
Public Function FindAll(ByVal query As stdICallable, optional byval searchType as EAccFindType=EAccFindType.DepthFirst) As Collection
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
            case EAccFindType.DepthFirst
                'Depth first search, so pop the item out of the stack
                part = stackPopV(stack,length)
            case EAccFindType.BreadthFirst
                'Breadth first search, get item directly out of array, no need to change array size
                index = index + 1
                part = stack(index)
            case else
                Err.Raise 1, "stdAcc#FindAll", "Invalid search type given. Please use EAccFindType"
        end select
        
        With part
            'When hitting the edge of the stack quit
            if not .initialised then Exit Function

            'Run query and test options
            Dim child as stdAcc
            select case query.Run(.element,.depth)
                Case EAccFindResult.NoMatchFound
                    'Nothing found, search descendents
                    for each child in .element.children
                        Call stackPushV(stack,length,CreateFindNode(.depth+1,child))
                    next
                Case EAccFindResult.MatchFound, True
                    'Found, add element
                    Call FindAll.add(.element)
                Case EAccFindResult.NoMatchCancelSearch
                    'Nothing found, cancel function
                    Exit Function
                Case EAccFindResult.NoMatchSkipDescendents
                    'Nothing found, don't search descendents
                Case EAccFindResult.MatchFoundSearchDescendents
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

'Return the parent of the IAccessible object
'@returns stdAcc - Parent of this element
Public Property Get Parent() As stdAcc
    If this.isProxy Then
        Set Parent = stdAcc.CreateFromIAccessible(this.Proxy.Parent)
    else
        On Error GoTo handle_error
            Set Parent = stdAcc.CreateFromIAccessible(protAccessible.accParent)
        On Error GoTo 0
    End if
    Exit Property
handle_error:
    Set Parent = Nothing
End Property

'Obtain the ancestral root of the IAccessible object (A child of the desktop)
'@returns - Ancestral root of this element (child window of the desktop)
Public Property Get AncestralRoot() As stdAcc
    Dim desktop as stdAcc: set desktop = stdAcc.CreateFromDesktop()
    'Walk up the tree until we hit the element just before desktop
    Dim acc As stdAcc: Set acc = Me
    While Not acc.Parent.hwnd <> desktop.hwnd
        Set acc = acc.Parent
    Wend
    Set AncestralRoot = acc
End Property

'Return the children of the IAccessible object
'@returns Collection<stdAcc> - Children of this element
Public Property Get children() As Collection
    On Error GoTo ErrorHandler
        Set children = getChildrenAsIAcc()
    On Error GoTo 0
    Exit Property
ErrorHandler:
    Set children = New Collection
End Property

'Return the hwnd of the IAccessible object
'@returns LongPtr - hwnd of this element
#if VBA7 then
  Public Property Get hwnd() As LongPtr
#Else
  Public Property Get hwnd() As Long
#End If
    On Error GoTo handle_error
        Dim lHwnd As LongPtr
        Dim res As Long
        res = WindowFromAccessibleObject(protAccessible, lHwnd)
        If res = 0 Then
            hwnd = lHwnd
        Else
            Call Err.Raise(1, "", "Error " & Hex(res) & " - " & COMErrorDescription(res) & " in stdAcc#hwnd")
        End If
    On Error GoTo 0
    Exit Property
handle_error:
    hwnd = 0
End Property

'Get a unique identifier for an element.
'@returns string - Hex string identifier for element.
'@example `stdAcc.CreateFromMouse().Identity`
Public Property Get Identity() As String
  On Error GoTo Fallback
  Dim accIdent As IUnknown: Set accIdent = unkQueryInterface(protAccessible, "{7852B78D-1CFD-41C1-A615-9C0C85960B5F}")
  If Not accIdent Is Nothing Then Identity = accGetIdentity(accIdent)
  Exit Property
Fallback:
  'TODO: Fallback identity
  Identity = "Unknown"
End Property

'Return the location of the element as a collection
'@returns Collection<Long|stdAcc> - The location of the element. Collection has 5 named keys: "Width", "Height", "Left", "Top" and "Parent". "Parent" refers to the element itself.
Public Property Get Location() As Collection
    On Error GoTo ErrorHandler
        'Get location from protAccessible
        Dim pcxWidth As Long
        Dim pcyHeight As Long
        Dim pxLeft As Long
        Dim pyTop As Long
        Call protAccessible.accLocation(pxLeft, pyTop, pcxWidth, pcyHeight, CHILDID_SELF)
        
        'Create location collection
        Set Location = New Collection
        Location.Add pcxWidth, "Width"
        Location.Add pcyHeight, "Height"
        Location.Add pxLeft, "Left"
        Location.Add pyTop, "Top"
        Location.Add Me, "Parent"
    On Error GoTo 0
    Exit Property
ErrorHandler:
    Set Location = Nothing
End Property

'Return the element under the specified location
'TODO: Is this needed?
'@param x - x coord
'@param y - y coord
'@returns stdAcc - element under point
Public Property Get HitTest(ByVal x As Long, ByVal y As Long) As stdAcc
  Set HitTest = New stdAcc
  Dim NewAcc As Object
  Call protAccessible.accHitTest(x, y, NewAcc)
  set HitTest.protAccessible = NewAcc
End Property

'Gets/Sets the value of the element
'@returns String - the current value of the element
Public Property Get value() As String
  On Error Resume Next 'FIX: Sometimes this is unsupported, so leave this error.
  if this.isProxy then
    value = this.Proxy.Parent.accValue(this.Proxy.Index)
  else
    value = protAccessible.accValue(CHILDID_SELF)
  end if
End Property
Public Property Let value(val As String)
  if this.isProxy then
    this.Proxy.Parent.accValue(this.Proxy.Index) = val
  else
    protAccessible.accValue(CHILDID_SELF) = val
  end if
End Property

'Returns the name of the element
'@returns String - the name of the element
Public Property Get name() As String
  On Error Resume Next 'FIX: Sometimes this is unsupported, so leave this error.
  If this.isProxy Then
    name = this.Proxy.Parent.accName(this.Proxy.Index)
  Else
    name = protAccessible.accName(CHILDID_SELF)
  End If
End Property

'Returns the description of the default action of the element
'@returns String - the description of the default action of the element
Public Property Get DefaultAction() As String
  On Error Resume Next 'FIX: Sometimes this is unsupported, so leave this error.
  If this.isProxy Then
    DefaultAction = this.Proxy.Parent.accDefaultAction(this.Proxy.Index)
  else
    DefaultAction = protAccessible.accDefaultAction(CHILDID_SELF)
  end if
End Property

'Returns the role of the element
'@returns String - the role of the element
Public Property Get Role() As String
  On Error Resume Next 'FIXME: On extremely rare occasions the OLE object will disconnect / not expect automation. `On Error Resume Next` silences the error but this isn't a real fix.
  if this.isProxy then
    Role = this.Singleton.Lookups("EAccRoles")("N2S")(this.Proxy.Parent.accRole(this.Proxy.Index))
  else
    Role = this.Singleton.Lookups("EAccRoles")("N2S")(protAccessible.accRole(CHILDID_SELF))
  end if
End Property

'Returns the state of the element
'@returns String - the state of the element
Public Property Get States() As String
  Dim iStateData As Long: iStateData = StateData
  Dim oLookup As Object: Set oLookup = this.Singleton.Lookups("EAccStates")("S2N")
  Dim sStates As String, vKey as variant
  For Each vKey In oLookup.keys()
    If (oLookup(vKey) And iStateData) = oLookup(vKey) Then
        sStates = sStates & vKey & ";"
    End If
  Next
  
  States = sStates
End Property

'Returns the state(s) of the element
'@returns Long - the state(s) of the element
Public Property Get StateData() As Long
  On Error Resume Next 'FIXME: On extremely rare occasions the OLE object will disconnect / not expect automation. `On Error Resume Next` silences the error but this isn't a real fix.
  if this.isProxy then
    StateData = this.Proxy.Parent.accState(this.Proxy.Index)
  else
    StateData = protAccessible.accState(CHILDID_SELF)
  end if
End Property

'Returns the description of the element
'@returns String - the description of the element
Public Property Get Description() As String
  On Error Resume Next 'FIX: Sometimes this is unsupported, so leave this error.
  if this.isProxy then
    Description = this.Proxy.Parent.accDescription(this.Proxy.Index)
  else
    Description = protAccessible.accDescription(CHILDID_SELF)
  end if
End Property

'Returns the keyboard shortcut used to trigger the default action of the element
'@returns String - The keyboard shortcut used to trigger the default action of the element
Public Property Get KeyboardShortcut() As String
  On Error Resume Next 'FIX: Sometimes this is unsupported, so leave this error handling.
  if this.isProxy then
    KeyboardShortcut = this.Proxy.Parent.accKeyboardShortcut(this.Proxy.Index)
  else
    KeyboardShortcut = protAccessible.accKeyboardShortcut(CHILDID_SELF)
  end if
End Property

'Get the Focus of the UI element
'@returns Boolean - Whether the element has focus
Public Property Get Focus() As Boolean
  'On Error Resume Next
  if this.isProxy then
    Focus = this.Proxy.Parent.accFocus(this.Proxy.Index)
  else
    Focus = protAccessible.accFocus(CHILDID_SELF)
  end if
End Property

'Set the Focus of the UI element
'@param Boolean - `True` if the element should be focussed, otherwise `false`
Public Property Let Focus(ByVal val As Boolean)
  'On Error Resume Next
  if this.isProxy then
    this.Proxy.Parent.accFocus(this.Proxy.Index) = val
  else
    protAccessible.accFocus(CHILDID_SELF) = val
  end if
End Property

'Returns the help
'@returns String - the help for the element
Public Property Get Help() As String
  On Error Resume Next 'FIX#97: Sometimes this is unsupported, so leave this error handling.
  if this.isProxy then
    Help = this.Proxy.Parent.accHelp(this.Proxy.Index)
  else
    Help = protAccessible.accHelp(CHILDID_SELF)
  end if
End Property

'Returns the help topic
'@param sFile - File to get help topic for
'@returns - the help topic for the element
Public Property Get HelpTopic(Optional ByVal sFile as string = "") As String
  On Error Resume Next 'FIX: Sometimes this is unsupported, so leave this error.
  if this.isProxy then
    HelpTopic = this.Proxy.Parent.accHelpTopic(sFile, this.Proxy.Index)
  else
    HelpTopic = protAccessible.accHelpTopic(sFile, CHILDID_SELF)
  end if
End Property

'Returns a string containing numerous properties from stdAcc concatenated together.
'@returns String - Concatenation of Name, Value, DefaultAction, Description, Role, Help, HelpTopic and KeyboardShortcut.
Public Property Get Text() As String
  If Len(name & value & DefaultAction & Description & Help & HelpTopic & KeyboardShortcut) > 0 Then
    Text = "Name: """ & name & """; " & _
            "Value: """ & value & """; " & _
            "DefaultAction: """ & DefaultAction & """; " & _
            "Description: """ & Description & """; " & _
            "Role: """ & Role & """; " & _
            "Help: """ & Help & """; " & _
            "HelpTopic: """ & HelpTopic & """; " & _
            "KeyboardShortcut: """ & KeyboardShortcut & """;"
  Else
    Text = ""
  End If
End Property

'Performs the default action of the `IAccessible` object
'@example `app.getButton().DoDefaultAction()`
Public Sub DoDefaultAction()
  If this.isProxy Then
    Call this.Proxy.Parent.accDoDefaultAction(this.Proxy.Index)
  Else
    Call protAccessible.accDoDefaultAction(CHILDID_SELF)
  End If
End Sub

'Sends a message to the window
'@deprecated Use `stdWindow#SendMessage()` instead.
'@param Msg - Message to send
'@param wParam - WParam to send
'@param lParam - LParam to send
'@returns Long - Return value
Public Function SendMessage(ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If hwnd() > 0 Then
        SendMessage = APISendMessage(hwnd(), Msg, wParam, lParam)
    Else
        Err.Raise 1, "No hwnd on this window"
    End If
End Function

'Print the text of all child elements. Useful for debugging.
Public Sub PrintChildTexts()
    Dim child As stdAcc
    Dim iIndex As Long: iIndex = 0
    For Each child In children
        iIndex = iIndex + 1
        Debug.Print iIndex & ". " & child.Text
    Next
End Sub

'Prints all descendent texts. Useful for debugging. Also includes all paths of all descendents which can be very beneficial.
'@param sToFilePath   - File path to dump text to.
'@param bPrintToDebug - Whether to print output to debug window or not.
'@param sPath         - Prefix path (sometimes you might wnat to use Me.getPath(...))
'@param fileNum       - Filenum to output too, typically this will be left unprovided
Public Sub PrintDescTexts(Optional ByVal sToFilePath as string = "", Optional ByVal bPrintToDebug as boolean = true, Optional ByVal sPath As String = "P", Optional ByVal fileNum as long = 0)
    'Open file is file name passed and file num not given
    Dim bFileOpened as boolean
    if fileNum = 0 and len(sToFilePath) > 0 then
        bFileOpened = true
        fileNum = FreeFile()
        open sToFilePath for Output as #fileNum
    end if

    'Loop over all children
    Dim child As stdAcc
    Dim iIndex As Long: iIndex = 0
    For Each child In children
        'Get child index for path
        iIndex = iIndex + 1

        'Create path string
        Dim myPath As String: myPath = sPath & "." & iIndex

        'If file has been opened, print data to file
        If fileNum <> 0 Then Print #fileNum, myPath & ". " & child.Text
        
        'Only print to debug if bPrintToDebug
        if bPrintToDebug then Debug.Print myPath & ". " & child.Text

        'Recurse to all children
        Call child.PrintDescTexts(sToFilePath, bPrintToDebug, myPath, fileNum)
    Next

    'Only close file if file was opened within sub
    if bFileOpened then
        Close #fileNum
    end if
End Sub

'Returns the path to an element
'@param toAccessible - Stop creating path at this element
'@returns String - String representing path to element from either Desktop or toAccessible.
Public Function getPath(Optional toAccessible As stdAcc = Nothing) As String
    'Initialise trace
    Dim acc As stdAcc
    Set acc = Me
    
    'Collection to store path
    Dim col As Collection: Set col = New Collection
        
    'Collect parents
    if not acc.parent is nothing then
      While Not acc.parent.name = "Desktop"
          Dim child As stdAcc
          Dim index As Long: index = 0
          Dim savedIndex As Long: savedIndex = 0
          For Each child In acc.parent.children
              index = index + 1
              If child.hwnd = acc.hwnd And child.Role = acc.Role Then
                  savedIndex = index
                  Exit For
              End If
          Next
          
          'Add index to stack
          If savedIndex > 0 Then
              col.add savedIndex
          Else
              'In some scenarios this has occurred where acc tree has changed after initial pass. Not much we can do if this is the case.
              getPath = "UNKNOWN"
              Exit Function
          End If
                  
          'Elevate parent
          Set acc = acc.parent
      Wend
      
      'Create path
      Dim path As String
      Dim i As Integer
      For i = col.Count To 1 Step -1
          path = path & "." & col(i)
      Next i
      
      'Return path
      getPath = "D.W" & path
    else
      getPath = "D"
    end if
End Function

'Returns this element and all descendents as a JSON string. Useful for debugging
'@returns String - JSON representation of element tree
Public Function toJSON() As String
    'Get children string
    Dim sChildren As String
    sChildren = ""
    Dim acc As stdAcc
    For Each acc In children()
        sChildren = sChildren & "," & acc.toJSON()
    Next
    sChildren = Mid(sChildren, 2)
    
    'Convert to json
    toJSON = "{" & _
        """name"":""" & jsonEscape(name()) & """," & _
        """desc"":""" & jsonEscape(Description()) & """," & _
        """value"":""" & jsonEscape(value()) & """," & _
        """role"":""" & jsonEscape(Role()) & """," & _
        """states"":""" & jsonEscape(States()) & """," & _
        """ks"":""" & jsonEscape(KeyboardShortcut()) & """," & _
        """help"":""" & jsonEscape(Help()) & """," & _
        """ht"":""" & jsonEscape(HelpTopic()) & """," & _
        """children"":[" & sChildren & "]}"
End Function

'Await for the existence of an element which meets a specified condition.
'@param awaitFor as stdICallable<(stdAcc)=>Boolean> - If this callback returns true, the function will immediately set the result of the function to the awaited element, otherwise will continue awaiting.
'@param error as stdICallable<(stdAcc)=>Boolean>    - If this callback returns true, the function will immediately set the result of the function to `nothing`, oterwise will continue awaiting. If none is provided no error is checked.
'@param timeout                                   - Timeout in seconds. Default = NoTimeout. When timeout is reached the return value of the function is `nothing`
'@returns stdAcc                                  - Element awaited for
'@remark calls `Me.FindFirst(awaitFor)` to wait for chosen element
Public Function AwaitForElement(awaitFor as stdICallable, Optional error as stdICallable = Nothing, Optional timeout As Long = -1) As stdAcc
  Dim bCond As Boolean: bCond = True
  Dim el As Object: Set el = Nothing
  Dim oldTime As Date: oldTime = Now()
  While bCond
      Set el = FindFirst(awaitFor)
      If Not el Is Nothing Then bCond = False
      If Not error Is Nothing Then
          If error.Run(Me) Then bCond = False
      End If
      If timeout <> -1 And Second(Now() - oldTime) > timeout Then bCond = False
      DoEvents
  Wend
  Set AwaitForElement = el
End Function

'Returns the lookups object
'@returns Object<Dictionary<string,Dictionary<string,Dictionary<string,long>>>>
Friend Function protGetLookups() as Object
    if this.Singleton.Lookups is nothing then
        set this.Singleton.Lookups = CreateObject("Scripting.Dictionary")
        set this.Singleton.Lookups("EAccStates") = CreateLookupDict(Array( _
            "STATE_NORMAL", &H0,                "STATE_UNAVAILABLE", &H1, _
            "STATE_SELECTED", &H2,              "STATE_FOCUSED", &H4, _
            "STATE_PRESSED", &H8,               "STATE_CHECKED", &H10, _
            "STATE_MIXED", &H20,                "STATE_INDETERMINATE", &H99, _
            "STATE_READONLY", &H40,             "STATE_HOTTRACKED", &H80, _
            "STATE_DEFAULT", &H100,             "STATE_EXPANDED", &H200, _
            "STATE_COLLAPSED", &H400,           "STATE_BUSY", &H800, _
            "STATE_FLOATING", &H1000,           "STATE_MARQUEED", &H2000, _
            "STATE_ANIMATED", &H4000,           "STATE_INVISIBLE", &H8000, _
            "STATE_OFFSCREEN", &H10000,         "STATE_SIZEABLE", &H20000, _
            "STATE_MOVEABLE", &H40000,          "STATE_SELFVOICING", &H80000, _
            "STATE_FOCUSABLE", &H100000,        "STATE_SELECTABLE", &H200000, _
            "STATE_LINKED", &H400000,           "STATE_TRAVERSED", &H800000, _
            "STATE_MULTISELECTABLE", &H1000000, "STATE_EXTSELECTABLE", &H2000000, _
            "STATE_ALERT_LOW", &H4000000,       "STATE_ALERT_MEDIUM", &H8000000, _
            "STATE_ALERT_HIGH", &H10000000,     "STATE_PROTECTED", &H20000000, _
            "STATE_VALID", &H7FFFFFFF _
        ))
        set this.Singleton.Lookups("EAccRoles") = CreateLookupDict(Array( _
            "ROLE_TITLEBAR", &H1&            , "ROLE_MENUBAR", &H2&          , "ROLE_SCROLLBAR", &H3&, _
            "ROLE_GRIP", &H4&                , "ROLE_SOUND", &H5&            , "ROLE_CURSOR", &H6&, _
            "ROLE_CARET", &H7&               , "ROLE_ALERT", &H8&            , "ROLE_WINDOW", &H9&, _
            "ROLE_CLIENT", &HA&              , "ROLE_MENUPOPUP", &HB&        , "ROLE_MENUITEM", &HC&, _
            "ROLE_TOOLTIP", &HD&             , "ROLE_APPLICATION", &HE&      , "ROLE_DOCUMENT", &HF&, _
            "ROLE_PANE", &H10&               , "ROLE_CHART", &H11&           , "ROLE_DIALOG", &H12&, _
            "ROLE_BORDER", &H13&             , "ROLE_GROUPING", &H14&        , "ROLE_SEPARATOR", &H15&, _
            "ROLE_TOOLBAR", &H16&            , "ROLE_STATUSBAR", &H17&       , "ROLE_TABLE", &H18&, _
            "ROLE_COLUMNHEADER", &H19&       , "ROLE_ROWHEADER", &H1A&       , "ROLE_COLUMN", &H1B&, _
            "ROLE_ROW", &H1C&                , "ROLE_CELL", &H1D&            , "ROLE_LINK", &H1E&, _
            "ROLE_HELPBALLOON", &H1F&        , "ROLE_CHARACTER", &H20&       , "ROLE_LIST", &H21&, _
            "ROLE_LISTITEM", &H22&           , "ROLE_OUTLINE", &H23&         , "ROLE_OUTLINEITEM", &H24&, _
            "ROLE_PAGETAB", &H25&            , "ROLE_PROPERTYPAGE", &H26&    , "ROLE_INDICATOR", &H27&, _
            "ROLE_GRAPHIC", &H28&            , "ROLE_STATICTEXT", &H29&      , "ROLE_TEXT", &H2A&, _
            "ROLE_PUSHBUTTON", &H2B&         , "ROLE_CHECKBUTTON", &H2C&     , "ROLE_RADIOBUTTON", &H2D&, _
            "ROLE_COMBOBOX", &H2E&           , "ROLE_DROPLIST", &H2F&        , "ROLE_PROGRESSBAR", &H30&, _
            "ROLE_DIAL", &H31&               , "ROLE_HOTKEYFIELD", &H32&     , "ROLE_SLIDER", &H33&, _
            "ROLE_SPINBUTTON", &H34&         , "ROLE_DIAGRAM", &H35&         , "ROLE_ANIMATION", &H36&, _
            "ROLE_EQUATION", &H37&           , "ROLE_BUTTONDROPDOWN", &H38&  , "ROLE_BUTTONMENU", &H39&, _
            "ROLE_BUTTONDROPDOWNGRID", &H3A& , "ROLE_WHITESPACE", &H3B&      , "ROLE_PAGETABLIST", &H3C&, _
            "ROLE_CLOCK",&H3D&  _
        ))
        set this.Singleton.Lookups("EAccFindResult") = CreateLookupDict(Array( _
            "MatchFound", EAccFindResult.MatchFound, _
            "MatchFoundSearchDescendents", EAccFindResult.MatchFoundSearchDescendents, _
            "NoMatchFound", EAccFindResult.NoMatchFound, _
            "NoMatchCancelSearch", EAccFindResult.NoMatchCancelSearch, _
            "NoMatchSkipDescendents", EAccFindResult.NoMatchSkipDescendents _
        ))
    end if
    set protGetLookups = this.Singleton.Lookups
End Function



#If Win64 Then
  'Get an IAccessible object from a screen point
  '@param x - X coordinate (Horizontal)
  '@param y - Y coordinate (Vertical)
  '@returns - Accessibility object
  Private Function AccessibilityObjectFromPointEx(ByVal x as long, ByVal y as long) As IAccessible
    Dim o As IAccessible, v As Variant
    Call AccessibleObjectFromPoint(PointToLongLong(x,y), o, v)
    Set AccessibilityObjectFromPointEx = o
  End Function

  'Convert a Point's X (Long) and Y (Long) into a `LongLong`
  '@param x - X coordinate (Horizontal)
  '@param y - Y coordinate (Vertical)
  '@returns - LongLong representation of point
  '@remark You may notice that we are putting the arguments x and y into the function in the reverse order.
  '        I believe this is because stdCall convention puts arguments onto the stack in the reverse order, but in this case
  '        Since we are providing a single argument, we have to do that instead of the compiler. This ordering has been confirmed
  '        experimentally, this is just my hypothesis for this weird ordering.
  Private Function PointToLongLong(ByVal x as long, ByVal y as long) As LongLong
    'If sign bit is set (10000000 00000000 00000000 00000000) then shifting by 4 bytes left will cause an overflow.
    'Therefore exclude it (XOR part) but set the sign bit.
    Dim sign As LongLong
    If y And &H80000000 Then
      y = y Xor &H80000000
      sign = &H8000000000000000^
    End If
    PointToLongLong = (y * &H100000000^) Or (x And &HFFFFFFFF) Or sign
  End Function
#Else
  'Get an IAccessible object from a screen point
  '@param x - X coordinate (Horizontal)
  '@param y - Y coordinate (Vertical)
  '@returns - Accessibility object
  Private Function AccessibilityObjectFromPointEx(ByVal x as long, ByVal y as long) As IAccessible
    Dim o As IAccessible, v As Variant
    Call AccessibleObjectFromPoint(x, y, o, v)
    Set AccessibilityObjectFromPointEx = o
  End Function
#End If

''Creates a userform with transparent background and red border around the element. Useful for debugging.
''@TODO: Write this function
'Private Function Highlight() as stdUserform
'
'End Function

'Retrieve the IAccessible interface from a window handle
'@param {LongPtr} hwnd
'@returns {IAccessible} IAccessible interface for hwnd
Private Function IAccessibleFromHwnd(ByVal hwnd As LongPtr) As IAccessible
    ' Define the GUID for the IAccessible object {618736E0-3C3D-11CF-810C-00AA00389B71}
    Dim Guid As tGUID
    Guid = convertGUID("618736E0-3C3D-11CF-810C-00AA00389B71")
   
    ' Retrieve the IAccessible object from the window
    Dim oIA As IAccessible, lReturn As Long
    lReturn = AccessibleObjectFromWindow(hwnd, 0, Guid, oIA)
    Set IAccessibleFromHwnd = oIA
End Function

'Create a GUID struct from input string
'@param guid - GUID string to convert
'@returns - guid struct for use in API calls
Private Function convertGUID(Guid As String) As tGUID
    'Lookups defined by:
    ' Define the GUID for the IAccessible object
    ' {618736E0-3C3D-11CF-810C-00AA00389B71}
    ' With tg
    '    .lData1 = &H618736E0
    '    .nData2 = &H3C3D
    '    .nData3 = &H11CF
    '    .abytData4(0) = &H81
    '    .abytData4(1) = &HC
    '    .abytData4(2) = &H0
    '    .abytData4(3) = &HAA
    '    .abytData4(4) = &H0
    '    .abytData4(5) = &H38
    '    .abytData4(6) = &H9B
    '    .abytData4(7) = &H71
    ' End With
    
    Dim vArr: vArr = Split(Guid, "-")
    
    'Generate GUID structure from string of hex
    convertGUID.lData1       = CLng("&H" & vArr(0))             '618736E0
    convertGUID.nData2       = CInt("&H" & vArr(1))             '3C3D
    convertGUID.nData3       = CInt("&H" & vArr(2))             '11CF
    convertGUID.abytData4(0) = CInt("&H" & Mid(vArr(3), 1, 2))  '81
    convertGUID.abytData4(1) = CInt("&H" & Mid(vArr(3), 3, 2))  '0C
    convertGUID.abytData4(2) = CInt("&H" & Mid(vArr(4), 1, 2))  '00
    convertGUID.abytData4(3) = CInt("&H" & Mid(vArr(4), 3, 2))  'AA
    convertGUID.abytData4(4) = CInt("&H" & Mid(vArr(4), 5, 2))  '00
    convertGUID.abytData4(5) = CInt("&H" & Mid(vArr(4), 7, 2))  '38
    convertGUID.abytData4(6) = CInt("&H" & Mid(vArr(4), 9, 2))  '9B
    convertGUID.abytData4(7) = CInt("&H" & Mid(vArr(4), 11, 2)) '71
End Function

'Bind standard acc global variables to a callable
'@param stdICallable - A callable to bind global variables to (if the callable supports it)
Private Sub BindGlobals(ByRef query as stdICallable)
    Dim bSuccess as boolean
    Call query.SendMessage("bindGlobal", bSuccess, Array("EAccRoles", this.Singleton.Lookups("EAccRoles")("S2N")))
    Call query.SendMessage("bindGlobal", bSuccess, Array("EAccStates", this.Singleton.Lookups("EAccStates")("S2N")))
    Call query.SendMessage("bindGlobal", bSuccess, Array("EAccFindResult", this.Singleton.Lookups("EAccFindResult")("S2N")))
End Sub

'Creates a collection of stdAcc objects representing the children of the element
'@returns Collection<stdAcc> - The returned collection of children
Private Function getChildrenAsIAcc() As Collection
  'Get IAccessible children as Array of children
  Dim vChildren()
  ReDim vChildren(1 To protAccessible.accChildCount)
  Dim iChildrenObtained as long
  Call AccessibleChildren(protAccessible, 0, protAccessible.accChildCount, vChildren(1), iChildrenObtained)
  
  Set getChildrenAsIAcc = New Collection
  Dim child As Variant
  For iChildID = 1 To iChildrenObtained
    'Check if child is an object, if it is set it, otherwise let it
    if vartype(vChildren(iChildID)) = vbObject Then
      getChildrenAsIAcc.Add CreateFromIAccessible(vChildren(iChildID))
    Else
      getChildrenAsIAcc.Add createProxyChild(vChildren(iChildID))
    end if
  Next
End Function

'Creates a proxy element under this element (this element is parent). Proxy elements get there information from the parent directly.
'@param childID - Child ID to create proxy element on
'@returns - stdAcc element for child
Private Function createProxyChild(ByVal childID As Long) As stdAcc
    Set createProxyChild = New stdAcc
    Call createProxyChild.protInitWithProxy(Me, childID)
End Function

'Escape string for compatibility with JSON
'@param s - string to escape
'@returns - Escaped string
Private Function jsonEscape(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, """", "\""")
    jsonEscape = s
End Function

'Whenever we initialise ensure to get Lookups from `stdAcc`
Private Sub Class_Initialize()
    set this.Singleton.Lookups = stdAcc.protGetLookups()
End Sub

'From an array of Key-Value pairs, create a dictionary of Key-->Value and Value-->Key lookups
'@param arr as Array<Variant> - 1D array of Key,Value pairs
'@returns Object<Dictionary<S2N: Dictionary<key: string,value: variant>, N2S: Dictionary<value: variant, key: string>>> - Dictionary of dictionaries containing the lookups
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
        Call oRet("N2S").add(arr(i+1), arr(i))
    next

    'Return dictionary
    set CreateLookupDict = oRet
End Function

'****************************************
'* HELPERS FOR FindFirst() and FindAll()
'****************************************
'Constructor for a tFindNode type
'@param depth - depth of element
'@param element - element to store in tFindNode
'@returns - FindNode struct
Private Function CreateFindNode(ByVal depth as long, ByVal element as stdAcc) as tFindNode
    CreateFindNode.initialised = true
    CreateFindNode.depth = depth
    set CreateFindNode.element = element
End Function

'Given an array `stack`, and length `index`, append `item` onto the stack 
'@param stack as Array<tFindNode> - stack of elements
'@param index - current index in stack
'@param item - Item to insert into stack
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
'@param stack as Array<tFindNode> - stack of elements
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
    set stackPopV.element = stack(index).element
    stack(index).initialised = False
    stack(index).depth = 0
    Set stack(index).element = Nothing
End Function


'Returns the COM Error Description for the passed error code
'@param code - Error code
'@returns - Description of error
Private Function COMErrorDescription(ByVal code As Long) As String
    Static pErrorCodes As Object
    If pErrorCodes Is Nothing Then
        Set pErrorCodes = CreateObject("Scripting.Dictionary")
        pErrorCodes(Hex(-2147352563)) = "DISP_E_ARRAYISLOCKED"
        pErrorCodes(Hex(-2147352560)) = "DISP_E_BADCALLEE"
        pErrorCodes(Hex(-2147352565)) = "DISP_E_BADINDEX"
        pErrorCodes(Hex(-2147352562)) = "DISP_E_BADPARAMCOUNT"
        pErrorCodes(Hex(-2147352568)) = "DISP_E_BADVARTYPE"
        pErrorCodes(Hex(-2147352557)) = "DISP_E_BUFFERTOOSMALL"
        pErrorCodes(Hex(-2147352558)) = "DISP_E_DIVBYZERO"
        pErrorCodes(Hex(-2147352567)) = "DISP_E_EXCEPTION"
        pErrorCodes(Hex(-2147352573)) = "DISP_E_MEMBERNOTFOUND"
        pErrorCodes(Hex(-2147352569)) = "DISP_E_NONAMEDARGS"
        pErrorCodes(Hex(-2147352559)) = "DISP_E_NOTACOLLECTION"
        pErrorCodes(Hex(-2147352566)) = "DISP_E_OVERFLOW"
        pErrorCodes(Hex(-2147352572)) = "DISP_E_PARAMNOTFOUND"
        pErrorCodes(Hex(-2147352561)) = "DISP_E_PARAMNOTOPTIONAL"
        pErrorCodes(Hex(-2147352571)) = "DISP_E_TYPEMISMATCH"
        pErrorCodes(Hex(-2147352575)) = "DISP_E_UNKNOWNINTERFACE"
        pErrorCodes(Hex(-2147352564)) = "DISP_E_UNKNOWNLCID"
        pErrorCodes(Hex(-2147352570)) = "DISP_E_UNKNOWNNAME"
        pErrorCodes(Hex(-2147467260)) = "E_ABORT"
        pErrorCodes(Hex(-2147024891)) = "E_ACCESSDENIED"
        pErrorCodes(Hex(-2147467259)) = "E_FAIL"
        pErrorCodes(Hex(-2147024890)) = "E_HANDLE"
        pErrorCodes(Hex(-2147024809)) = "E_INVALIDARG"
        pErrorCodes(Hex(-2147467262)) = "E_NOINTERFACE"
        pErrorCodes(Hex(-2147467263)) = "E_NOTIMPL"
        pErrorCodes(Hex(-2147024882)) = "E_OUTOFMEMORY"
        pErrorCodes(Hex(-2147483638)) = "E_PENDING"
        pErrorCodes(Hex(-2147467261)) = "E_POINTER"
        pErrorCodes(Hex(-2147418113)) = "E_UNEXPECTED"
    End If
    COMErrorDescription = pErrorCodes(Hex(code))
End Function



Private Function accGetIdentity(ByRef accIdent As IUnknown) As String
  Const GetIdentityOffset As Long = 3
  Dim accIdentPtr As LongPtr
  Dim accIdentLen As Long
  Dim hResult As Long: hResult = DispCallFuncEx(accIdent, GetIdentityOffset, vbLong, 0, VarPtr(accIdentPtr), VarPtr(accIdentLen))
  If hResult = 0 Then
    Dim ptr As LongPtr
    Dim accIdentByt() As Byte
    ReDim accIdentByt(0 To accIdentLen)
    Call CopyMemory(VarPtr(accIdentByt(0)), accIdentPtr, accIdentLen)
    Call CoTaskMemFree(accIdentPtr)

    Dim s As String: s = Space(accIdentLen * 2)
    Dim i As Long: For i = 0 To accIdentLen - 1
      Mid(s, i * 2 + 1, 2) = IIf(accIdentByt(i) < 10, "0", "") & Hex(accIdentByt(i))
    Next

    accGetIdentity = s
  End If
End Function

Private Function unkQueryInterface(ByRef unk As IUnknown, ByVal sIID As String) As IUnknown
  Const IUnknownQueryInterface = 0
  Const S_OK          As Long = 0
  Const E_NOINTERFACE As Long = &H80004002
  Const E_POINTER     As Long = &H80004003

  Dim IInterface As IUnknown
  Dim tIID As GUID
  If IIDFromString(StrPtr(sIID), tIID) = S_OK Then
    On Error Resume Next
    Dim hResult As Long: hResult = DispCallFuncEx(unk, IUnknownQueryInterface, vbLong, VarPtr(tIID.Data1), VarPtr(IInterface))
    On Error GoTo 0
    If hResult = S_OK Then
      If IInterface Is Nothing Then
        Set unkQueryInterface = Nothing
      Else
        Set unkQueryInterface = IInterface
      End If
    ElseIf hResult = E_POINTER Then
      Set unkQueryInterface = Nothing
    ElseIf hResult = E_NOINTERFACE Then
      Err.Raise E_NOINTERFACE, "stdAcc", "IUnknown::QueryInterface - Interface not implemented."
    Else
      Err.Raise hResult, "stdAcc", "IUnknown::QueryInterface - Unknown error occurred."
    End If
  Else
    Err.Raise Err.LastDllError, "stdAcc", "IUnknown::QueryInterface - Cannot retrieve IID from String."
  End If
End Function

'Call a function in the VTable of the object at a specified offset
'@param unk - The object to call the function on
'@param VTableOffset - The offset of the method to call. IUnknown::QueryInterface=0; IUnknown::AddRef=1; IUnknown::Release=2; IDispatch::Invoke=6
'@param ReturnType - Return value of function, typically Long (HResult) for COM Object calls
'@param FunctionParameters - Arguments to send to function.
'@returns - Return value of function
Private Function DispCallFuncEx(ByRef unk As IUnknown, ByVal VTableOffset As Long, ByVal ReturnType As VbVarType, ParamArray FunctionParameters() As Variant) As Variant
    #If Win64 Then
      Const PTRSIZE = 8
    #Else
      Const PTRSIZE = 4
    #End If
    Const CC_STDCALL As Long = 4

    'Create array from params
    Dim vParam() As Variant: vParam = FunctionParameters
    Dim iParamCount As Long: iParamCount = UBound(vParam) - LBound(vParam) + 1

    'Create array of types and pointers to vars, for function call
    #If VBA7 Then
      Dim vPtr() As LongPtr
    #Else
      Dim vPtr() As Long
    #End If
    Dim vType() As Integer

    'Populate array of pointers to params, and array of types for DispCallFunc
    If iParamCount = 0 Then
      ReDim vType(0 To 0)
      ReDim vPtr(0 To 0)
    Else
      ReDim vType(0 To iParamCount - 1)
      ReDim vPtr(0 To iParamCount - 1)

      Dim lIdx As Long
      For lIdx = 0 To iParamCount - 1
        vType(lIdx) = VarType(vParam(lIdx))
        vPtr(lIdx) = VarPtr(vParam(lIdx))
      Next
    End If

    'Call function
    Dim ox As LongPtr: ox = ObjPtr(unk)
    Dim hResult As Long
    hResult = DispCallFunc(ox, VTableOffset * PTRSIZE, CC_STDCALL, ReturnType, iParamCount, VarPtr(vType(0)), VarPtr(vPtr(0)), vv)
    If IsObject(vv) Then
      Set DispCallFuncEx = vv
    Else
      Let DispCallFuncEx = vv
    End If
    If hResult < 0 Then
        Err.Raise hResult, "DispCallFunc", "Error in calling DispCallFunc"
    End If
End Function

