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


Private Declare PtrSafe Function AccessibleObjectFromPoint Lib "oleacc" (ByVal x As Long, ByVal y As Long, ppoleAcc As IAccessible, pvarElement As Variant) As Long
Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hwnd As LongPtr, ByVal dwId As Long, riid As tGUID, ppvObject As Object) As Long
Private Declare PtrSafe Function AccessibleObjectFromEvent Lib "oleacc.dll" (hwnd As LongPtr, dwObjectId As Integer, dwChildID As Integer, ppacc As IAccessible, pVarChild As Object) As Long
Private Declare PtrSafe Function AccessibleChildren Lib "oleacc" (ByVal paccContainer As IAccessible, ByVal iChildStart As Long, ByVal cChildren As Long, rgvarChildren As Variant, pcObtained As Long) As Long
Private Declare PtrSafe Function WindowFromAccessibleObject Lib "oleacc" (ByVal pacc As IAccessible, ByRef pHwnd As LongPtr) As Long
Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As tPOINT) As Long 'For FromMouse()
Private Declare PtrSafe Function APISendMessage Lib "user32" Alias "SendMessage" (ByVal hwnd As LongPtr, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


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

Private isProxy As Boolean
Private proxyParent As IAccessible 'IAcc
Private proxyIndex As Long
Private Lookups As Object

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


Const CHILDID_SELF = 0&

'The object which implements `IAccessible`, with which this class wraps. Please don't modify this object, unless you know what you are doing.
'@protected 
Public protAccessible As Object

'Constructors:
'Creates an `stdAcc` object from an `X` and `Y` point location on the screen.
'@constructor
'@param x {ByVal Long} X Coordinate
'@param y {ByVal Long} Y Coordinate
'@returns {stdAcc} Object at the specified point
Public Function CreateFromPoint(ByVal x As Long, ByVal y As Long) As stdAcc
    Set CreateFromPoint = New stdAcc
    Dim obj As IAccessible
    Dim v As Variant
    Call AccessibleObjectFromPoint(x, y, obj, v)
    set CreateFromPoint.protAccessible = obj
End Function

'Creates an `stdAcc` object from a window handle.
'@constructor
'@param hwnd {ByVal LongPtr} Window handle to wrap
'@returns {stdAcc} stdAcc object for the window 
Public Function CreateFromHwnd(ByVal hwnd As LongPtr) As stdAcc
    Dim acc as stdAcc: set acc = new stdAcc
    Dim obj As IAccessible
    Set obj = IAccessibleFromHwnd(hwnd)
    set acc.protAccessible = obj
    Set CreateFromHwnd = acc
End Function

'Creates an `stdAcc` object from the current running application (e.g. Excel / Word / Powerpoint).
'@constructor
'@returns {stdAcc}
Public Function CreateFromApplication() As stdAcc
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

'Creates an `stdAcc` object from the desktop.
'@constructor
'@returns {stdAcc}
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
'@param {IAccessible} Object implementing `IAccessible` interface.
'@constructor
'@returns {stdAcc}
Public Function CreateFromIAccessible(ByRef obj As IAccessible) As stdAcc
    Set CreateFromIAccessible = New stdAcc
    set CreateFromIAccessible.protAccessible = obj
End Function

'Creates an `stdAcc` object for the element the mouse currently hovers over.
'@constructor
'@returns {stdAcc}
Public Function CreateFromMouse() As stdAcc
    Dim pT As tPOINT
    Dim success As Long
    success = GetCursorPos(pT)
    Set CreateFromMouse = Me.CreateFromPoint(pT.x, pT.y)
End Function

'Creates an `stdAcc` object for the element at a given path from the current element.
'@constructor
'@returns {stdAcc}
'@example `Debug.Print stdAcc.CreateFromApplication().CreateFromPath("3.1").name`
Public Function CreateFromPath(ByVal sPath As String) As stdAcc
    'If starting with "." remove it
    If left(sPath, 1) = "." Then sPath = Mid(sPath, 2)
    
    'Get descendants list
    Dim descendants As Variant
    descendants = Split(sPath, ".")
    
    'Initiate acc (used for tracing through descendants)
    Dim acc As stdAcc
    Set acc = Me
    
    'Loop over descendants
    Dim i As Integer
    For i = 0 To UBound(descendants)
        If CLng(descendants(i)) > acc.children.Count Then Exit Function
        Set acc = acc.children(CLng(descendants(i)))
    Next i
    
    'Return descendant
    Set CreateFromPath = acc
End Function

'Initialises an stdAcc object as a `Proxy` object, who's methods are implemented on the parent instead of on the element itself
'@constructor
'@protected
'@returns {VOID}
Public Sub InitWithProxy(ByRef oParent as stdAcc, ByVal index as long)
    isProxy = true
    set proxyParent = oParent.protAccessible
    proxyIndex = index
End Sub


'Get all descendents of the stdAcc control
'@returns {Collection} Collection of descendents
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
'@param {ByVal stdICallable<(stdAcc,depth)=>EAccFindResult>} Callback returning
'  EAccFindResult options:
'    EAccFindResult.NoMatchFound/0/False             - Not found, countinue walking
'    EAccFindResult.MatchFound/1/-1/True             - Found, return this element
'    EAccFindResult.NoMatchCancelSearch/2            - Not found, cancel search
'    EAccFindResult.NoMatchSkipDescendents/3,else    - Not found, don't search descendents
'    EAccFindResult.MatchFoundSearchDescendents/4    - Same as EAccFindResult.MatchFound
'@param {EAccFindType=1} - The type of search, 0 for Breadth First Search (BFS) and 1 for Depth First Search (DFS).
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
'@param {ByVal stdICallable<(stdAcc,depth)=>EAccFindResult>} Callback returning
'  EAccFindResult options:
'    EAccFindResult.NoMatchFound/0/False             - Not found, countinue walking
'    EAccFindResult.MatchFound/1/-1/True             - Found, return this element, won't search descendents of elements found
'    EAccFindResult.NoMatchCancelSearch/2            - Not found, cancel search
'    EAccFindResult.NoMatchSkipDescendents/3,else    - Not found, don't search descendents
'    EAccFindResult.MatchFoundSearchDescendents/4    - Found, return this element, but continue searching descendents
'@param {EAccFindType} - The type of search, 0 for Breadth First Search (BFS) and 1 for Depth First Search (DFS).
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
'@returns {stdAcc} Parent of this element
Public Property Get Parent() As stdAcc
    If isProxy Then
        Set Parent = proxyParent
    else
        On Error GoTo handle_error
            Set Parent = New stdAcc
            set Parent.protAccessible = protAccessible.accParent
        On Error GoTo 0
    End if
    Exit Property
handle_error:
    Set Parent = Nothing
End Property

'Return the children of the IAccessible object
'@returns {Collection<stdAcc>} Children of this element
Public Property Get children() As Collection
    On Error GoTo ErrorHandler
        Set children = getChildrenAsIAcc()
    On Error GoTo 0
    Exit Property
ErrorHandler:
    Set children = New Collection
End Property

'Return the hwnd of the IAccessible object
'@returns {LongPtr} hwnd of this element
'Return the hwnd of the IAccessible object
'@returns {LongPtr} hwnd of this element
Public Property Get hwnd() As LongPtr
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

'Return the location of the element as a collection
'@returns {Collection<Long>} the location of the element. Has 5 named keys: "Width", "Height", "Left", "Top" and "Parent"
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
'@param {long} x coord
'@param {long} y coord
'@returns {stdAcc} element under point
Public Property Get HitTest(ByVal x As Long, ByVal y As Long) As stdAcc
  Set HitTest = New stdAcc
  Dim NewAcc As Object
  Call protAccessible.accHitTest(x, y, NewAcc)
  set HitTest.protAccessible = NewAcc
End Property

'Gets the value of the element
'@returns {String} the current value of the element
Public Property Get value() As String
  'On Error Resume Next
  if isProxy then
    value = proxyParent.accValue(proxyIndex)
  else
    value = protAccessible.accValue(CHILDID_SELF)
  end if
End Property

'Sets the value of the element
'@param {Variant} the new value of the element
Public Property Let value(val As String)
  'On Error Resume Next
  if isProxy then
    proxyParent.accValue(proxyIndex) = val
  else
    protAccessible.accValue(CHILDID_SELF) = val
  end if
End Property

'Returns the name of the element
'@returns {String} the name of the element
Public Property Get name() As String
  'On Error Resume Next
  If isProxy Then
    name = proxyParent.accName(proxyIndex)
  Else
    name = protAccessible.accName(CHILDID_SELF)
  End If
End Property

'Returns the description of the default action of the element
'@returns {String} the description of the default action of the element
Public Property Get DefaultAction() As String
  'On Error Resume Next
  If isProxy Then
    DefaultAction = proxyParent.accDefaultAction(proxyIndex)
  else
    DefaultAction = protAccessible.accDefaultAction(CHILDID_SELF)
  end if
End Property

'Returns the role of the element
'@returns {String} the role of the element
Public Property Get Role() As String
  'On Error Resume Next
  if isProxy then
    Role = Lookups("EAccRoles")("N2S")(proxyParent.accRole(proxyIndex))
  else
    Role = Lookups("EAccRoles")("N2S")(protAccessible.accRole(CHILDID_SELF))
  end if
End Property

'Returns the state of the element
'@returns {String} the state of the element
Public Property Get State() As String
  'On Error Resume Next
  if isProxy then
    State = Lookups("EAccStates")("N2S")(proxyParent.accState(proxyIndex))
  else
    State = Lookups("EAccStates")("N2S")(protAccessible.accState(CHILDID_SELF))
  end if
End Property

'Returns the state(s) of the element
'@returns {Long} the state(s) of the element
Public Property Get StateData() As Long
  'On Error Resume Next
  if isProxy then
    StateData = proxyParent.accState(proxyIndex)
  else
    StateData = protAccessible.accState(CHILDID_SELF)
  end if
End Property

'Returns the description of the element
'@returns {String} the description of the element
Public Property Get Description() As String
  'On Error Resume Next
  if isProxy then
    Description = proxyParent.accDescription(proxyIndex)
  else
    Description = protAccessible.accDescription(CHILDID_SELF)
  end if
End Property

'Returns the keyboard shortcut used to trigger the default action of the element
'@returns {String} The keyboard shortcut used to trigger the default action of the element
Public Property Get KeyboardShortcut() As String
  'On Error Resume Next
  if isProxy then
    KeyboardShortcut = proxyParent.accKeyboardShortcut(proxyIndex)
  else
    KeyboardShortcut = protAccessible.accKeyboardShortcut(CHILDID_SELF)
  end if
End Property

'Get the Focus of the UI element
'@returns {Boolean} Whether the element has focus
Public Property Get Focus() As Boolean
  'On Error Resume Next
  if isProxy then
    Focus = proxyParent.accFocus(proxyIndex)
  else
    Focus = protAccessible.accFocus(CHILDID_SELF)
  end if
End Property

'Set the Focus of the UI element
'@param {Boolean} True if the element should be focussed, otherwise false
Public Property Let Focus(ByVal val As Boolean)
  'On Error Resume Next
  if isProxy then
    proxyParent.accFocus(proxyIndex) = val
  else
    protAccessible.accFocus(CHILDID_SELF) = val
  end if
End Property

'Returns the help
'@returns {String} the help
Public Property Get Help() As String
  'On Error Resume Next
  if isProxy then
    Help = proxyParent.accHelp(proxyIndex)
  else
    Help = protAccessible.accHelp(CHILDID_SELF)
  end if
End Property

'Returns the help topic
'@returns {String} the help topic
Public Property Get HelpTopic(Optional ByVal sFile as string = "") As String
  'On Error Resume Next
  if isProxy then
    HelpTopic = proxyParent.accHelpTopic(sFile, proxyIndex)
  else
    HelpTopic = protAccessible.accHelpTopic(sFile, CHILDID_SELF)
  end if
End Property

'Returns a string containing numerous properties from stdAcc concatenated together.
'@returns {String} Concatenation of Name, Value, DefaultAction, Description, Role, Help, HelpTopic and KeyboardShortcut.
Public Property Get Text() As String
  If Len(Me.name & Me.value & Me.DefaultAction & Me.Description & Me.Help & Me.HelpTopic & Me.KeyboardShortcut) > 0 Then
    Text = "Name: """ & Me.name & """; " & _
            "Value: """ & Me.value & """; " & _
            "DefaultAction: """ & Me.DefaultAction & """; " & _
            "Description: """ & Me.Description & """; " & _
            "Role: """ & Me.Role & """; " & _
            "Help: """ & Me.Help & """; " & _
            "HelpTopic: """ & Me.HelpTopic & """; " & _
            "KeyboardShortcut: """ & Me.KeyboardShortcut & """;"
  Else
    Text = ""
  End If
End Property

'Performs the default action of the IAccessible object
'@returns {VOID}
Public Sub DoDefaultAction()
  If isProxy Then
    Call proxyParent.accDoDefaultAction(proxyIndex)
  Else
    Call protAccessible.accDoDefaultAction(CHILDID_SELF)
  End If
End Sub

'Sends a message to the window
'TODO: Move this out to stdWindow when ready `stdWindow.Create(hwnd).sendMessage(...)`
'@param {Long} Message to send
'@param {Long} WParam to send
'@param {Long} LParam to send
'@returns {Long} Return value
Public Function SendMessage(ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Me.hwnd > 0 Then
        SendMessage = APISendMessage(Me.hwnd, Msg, wParam, lParam)
    Else
        Err.Raise 1, "No hwnd on this window"
    End If
End Function

'Print the text of all child elements
'@returns {VOID}
Public Sub PrintChildTexts()
    Dim child As stdAcc
    Dim iIndex As Long: iIndex = 0
    For Each child In children
        iIndex = iIndex + 1
        Debug.Print iIndex & ". " & child.Text
    Next
End Sub

'Prints all descendent texts. Useful for debugging. Also includes all paths of all descendents which can be very beneficial.
'@param {ByVal String = ""} Optional File path to dump text to
'@param {ByVal Boolean = True} Optional whether to print output to debug window or not.
'@param {ByVal String = "P"} Optional prefix path (sometimes you might wnat to use Me.getPath(...))
'@param {ByVal Long = 0} Filenum to output too, typically this will be left unprovided
'@returns {VOID}
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
'@param {stdAcc} Stop creating path at this element
'@returns {String} String representing path to element from either Desktop or toAccessible
Public Function getPath(Optional toAccessible As stdAcc = Nothing) As String
    On Error GoTo ErrorOccurred
        'Initialise trace
        Dim acc As stdAcc
        Set acc = Me
        
        'Collection to store path
        Dim col As Collection: set col = new collection
            
        'Collect parents
        While Not acc.Parent.name = "Desktop"
            Dim child As stdAcc
            Dim index As Long: index = 0
            Dim savedIndex As Long: savedIndex = 0
            For Each child In acc.Parent.children
                index = index + 1
                If child.hwnd = acc.hwnd And child.Role = acc.Role Then
                    savedIndex = index
                    Exit For
                End If
            Next
            
            'Add index to stack
            If savedIndex > 0 Then
                col.Add savedIndex
            Else
                'In some scenarios this has occurred where acc tree has changed after initial pass. Not much we can do if this is the case.
                Call Err.Raise(1, "", "Cannot find path because acc ancestor's parent-child relationship cannot be distinguished")
                getPath = ""
                Exit Function
            End If
                    
            'Elevate parent
            Set acc = acc.Parent
        Wend
        
        'Create path
        Dim path As String
        Dim i As Integer
        For i = col.count To 1 Step -1
            path = path & "." & col(i)
        Next i
        
        'Return path
        getPath = "D.W" & path
    On Error GoTo 0
    Exit Function
ErrorOccurred:
    Err.Raise 1, "stdAcc#getPath", "An unknown error occurred while finding stdAcc path." & vbCrLf & Err.Description
End Function

'Returns this element and all descendents as a JSON string. Useful for debugging
'@returns {String} JSON representation of element tree
Public Function toJSON() As String
    'Get children string
    Dim sChildren As String
    sChildren = ""
    Dim acc As stdAcc
    For Each acc In Me.children
        sChildren = sChildren & "," & acc.toJSON()
    Next
    sChildren = Mid(sChildren, 2)
    
    'Convert to json
    toJSON = "{" & _
        "name:""" & jsonEscape(Me.name) & """," & _
        "desc:""" & jsonEscape(Me.Description) & """," & _
        "value:""" & jsonEscape(Me.value) & """," & _
        "role:""" & jsonEscape(Me.Role) & """," & _
        "state:""" & jsonEscape(Me.State) & """," & _
        "ks:""" & jsonEscape(Me.KeyboardShortcut) & """," & _
        "help:""" & jsonEscape(Me.Help) & """," & _
        "ht:""" & jsonEscape(Me.HelpTopic) & """," & _
        "children:[" & sChildren & "]}"
End Function

'Returns the lookups object
'@returns {Dictionary<Dictionary<Dictionary<string|long>>>}
Public Function protGetLookups()
    if Lookups is nothing then
        set Lookups = CreateObject("Scripting.Dictionary")
        set Lookups("EAccStates") = CreateLookupDict(Array( _ 
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
        set Lookups("EAccRoles") = CreateLookupDict(Array( _ 
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
        set Lookups("EAccFindResult") = CreateLookupDict(Array( _ 
            "MatchFound", EAccFindResult.MatchFound, _ 
            "MatchFoundSearchDescendents", EAccFindResult.MatchFoundSearchDescendents, _ 
            "NoMatchFound", EAccFindResult.NoMatchFound, _ 
            "NoMatchCancelSearch", EAccFindResult.NoMatchCancelSearch, _ 
            "NoMatchSkipDescendents", EAccFindResult.NoMatchSkipDescendents _ 
        ))
    end if
    set protGetLookups = lookups
End Function



'TODO: Highlight() - Creates a userform with transparent background and red border
'Public Function Highlight() as stdUserform
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
'@param {String} GUID to convert
'@returns {tGUID} for use in API calls
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

'@param {stdICallable} A callable that may 
'@returns {VOID}
Private Sub BindGlobals(ByRef query as stdICallable)
    Dim bSuccess as boolean
    Call query.SendMessage("bindGlobal", bSuccess, Array("EAccRoles", Lookups("EAccRoles")("S2N")))
    Call query.SendMessage("bindGlobal", bSuccess, Array("EAccStates", Lookups("EAccStates")("S2N")))
    Call query.SendMessage("bindGlobal", bSuccess, Array("EAccFindResult", Lookups("EAccFindResult")("S2N")))
End Sub

'Creates a collection of stdAcc objects representing the children of the element
'@returns {Collection<stdAcc>} The returned collection of children
Private Function getChildrenAsIAcc() As Collection
    Dim children As Collection
    Set children = getAccChildren(protAccessible)
    
    Set getChildrenAsIAcc = New Collection
    Dim child As Variant
    For iChildID = 1 To protAccessible.accChildCount
        'Check if child is an object, if it is set it, otherwise let it
        if vartype(children(iChildID)) = vbObject Then
            Set child = children(iChildID)
        Else
            child = children(iChildID)
        end if

        'Create proxy as required
        If VarType(child) = vbLong then
            getChildrenAsIAcc.Add createProxyChild(iChildID)
        Else
            Dim oAccessible As IAccessible
            Set oAccessible = child
            getChildrenAsIAcc.Add Me.CreateFromIAccessible(oAccessible)
        End If
    Next
End Function

'Creates a proxy element under this element (this element is parent). Proxy elements get there information from the parent directly.
'@param {Long} Child ID to create proxy element on
'@returns {stdAcc} stdAcc element for child
Private Function createProxyChild(ByVal childID As Long) As stdAcc
    Set createProxyChild = New stdAcc
    Call createProxyChild.InitWithProxy(Me, childID)
End Function

'Get children as IAccessible
'@param {Long} Child ID to create proxy element on
'@returns {Collection<IAccessible|VT_I4>} Collection of child elements
Private Function getAccChildren(obj As IAccessible) As Collection
    Dim arr As Variant
    ReDim arr(1 To obj.accChildCount)
    
    Dim d
    Call AccessibleChildren(obj, 0, obj.accChildCount, arr(1), d)
    
    Dim col As Collection: set col = new collection
    Dim i As Integer
    For i = 1 To d    'note d is used not accChildCount, because sometimes AccessibleChildren does not return all children (no explanation for this?)
        col.Add arr(i)
    Next
    
    Set getAccChildren = col
End Function

'Escape string for compatibility with JSON
'@param {String} string to escape
'@returns {String} Escaped string
Private Function jsonEscape(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, """", "\""")
    jsonEscape = s
End Function

'Whenever we initialise ensure to get lookups from `stdAcc`
Private Sub Class_Initialize()
    set Lookups = stdAcc.protGetLookups()
End Sub

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
        Call oRet("N2S").add(arr(i+1), arr(i))
    next

    'Return dictionary
    set CreateLookupDict = oRet
End Function

'****************************************
'* HELPERS FOR FindFirst() and FindAll()
'****************************************
'Constructor for a tFindNode type
'@param {Long} depth of element
'@param {stdAcc} element
'@returns {tFindNode} FindNode struct
Private Function CreateFindNode(ByVal depth as long, ByVal element as stdAcc) as tFindNode
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


'Returns the COM Error Description for the passed error code
'@param {Long} Error code
'@returns {String} Description of error
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
