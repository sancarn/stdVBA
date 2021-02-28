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
'INSTANCE METHODS:
'    [X] CreateFromPath
'    [X] GetDescendents
'    [X] FindFirst(stdICallable)
'    [X] FindAll(stdICallable)
'    [X] DoDefaultAction()
'    [X] SendMessage         TODO: Consider removing when we have stdWindow?
'    [X] Highlight(seconds?)
'    [X] getPath
'    [X] PrintChildTexts()
'    [X] PrintDescTexts(sFile?)
'PROPERTIES:
'    [X] R   Parent
'    [X] R   Children
'    [X] R   hwnd
'    [X] R   location
'    [X] R   HitTest
'    [X] R   Selection
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
'
'INSTANCE METHODS
'Many methods were inspired by those in Ruby's Enumerable: https://ruby-doc.org/core-2.7.2/Enumerable.html






'TODO:
'[?] SendKeys() method
'[?] SendKeysRaw() method
'[?] Focus method


Private Declare PtrSafe Function AccessibleObjectFromPoint Lib "oleacc" (ByVal x As Long, ByVal y As Long, ppoleAcc As IAccessible, pvarElement As Variant) As Long
Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hwnd As LongPtr, ByVal dwId As Long, riid As tGUID, ppvObject As Object) As Long
Private Declare PtrSafe Function AccessibleObjectFromEvent Lib "oleacc.dll" (hwnd As LongPtr, dwObjectId As Integer, dwChildID As Integer, ppacc As IAccessible, pVarChild As Object) As Long
Private Declare PtrSafe Function AccessibleChildren Lib "oleacc" (ByVal paccContainer As IAccessible, ByVal iChildStart As Long, ByVal cChildren As Long, rgvarChildren As Variant, pcObtained As Long) As Long
Private Declare PtrSafe Function WindowFromAccessibleObject Lib "oleacc" (ByVal pacc As IAccessible, ByRef pHwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As tPOINT) As Long 'For FromMouse()
Private Declare PtrSafe Function APISendMessage Lib "user32" Alias "SendMessage" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Enum ESearchType
    BredthFirst = 0
    DepthFirst = 1
End Enum

Public Enum EMSAARoles
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
Public Enum EMSAAStates
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

'Enums
Public proxy As Boolean
Public proxyParent As Object 'IAcc
Public proxyInfo As Object
Public proxyIndex As Long
Public Lookups As Object

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

'Should only be set from IAcc
Public protAccessible As Object

'Constructors:
Public Function CreateFromPoint(x As Long, y As Long) As stdAcc
    Set CreateFromPoint = New stdAcc
    Dim obj As IAccessible
    Dim v As Variant
    Call AccessibleObjectFromPoint(x, y, obj, v)
    set CreateFromPoint.protAccessible = obj
End Function
Public Function CreateFromHwnd(hwnd As LongPtr) As stdAcc
    Dim acc as New stdAcc
    Dim obj As IAccessible
    Set obj = IAccessibleFromHwnd(hwnd)
    set acc.protAccessible = obj
    Set CreateFromHwnd = acc
End Function
Public Function CreateFromApplication() As stdAcc
    Set CreateFromApplication = CreateFromHwnd(Application.hwnd)
End Function
Public Function CreateFromDesktop() As stdAcc
    'Get this application's accessibility object
    Dim accThis As IAccessible
    Set accThis = IAccessibleFromHwnd(Application.hwnd)
    
    'Set desktop
    Dim accDesktop As IAccessible
    Set accDesktop = accThis.accParent
    
    Dim acc As New stdAcc
    set acc.protAccessible = accDesktop
    Set CreateFromDesktop = acc
End Function
Public Function CreateFromIAccessible(ByRef obj As IAccessible) As stdAcc
    Set CreateFromIAccessible = New stdAcc
    set CreateFromIAccessible.protAccessible = obj
End Function
Public Function CreateFromMouse() As stdAcc
    Dim pT As tPOINT
    Dim success As Long
    success = GetCursorPos(pT)
    Set CreateFromMouse = Me.FromPoint(pT.x, pT.y)
End Function
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
        Set acc = acc.children(CLng(descendants(i)))
    Next i
    
    'Return descendant
    Set CreateFromPath = acc
End Function

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

'Search the IAccessible tree for elements which match a certain criteria. When found exit the function.
'@param {ByVal stdICallable<(stdAcc,eSearchType),eWalkType>} Callback returning
'  eWalkType options:
'    0/False   - Not found, countinue walking
'    1/-1/True - Found, return this element
'    2         - Not found, cancel search
'    else      - Not found, don't search descendents
'@param {eSearchType} - The type of search, 0 for Breadth First Search (BFS) and 1 for Depth First Search (DFS).
' To understand the difference between BFS and DFS take this tree:
'        A
'       / \
'      B   C
'     /   / \
'    D   E   F
' A BFS will walk this tree in the following order: A, B, C, D, E, F
' A DFS will walk this tree in a different order:   A, C, F, E, B, D
' 
'@examples
' ```
' 'Find where name is "hello" and class is "world":
' el.FindFirst(stdLambda.Create("$1.name=""hello"" and $1.class=""world"""))
' 'Find where name is "hello" but restrict search to 4 levels deep:
' el.FindFirst(stdLambda.Create("if $2 = 4 then 2 else $1.name=""hello"""))
' ```
Public Function FindFirst(ByVal query As stdICallable, optional byval searchType as ESearchType=0) As stdAcc
    Dim stack() As tFindNode
    ReDim stack(0 To 0)
    stack(0).initialised = true
    stack(0).depth = 0
    Set stack(0).element = Me
    
    Dim length As Long: length = 1
    Dim index As Long: index = -1

    'Loop over the stack/array
    While length > 0 And index < length
        Dim part As tFindNode
        If searchType = 0 Then
            'Bredth first search, get item directly out of array, no need to change array size
            index = index + 1
            part = stack(index)
        Else
            'Depth first search, so pop the item out of the stack
            part = stackPopV(stack, length)
        End If
        
        With part
            If not .initialised Then Exit Function
            
            'Run query and test options
            Select Case query.Run(.element, .depth)
                Case 0
                    'Nothing found, search descendents
                    Dim child As stdAcc
                    For Each child In part.element.children
                        Call stackPushV(stack, length, CreateFindNode(.depth + 1, child))
                    Next
                Case 1, True
                    'Found, return element
                    Set FindFirst = .element
                    Exit Function
                Case 2
                    'Nothing found, cancel function
                    Set FindFirst = Nothing
                    Exit Function
                'case 3
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
'@param {ByVal stdICallable<(stdAcc,eSearchType),eWalkType>} Callback returning
'  eWalkType options:
'    0/False   - Not found, countinue walking
'    1/-1/True - Found, return this element
'    2         - Not found, cancel search
'    else      - Not found, don't search descendents
'@param {eSearchType} - The type of search, 0 for Breadth First Search (BFS) and 1 for Depth First Search (DFS).
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
' 'Find where name is "hello" but restrict search to 4 levels deep:
' el.FindAll(stdLambda.Create("if $2 = 4 then 2 else $1.name=""hello"""))
' ```
Public Function FindAll(ByVal query As stdICallable, optional byval searchType as ESearchType=0) As Collection
    Dim stack() As tFindNode
    ReDim stack(0 To 0)
    stack(0).initialised = true
    stack(0).depth = 0
    Set stack(0).element = Me
    
    Dim length As Long: length = 1
    Dim index As Long: index = -1
    

    'Initialise collection
    set FindAll = new Collection

    'Loop over the stack/array
    While length > 0 And index < length
        Dim part as tFindNode
        if searchType = 0 then 
            'Bredth first search, get item directly out of array, no need to change array size
            index = index + 1
            part = stack(index)
        else
            'Depth first search, so pop the item out of the stack
            part = stackPopV(stack,length)
        end if
        
        With part
            'When hitting the edge of the stack quit
            if not .initialised then Exit Function

            'Run query and test options
            select case query.Run(.element,.depth)
                case 0
                    'Nothing found, search descendents
                    Dim child as stdAcc
                    for each child in .element.children
                        Call stackPushV(stack,length,CreateFindNode(.depth+1,child))
                    next
                case 1,True
                    'Found, add element
                    Call FindAll.add(.element)
                case 2
                    'Nothing found, cancel function
                    Exit Function
                'case 3
                '    Nothing found, don't search descendents
            end select
        End with
        'Just make sure no freezing occurs
        DoEvents
    Wend
End Function

'Complex properties
Public Property Get Parent() As stdAcc
    If Me.proxy Then
        Set Parent = Me.proxyParent
    End If
    On Error GoTo handle_error
        Set Parent = New stdAcc
        set Parent.protAccessible = protAccessible.accParent
    On Error GoTo 0
    Exit Property
handle_error:
    Set Parent = Nothing
End Property

Public Property Get children() As Collection
    On Error GoTo ErrorHandler
        Set children = getChildrenAsIAcc()
    On Error GoTo 0
    Exit Property
ErrorHandler:
    Set children = New Collection
End Property

Public Property Get hwnd()
    On Error GoTo handle_error
        Dim lHwnd As LongPtr
        WindowFromAccessibleObject protAccessible, lHwnd
        hwnd = lHwnd
    On Error GoTo 0
    Exit Property
handle_error:
    hwnd = 0
End Property

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

Public Property Get HitTest(x As Long, y As Long) As stdAcc
    Set HitTest = New stdAcc
    Dim NewAcc As Object
    Call protAccessible.accHitTest(x, y, NewAcc)
    set HitTest.protAccessible = NewAcc
End Property

Public Property Get Selection() As Object
    On Error Resume Next
    Set Selection = protAccessible.Selection
End Property

Public Property Get value() As Variant
    On Error Resume Next
    With protAccessible
        If VarType(.accValue) = vbObject Then
            Set value = .accValue
        Else
            Let value = .accValue
        End If
    End With
End Property
Public Property Let value(val As Variant)
    On Error Resume Next
    With protAccessible
        If VarType(.accValue) = vbObject Then
            Set .accValue = val
        Else
            Let .accValue = val
        End If
    End with
End Property



'Simple properties:
Public Property Get name() As String
    On Error Resume Next
    name = protAccessible.accName
End Property
Public Property Get DefaultAction() As String
    On Error Resume Next
    DefaultAction = protAccessible.accDefaultAction
End Property
Public Property Get Role() As String
    On Error Resume Next
    Role = Me.Lookups("roles")("S2N")(protAccessible.AccRole)
End Property
Public Property Get State() As String
    On Error Resume Next
    State = Me.Lookups("states")("S2N")(protAccessible.AccState)
End Property
Public Property Get StateData() As Long
    On Error Resume Next
    StateData = protAccessible.AccState
End Property
Public Property Get Description() As String
    On Error Resume Next
    Description = protAccessible.accDescription
End Property
Public Property Get KeyboardShortcut() As String
    On Error Resume Next
    KeyboardShortcut = protAccessible.accKeyboardShortcut
End Property
Public Property Get Focus() As Boolean
    On Error Resume Next
    Focus = protAccessible.accFocus
End Property
Public Property Let Focus(val As Boolean)
    On Error Resume Next
    protAccessible.accFocus = val
End Property
Public Property Get Help() As String
    On Error Resume Next
    Help = protAccessible.accHelp
End Property
Public Property Get HelpTopic() As String
    On Error Resume Next
    HelpTopic = protAccessible.accHelpTopic
End Property

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

Public Function DoDefaultAction()
    If Me.proxy Then
    Else
        Call protAccessible.accDoDefaultAction
    End If
End Function

Public Function SendMessage(ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Me.hwnd > 0 Then
        SendMessage = APISendMessage(Me.hwnd, Msg, wParam, lParam)
    Else
        Err.Raise 1, "No hwnd on this window"
    End If
End Function

'Useful for debugging
Public Sub PrintChildTexts()
    Dim child As stdAcc
    Dim iIndex As Long: iIndex = 0
    For Each child In children
        iIndex = iIndex + 1
        Debug.Print iIndex & ". " & child.Text
    Next
End Sub

'Prints all descendent texts. Useful for debugging. Also includes all paths of all descendents which can be very beneficial.
'
'
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



'TODO: Highlight() - Creates a userform with transparent background and red border
'Public Function Highlight() as stdUserform
'
'End Function

'IAccessible:
'  accName              - Get name of control or window
'  accDefaultAction     - Get default action
'  accDoDefaultAction   - Do default action
'  accRole              - Get role of control or window
'  accState             - Get state of control or window
'  accParent            - Get parent IAccessible object
'  accDescription       - Get description of control or window
'  accKeyboardShortcut  - Get keyboard shortcut if it exists
'  accChildCount        - Get count of children
'  accFocus             - Get Focus state
'  accLocation          - Get Position of the object on the screen
'  accHitTest           - Subobject under the mouse cursor
'  accSelection         - List of selected objects
'  accSelect            - Modify selection of objects
'  accHelp              - Get help text if provided
'  accHelpTopic         - Get help file if provided

'Retrieve the IAccessible interface from a window handle
Private Function IAccessibleFromHwnd(ByVal hwnd As LongPtr) As IAccessible
    ' Define the GUID for the IAccessible object {618736E0-3C3D-11CF-810C-00AA00389B71}
    Dim Guid As tGUID
    Guid = convertGUID("618736E0-3C3D-11CF-810C-00AA00389B71")
   
    ' Retrieve the IAccessible object from the window
    Dim oIA As IAccessible, lReturn As Long
    lReturn = AccessibleObjectFromWindow(hwnd, 0, Guid, oIA)
    Set IAccessibleFromHwnd = oIA
End Function

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


'  accName              - Get name of control or window
'  accDefaultAction     - Get default action
'  accDoDefaultAction   - Do default action
'  accRole              - Get role of control or window
'  accState             - Get state of control or window
'  accParent            - Get parent IAccessible object
'  accDescription       - Get description of control or window
'  accKeyboardShortcut  - Get keyboard shortcut if it exists
'  accChildCount        - Get count of children
'  accFocus             - Get Focus state
'  accLocation          - Get Position of the object on the screen
'  accHitTest           - Subobject under the mouse cursor
'  accSelection         - List of selected objects
'  accSelect            - Modify selection of objects
'  accHelp              - Get help text if provided
'  accHelpTopic         - Get help file if provided
Private Function createProxyChild(ByVal childID As Long) As stdAcc
    Dim acc As stdAcc
    Set acc = New stdAcc
    acc.proxy = True
    
    On Error Resume Next
    Set acc.Parent = Me
    acc.proxyIndex = childID
    acc.proxyInfo = CreateObject("Scripting.Dictionary")
    acc.proxyInfo.Add "Name", protAccessible.accName(childID - 1)
    acc.proxyInfo.Add "DefaultAction", protAccessible.accDefaultAction(childID - 1)
    acc.proxyInfo.Add "Role", protAccessible.AccRole(childID - 1)
    acc.proxyInfo.Add "State", protAccessible.AccState(childID - 1)
    acc.proxyInfo.Add "Parent", protAccessible.accParent(childID - 1)
    acc.proxyInfo.Add "Description", protAccessible.accDescription(childID - 1)
    acc.proxyInfo.Add "KeyboardShortcut", protAccessible.accKeyboardShortcut(childID - 1)
    acc.proxyInfo.Add "ChildCount", protAccessible.accChildCount(childID - 1)
    acc.proxyInfo.Add "Focus", protAccessible.accFocus(childID - 1)
    acc.proxyInfo.Add "Location", protAccessible.accLocation(childID - 1)
    acc.proxyInfo.Add "HitTest", protAccessible.accHitTest(childID - 1)
    acc.proxyInfo.Add "Selection", protAccessible.accSelection(childID - 1)
    acc.proxyInfo.Add "Help", protAccessible.accHelp(childID - 1)
    acc.proxyInfo.Add "HelpTopic", protAccessible.accHelpTopic(childID - 1)
    
    Set createProxyChild = acc
End Function

Private Function getAccChildren(obj As IAccessible) As Collection
    Dim arr As Variant
    ReDim arr(1 To obj.accChildCount)
    
    Dim d
    Call AccessibleChildren(obj, 0, obj.accChildCount, arr(1), d)
    
    Dim col As New Collection
    Dim i As Integer
    For i = 1 To d    'note d is used not accChildCount, because sometimes AccessibleChildren does not return all children (no explanation for this?)
        col.Add arr(i)
    Next
    
    Set getAccChildren = col
End Function

Public Function getPath(Optional toAccessible As stdAcc = Nothing) As String
    'Initialise trace
    Dim acc As stdAcc
    Set acc = Me
    
    'Collection to store path
    Dim col As New Collection
    
    
    'Collect parents
    While Not acc.Parent.name = "Desktop"
        Dim child As stdAcc
        Dim index As Long
        index = 0
        For Each child In acc.Parent.children
            index = index + 1
            If child.hwnd = acc.hwnd And child.Role = acc.Role Then
                Exit For
            End If
        Next
        
        'Add index to stack
        col.Add index
                
        'Elevate parent
        Set acc = acc.Parent
    Wend
    
    'Create path
    Dim path As String
    Dim i As Integer
    For i = col.Count To 1 Step -1
        path = path & "." & col(i)
    Next i
    
    'Return path
    getPath = "D.W" & path
End Function


'  accName              - Get name of control or window
'  accDefaultAction     - Get default action
'  accDoDefaultAction   - Do default action
'  accRole              - Get role of control or window
'  accState             - Get state of control or window
'  accParent            - Get parent IAccessible object
'  accDescription       - Get description of control or window
'  accKeyboardShortcut  - Get keyboard shortcut if it exists
'  accChildCount        - Get count of children
'  accFocus             - Get Focus state
'  accLocation          - Get Position of the object on the screen
'  accHitTest           - Subobject under the mouse cursor
'  accSelection         - List of selected objects
'  accSelect            - Modify selection of objects
'  accHelp              - Get help text if provided
'  accHelpTopic         - Get help file if provided
Function toJSON() As String
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

Private Function jsonEscape(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, """", "\""")
    jsonEscape = s
End Function

Private Sub Class_Initialize()
    set Lookups = stdAcc.getLookups()
End Sub

'Returns the lookups object
Public Function getLookups()
    if Lookups is nothing then
        set Lookups = CreateObject("Scripting.Dictionary")
        set Lookups("states") = CreateLookupDict(Array( _ 
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
        set Lookups("roles") = CreateLookupDict(Array( _ 
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
    end if
    set getLookups = lookups
End Function

Private Function CreateLookupDict(arr as Variant) as Object
    Dim oRet as object: set oRet = CreateObject("Scripting.Dictionary")
    Dim i as long

    'Create string --> number dictionary
    set oRet("S2N") = CreateObject("Scripting.Dictionary")
    For i = lbound(arr) to ubound(arr) step 2
        Call oRet("S2N").add(arr(i), arr(i+1))
    next

    'Create number --> string dictionary
    set oRet("N2S") = CreateObject("Scripting.Dictionary")
    For i = lbound(arr) to ubound(arr) step 2
        Call oRet("S2N").add(arr(i+1), arr(i))
    next

    'Return dictionary
    set CreateLookupDict = oRet
End Function




'****************************************
'* HELPERS FOR FindFirst() and FindAll()
'****************************************
Private Function CreateFindNode(ByVal depth as long, ByVal element as stdAcc) as tFindNode
    CreateFindNode.initialised = true
    CreateFindNode.depth = depth
    set CreateFindNode.element = element
End Function

'Given an array `stack`, and length `index`, append `item` onto the stack 
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

'Used in DFS
Private Function stackPopV(ByRef stack() As tFindNode, ByRef index As Variant) As tFindNode
    Dim size As Long: size = UBound(stack) + 1
    If index < size / 3 And index > minStackSize Then
        ReDim Preserve stack(0 To CLng(size / 2))
    End If
    index = index - 1
    
    stackPopV.initialised = stack(index).initialised
    stackPopV.depth = stack(index).depth
    set stackPopV.element = stack(index).element
    stack(index) = Empty
End Function
