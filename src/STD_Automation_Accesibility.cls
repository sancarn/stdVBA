VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "STD_Automation_Accesibility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare PtrSafe Function AccessibleObjectFromPoint Lib "oleacc" (ByVal x As Long, ByVal y As Long, ppoleAcc As IAccessible, pvarElement As Variant) As Long
Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hwnd As Long, ByVal dwId As Long, riid As tGUID, ppvObject As Object) As Long
Private Declare PtrSafe Function AccessibleObjectFromEvent Lib "oleacc.dll" (hwnd As LongPtr, dwObjectId As Integer, dwChildID As Integer, ppacc As IAccessible, pVarChild As Object) As Long
Private Declare PtrSafe Function AccessibleChildren Lib "oleacc" (ByVal paccContainer As IAccessible, ByVal iChildStart As Long, ByVal cChildren As Long, rgvarChildren As Variant, pcObtained As Long) As Long
Private Declare PtrSafe Function WindowFromAccessibleObject Lib "oleacc" (ByVal pacc As IAccessible, ByRef pHwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As tPOINT) As Long 'For FromMouse()

'Enums
Public ACC_STATES As Object
Public ACC_ROLES As Object
Public proxy As Boolean
Public proxyParent As IAcc
Public proxyInfo As Object
Public proxyIndex As Long

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

Const CHILDID_SELF = 0&

'Should only be set from IAcc
Private oAcc As Object

'Constructors:
Public Function FromPoint(x As Long, y As Long) As STD_Automation_Accesibility
    Set FromPoint = New STD_Automation_Accesibility
    Dim obj As IAccessible
    Dim v As Variant
    Call AccessibleObjectFromPoint(x, y, obj, v)
    Call FromPoint.SetOAcc(obj)
End Function
Public Function FromHwnd(hwnd As LongPtr) As STD_Automation_Accesibility
    Dim acc As New STD_Automation_Accesibility
    Dim obj As IAccessible
    Set obj = IAccessibleFromHwnd(hwnd)
    Call acc.SetOAcc(obj)
    Set FromHwnd = acc
End Function
Public Function FromExcel() As STD_Automation_Accesibility
    Set FromExcel = FromHwnd(Application.hwnd)
End Function
Public Function FromDesktop() As STD_Automation_Accesibility
    'Get this application's accessibility object
    Dim accThis As IAccessible
    Set accThis = IAccessibleFromHwnd(Application.hwnd)
    
    'Set desktop
    Dim accDesktop As IAccessible
    Set accDesktop = accThis.accParent
    
    Dim acc As New STD_Automation_Accesibility
    Call acc.SetOAcc(accDesktop)
    Set FromDesktop = acc
End Function
Public Function FromIAccessible(ByRef obj As IAccessible) As STD_Automation_Accesibility
    Set FromIAccessible = New STD_Automation_Accesibility
    Call FromIAccessible.SetOAcc(obj)
End Function
Public Function FromMouse() As STD_Automation_Accesibility
    Dim pt As tPOINT
    Dim success As Long
    success = GetCursorPos(pt)
    Set FromMouse = Me.FromPoint(pt.x, pt.y)
End Function
Public Function FromPath(ByVal sPath As String) As STD_Automation_Accesibility
    'If starting with "." remove it
    If Left(sPath, 1) = "." Then sPath = Mid(sPath, 2)
    
    'Get descendants list
    Dim descendants As Variant
    descendants = Split(sPath, ".")
    
    'Initiate acc (used for tracing through descendants)
    Dim acc As STD_Automation_Accesibility
    Set acc = Me
    
    'Loop over descendants
    Dim i As Integer
    For i = 0 To UBound(descendants)
        Set acc = acc.children(CLng(descendants(i)))
    Next i
    
    'Return descendant
    Set FromPath = acc
End Function


'Complex properties
Public Property Get parent() As STD_Automation_Accesibility
    If Me.proxy Then
        Set parent = Me.proxyParent
    End If
    On Error GoTo handle_error
        Set parent = New STD_Automation_Accesibility
        Call parent.SetOAcc(oAcc.accParent)
    On Error GoTo 0
    Exit Property
handle_error:
    Set parent = Nothing
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
        Dim lHwnd As Long
        WindowFromAccessibleObject oAcc, lHwnd
        hwnd = lHwnd
    On Error GoTo 0
    Exit Property
handle_error:
    hwnd = 0
End Property

Public Property Get Location() As Collection
    On Error GoTo ErrorHandler
        'Get location from oAcc
        Dim pcxWidth As Long
        Dim pcyHeight As Long
        Dim pxLeft As Long
        Dim pyTop As Long
        Call oAcc.accLocation(pxLeft, pyTop, pcxWidth, pcyHeight, CHILDID_SELF)
        
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

Public Property Get HitTest() As STD_Automation_Accesibility
    Set HitTest = New STD_Automation_Accesibility
    Call HitTest.SetOAcc(oAcc.accHitTest())
End Property

Public Property Get Selection() As Object
    On Error Resume Next
    Set Selection = oAcc.Selection
End Property

Public Property Get Value() As Variant
    On Error Resume Next
    If VarType(oAcc.accValue) = vbObject Then
        Set Value = oAcc.accValue
    Else
        Value = oAcc.accValue
    End If
End Property
Public Property Let Value(val As Variant)
    On Error Resume Next
    If VarType(oAcc.accValue) = vbObject Then
        Set oAcc.accValue = val
    Else
        oAcc.accValue = val
    End If
End Property



'Simple properties:
Public Property Get Name() As String
    On Error Resume Next
    Name = oAcc.accName
End Property
Public Property Get DefaultAction() As String
    On Error Resume Next
    DefaultAction = oAcc.accDefaultAction
End Property
Public Property Get Role() As String
    On Error Resume Next
    Role = Me.ACC_ROLES(oAcc.AccRole)
End Property
Public Property Get State() As String
    On Error Resume Next
    State = Me.ACC_STATES(oAcc.AccState)
End Property
Public Property Get Description() As String
    On Error Resume Next
    Description = oAcc.accDescription
End Property
Public Property Get KeyboardShortcut() As String
    On Error Resume Next
    KeyboardShortcut = oAcc.accKeyboardShortcut
End Property
Public Property Get Focus() As Boolean
    On Error Resume Next
    Focus = oAcc.accFocus
End Property
Public Property Let Focus(val As Boolean)
    On Error Resume Next
    oAcc.accFocus = val
End Property
Public Property Get Help() As String
    On Error Resume Next
    Help = oAcc.accHelp
End Property
Public Property Get HelpTopic() As String
    On Error Resume Next
    HelpTopic = oAcc.accHelpTopic
End Property


Public Function DoDefaultAction()
    If Me.proxy Then
    Else
        Call oAcc.accDoDefaultAction
    End If
End Function

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
Private Function IAccessibleFromHwnd(hwnd As Long) As IAccessible
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
    
    '618736E0-3C3D-11CF-810C-00AA00389B71
    convertGUID.lData1 = CLng(Application.WorksheetFunction.Hex2Dec(vArr(0))) '618736E0
    convertGUID.nData2 = CInt(Application.WorksheetFunction.Hex2Dec(vArr(1))) '3C3D
    convertGUID.nData3 = CInt(Application.WorksheetFunction.Hex2Dec(vArr(2))) '11CF
    convertGUID.abytData4(0) = CInt(Application.WorksheetFunction.Hex2Dec(Mid(vArr(3), 1, 2))) '81
    convertGUID.abytData4(1) = CInt(Application.WorksheetFunction.Hex2Dec(Mid(vArr(3), 3, 2))) '0C
    convertGUID.abytData4(2) = CInt(Application.WorksheetFunction.Hex2Dec(Mid(vArr(4), 1, 2))) '00
    convertGUID.abytData4(3) = CInt(Application.WorksheetFunction.Hex2Dec(Mid(vArr(4), 3, 2))) 'AA
    convertGUID.abytData4(4) = CInt(Application.WorksheetFunction.Hex2Dec(Mid(vArr(4), 5, 2))) '00
    convertGUID.abytData4(5) = CInt(Application.WorksheetFunction.Hex2Dec(Mid(vArr(4), 7, 2))) '38
    convertGUID.abytData4(6) = CInt(Application.WorksheetFunction.Hex2Dec(Mid(vArr(4), 9, 2))) '9B
    convertGUID.abytData4(7) = CInt(Application.WorksheetFunction.Hex2Dec(Mid(vArr(4), 11, 2))) '71
End Function

Private Function accDesktop() As IAccessible
    'Get this application's accessibility object
    Dim accThis As IAccessible
    Set accThis = IAccessibleFromHwnd(Application.hwnd)
    
    'Set desktop
    Set accDesktop = accThis.accParent
End Function

Private Function getChildrenAsIAcc() As Collection
    Dim children As Collection
    Set children = getAccChildren(oAcc)
    
    Set getChildrenAsIAcc = New Collection
    Dim child As Variant
    For Each child In children
        If VarType(child) = vbLong Then 'creates a proxy child to proxy all requests.
            getChildrenAsIAcc.Add createProxyChild(child)
        Else
            Dim oAccessible As IAccessible
            Set oAccessible = child
            getChildrenAsIAcc.Add Me.FromIAccessible(oAccessible)
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
Private Function createProxyChild(ByVal childID As Long) As STD_Automation_Accesibility
    Dim acc As New STD_Automation_Accesibility
    acc.proxy = True
    
    On Error Resume Next
    Set acc.parent = Me
    acc.proxyIndex = childID
    acc.proxyInfo = CreateObject("Scripting.Dictionary")
    acc.proxyInfo.Add "Name", oAcc.accName(childID - 1)
    acc.proxyInfo.Add "DefaultAction", oAcc.accDefaultAction(childID - 1)
    acc.proxyInfo.Add "Role", oAcc.AccRole(childID - 1)
    acc.proxyInfo.Add "State", oAcc.AccState(childID - 1)
    acc.proxyInfo.Add "Parent", oAcc.accParent(childID - 1)
    acc.proxyInfo.Add "Description", oAcc.accDescription(childID - 1)
    acc.proxyInfo.Add "KeyboardShortcut", oAcc.accKeyboardShortcut(childID - 1)
    acc.proxyInfo.Add "ChildCount", oAcc.accChildCount(childID - 1)
    acc.proxyInfo.Add "Focus", oAcc.accFocus(childID - 1)
    acc.proxyInfo.Add "Location", oAcc.accLocation(childID - 1)
    acc.proxyInfo.Add "HitTest", oAcc.accHitTest(childID - 1)
    acc.proxyInfo.Add "Selection", oAcc.accSelection(childID - 1)
    acc.proxyInfo.Add "Help", oAcc.accHelp(childID - 1)
    acc.proxyInfo.Add "HelpTopic", oAcc.accHelpTopic(childID - 1)
    
    Set createProxyChild = acc
End Function

Private Function getAccChildren(obj As IAccessible) As Collection
    Dim arr As Variant
    ReDim arr(1 To obj.accChildCount)
    
    Dim d
    Call MSAA.AccessibleChildren(obj, 0, obj.accChildCount, arr(1), d)
    
    Dim col As New Collection
    Dim i As Integer
    For i = 1 To obj.accChildCount
        col.Add arr(i)
    Next
    
    Set getAccChildren = col
End Function

Private Sub SetupDicts()
    'SETUP ACC_STATES DICTIONARY
    '*******************************
    Set ACC_STATES = CreateObject("Scripting.Dictionary")
    ACC_STATES.Add "NORMAL", &H0
    ACC_STATES.Add "UNAVAILABLE", &H1
    ACC_STATES.Add "SELECTED", &H2
    ACC_STATES.Add "FOCUSED", &H4
    ACC_STATES.Add "PRESSED", &H8
    ACC_STATES.Add "CHECKED", &H10
    ACC_STATES.Add "MIXED", &H20
    ACC_STATES.Add "INDETERMINATE", &H99
    ACC_STATES.Add "READONLY", &H40
    ACC_STATES.Add "HOTTRACKED", &H80
    ACC_STATES.Add "DEFAULT", &H100
    ACC_STATES.Add "EXPANDED", &H200
    ACC_STATES.Add "COLLAPSED", &H400
    ACC_STATES.Add "BUSY", &H800
    ACC_STATES.Add "FLOATING", &H1000
    ACC_STATES.Add "MARQUEED", &H2000
    ACC_STATES.Add "ANIMATED", &H4000
    ACC_STATES.Add "INVISIBLE", &H8000
    ACC_STATES.Add "OFFSCREEN", &H10000
    ACC_STATES.Add "SIZEABLE", &H20000
    ACC_STATES.Add "MOVEABLE", &H40000
    ACC_STATES.Add "SELFVOICING", &H80000
    ACC_STATES.Add "FOCUSABLE", &H100000
    ACC_STATES.Add "SELECTABLE", &H200000
    ACC_STATES.Add "LINKED", &H400000
    ACC_STATES.Add "TRAVERSED", &H800000
    ACC_STATES.Add "MULTISELECTABLE", &H1000000
    ACC_STATES.Add "EXTSELECTABLE", &H2000000
    ACC_STATES.Add "ALERT_LOW", &H4000000
    ACC_STATES.Add "ALERT_MEDIUM", &H8000000
    ACC_STATES.Add "ALERT_HIGH", &H10000000
    ACC_STATES.Add "PROTECTED", &H20000000
    ACC_STATES.Add "VALID", &H7FFFFFFF
    
    ACC_STATES.Add &H0, "NORMAL"
    ACC_STATES.Add &H1, "UNAVAILABLE"
    ACC_STATES.Add &H2, "SELECTED"
    ACC_STATES.Add &H4, "FOCUSED"
    ACC_STATES.Add &H8, "PRESSED"
    ACC_STATES.Add &H10, "CHECKED"
    ACC_STATES.Add &H20, "MIXED"
    ACC_STATES.Add &H99, "INDETERMINATE"
    ACC_STATES.Add &H40, "READONLY"
    ACC_STATES.Add &H80, "HOTTRACKED"
    ACC_STATES.Add &H100, "DEFAULT"
    ACC_STATES.Add &H200, "EXPANDED"
    ACC_STATES.Add &H400, "COLLAPSED"
    ACC_STATES.Add &H800, "BUSY"
    ACC_STATES.Add &H1000, "FLOATING"
    ACC_STATES.Add &H2000, "MARQUEED"
    ACC_STATES.Add &H4000, "ANIMATED"
    ACC_STATES.Add &H8000, "INVISIBLE"
    ACC_STATES.Add &H10000, "OFFSCREEN"
    ACC_STATES.Add &H20000, "SIZEABLE"
    ACC_STATES.Add &H40000, "MOVEABLE"
    ACC_STATES.Add &H80000, "SELFVOICING"
    ACC_STATES.Add &H100000, "FOCUSABLE"
    ACC_STATES.Add &H200000, "SELECTABLE"
    ACC_STATES.Add &H400000, "LINKED"
    ACC_STATES.Add &H800000, "TRAVERSED"
    ACC_STATES.Add &H1000000, "MULTISELECTABLE"
    ACC_STATES.Add &H2000000, "EXTSELECTABLE"
    ACC_STATES.Add &H4000000, "ALERT_LOW"
    ACC_STATES.Add &H8000000, "ALERT_MEDIUM"
    ACC_STATES.Add &H10000000, "ALERT_HIGH"
    ACC_STATES.Add &H20000000, "PROTECTED"
    ACC_STATES.Add &H7FFFFFFF, "VALID"
    
    'SETUP ACC_ROLES DICTIONARY
    '*******************************
    Set ACC_ROLES = CreateObject("Scripting.Dictionary")
    ACC_ROLES.Add "ROLE_TITLEBAR", &H1&
    ACC_ROLES.Add "ROLE_MENUBAR", &H2&
    ACC_ROLES.Add "ROLE_SCROLLBAR", &H3&
    ACC_ROLES.Add "ROLE_GRIP", &H4&
    ACC_ROLES.Add "ROLE_SOUND", &H5&
    ACC_ROLES.Add "ROLE_CURSOR", &H6&
    ACC_ROLES.Add "ROLE_CARET", &H7&
    ACC_ROLES.Add "ROLE_ALERT", &H8&
    ACC_ROLES.Add "ROLE_WINDOW", &H9&
    ACC_ROLES.Add "ROLE_CLIENT", &HA&
    ACC_ROLES.Add "ROLE_MENUPOPUP", &HB&
    ACC_ROLES.Add "ROLE_MENUITEM", &HC&
    ACC_ROLES.Add "ROLE_TOOLTIP", &HD&
    ACC_ROLES.Add "ROLE_APPLICATION", &HE&
    ACC_ROLES.Add "ROLE_DOCUMENT", &HF&
    ACC_ROLES.Add "ROLE_PANE", &H10&
    ACC_ROLES.Add "ROLE_CHART", &H11&
    ACC_ROLES.Add "ROLE_DIALOG", &H12&
    ACC_ROLES.Add "ROLE_BORDER", &H13&
    ACC_ROLES.Add "ROLE_GROUPING", &H14&
    ACC_ROLES.Add "ROLE_SEPARATOR", &H15&
    ACC_ROLES.Add "ROLE_TOOLBAR", &H16&
    ACC_ROLES.Add "ROLE_STATUSBAR", &H17&
    ACC_ROLES.Add "ROLE_TABLE", &H18&
    ACC_ROLES.Add "ROLE_COLUMNHEADER", &H19&
    ACC_ROLES.Add "ROLE_ROWHEADER", &H1A&
    ACC_ROLES.Add "ROLE_COLUMN", &H1B&
    ACC_ROLES.Add "ROLE_ROW", &H1C&
    ACC_ROLES.Add "ROLE_CELL", &H1D&
    ACC_ROLES.Add "ROLE_LINK", &H1E&
    ACC_ROLES.Add "ROLE_HELPBALLOON", &H1F&
    ACC_ROLES.Add "ROLE_CHARACTER", &H20&
    ACC_ROLES.Add "ROLE_LIST", &H21&
    ACC_ROLES.Add "ROLE_LISTITEM", &H22&
    ACC_ROLES.Add "ROLE_OUTLINE", &H23&
    ACC_ROLES.Add "ROLE_OUTLINEITEM", &H24&
    ACC_ROLES.Add "ROLE_PAGETAB", &H25&
    ACC_ROLES.Add "ROLE_PROPERTYPAGE", &H26&
    ACC_ROLES.Add "ROLE_INDICATOR", &H27&
    ACC_ROLES.Add "ROLE_GRAPHIC", &H28&
    ACC_ROLES.Add "ROLE_STATICTEXT", &H29&
    ACC_ROLES.Add "ROLE_TEXT", &H2A&
    ACC_ROLES.Add "ROLE_PUSHBUTTON", &H2B&
    ACC_ROLES.Add "ROLE_CHECKBUTTON", &H2C&
    ACC_ROLES.Add "ROLE_RADIOBUTTON", &H2D&
    ACC_ROLES.Add "ROLE_COMBOBOX", &H2E&
    ACC_ROLES.Add "ROLE_DROPLIST", &H2F&
    ACC_ROLES.Add "ROLE_PROGRESSBAR", &H30&
    ACC_ROLES.Add "ROLE_DIAL", &H31&
    ACC_ROLES.Add "ROLE_HOTKEYFIELD", &H32&
    ACC_ROLES.Add "ROLE_SLIDER", &H33&
    ACC_ROLES.Add "ROLE_SPINBUTTON", &H34&
    ACC_ROLES.Add "ROLE_DIAGRAM", &H35&
    ACC_ROLES.Add "ROLE_ANIMATION", &H36&
    ACC_ROLES.Add "ROLE_EQUATION", &H37&
    ACC_ROLES.Add "ROLE_BUTTONDROPDOWN", &H38&
    ACC_ROLES.Add "ROLE_BUTTONMENU", &H39&
    ACC_ROLES.Add "ROLE_BUTTONDROPDOWNGRID", &H3A&
    ACC_ROLES.Add "ROLE_WHITESPACE", &H3B&
    ACC_ROLES.Add "ROLE_PAGETABLIST", &H3C&
    ACC_ROLES.Add "ROLE_CLOCK", &H3D&
    
    ACC_ROLES.Add &H1&, "ROLE_TITLEBAR"
    ACC_ROLES.Add &H2&, "ROLE_MENUBAR"
    ACC_ROLES.Add &H3&, "ROLE_SCROLLBAR"
    ACC_ROLES.Add &H4&, "ROLE_GRIP"
    ACC_ROLES.Add &H5&, "ROLE_SOUND"
    ACC_ROLES.Add &H6&, "ROLE_CURSOR"
    ACC_ROLES.Add &H7&, "ROLE_CARET"
    ACC_ROLES.Add &H8&, "ROLE_ALERT"
    ACC_ROLES.Add &H9&, "ROLE_WINDOW"
    ACC_ROLES.Add &HA&, "ROLE_CLIENT"
    ACC_ROLES.Add &HB&, "ROLE_MENUPOPUP"
    ACC_ROLES.Add &HC&, "ROLE_MENUITEM"
    ACC_ROLES.Add &HD&, "ROLE_TOOLTIP"
    ACC_ROLES.Add &HE&, "ROLE_APPLICATION"
    ACC_ROLES.Add &HF&, "ROLE_DOCUMENT"
    ACC_ROLES.Add &H10&, "ROLE_PANE"
    ACC_ROLES.Add &H11&, "ROLE_CHART"
    ACC_ROLES.Add &H12&, "ROLE_DIALOG"
    ACC_ROLES.Add &H13&, "ROLE_BORDER"
    ACC_ROLES.Add &H14&, "ROLE_GROUPING"
    ACC_ROLES.Add &H15&, "ROLE_SEPARATOR"
    ACC_ROLES.Add &H16&, "ROLE_TOOLBAR"
    ACC_ROLES.Add &H17&, "ROLE_STATUSBAR"
    ACC_ROLES.Add &H18&, "ROLE_TABLE"
    ACC_ROLES.Add &H19&, "ROLE_COLUMNHEADER"
    ACC_ROLES.Add &H1A&, "ROLE_ROWHEADER"
    ACC_ROLES.Add &H1B&, "ROLE_COLUMN"
    ACC_ROLES.Add &H1C&, "ROLE_ROW"
    ACC_ROLES.Add &H1D&, "ROLE_CELL"
    ACC_ROLES.Add &H1E&, "ROLE_LINK"
    ACC_ROLES.Add &H1F&, "ROLE_HELPBALLOON"
    ACC_ROLES.Add &H20&, "ROLE_CHARACTER"
    ACC_ROLES.Add &H21&, "ROLE_LIST"
    ACC_ROLES.Add &H22&, "ROLE_LISTITEM"
    ACC_ROLES.Add &H23&, "ROLE_OUTLINE"
    ACC_ROLES.Add &H24&, "ROLE_OUTLINEITEM"
    ACC_ROLES.Add &H25&, "ROLE_PAGETAB"
    ACC_ROLES.Add &H26&, "ROLE_PROPERTYPAGE"
    ACC_ROLES.Add &H27&, "ROLE_INDICATOR"
    ACC_ROLES.Add &H28&, "ROLE_GRAPHIC"
    ACC_ROLES.Add &H29&, "ROLE_STATICTEXT"
    ACC_ROLES.Add &H2A&, "ROLE_TEXT"
    ACC_ROLES.Add &H2B&, "ROLE_PUSHBUTTON"
    ACC_ROLES.Add &H2C&, "ROLE_CHECKBUTTON"
    ACC_ROLES.Add &H2D&, "ROLE_RADIOBUTTON"
    ACC_ROLES.Add &H2E&, "ROLE_COMBOBOX"
    ACC_ROLES.Add &H2F&, "ROLE_DROPLIST"
    ACC_ROLES.Add &H30&, "ROLE_PROGRESSBAR"
    ACC_ROLES.Add &H31&, "ROLE_DIAL"
    ACC_ROLES.Add &H32&, "ROLE_HOTKEYFIELD"
    ACC_ROLES.Add &H33&, "ROLE_SLIDER"
    ACC_ROLES.Add &H34&, "ROLE_SPINBUTTON"
    ACC_ROLES.Add &H35&, "ROLE_DIAGRAM"
    ACC_ROLES.Add &H36&, "ROLE_ANIMATION"
    ACC_ROLES.Add &H37&, "ROLE_EQUATION"
    ACC_ROLES.Add &H38&, "ROLE_BUTTONDROPDOWN"
    ACC_ROLES.Add &H39&, "ROLE_BUTTONMENU"
    ACC_ROLES.Add &H3A&, "ROLE_BUTTONDROPDOWNGRID"
    ACC_ROLES.Add &H3B&, "ROLE_WHITESPACE"
    ACC_ROLES.Add &H3C&, "ROLE_PAGETABLIST"
    ACC_ROLES.Add &H3D&, "ROLE_CLOCK"
End Sub

Public Function getPath(Optional toAccessible As STD_Automation_Accesibility = Nothing) As String
    'Initialise trace
    Dim acc As STD_Automation_Accesibility
    Set acc = Me
    
    'Collection to store path
    Dim col As New Collection
    
    
    'Collect parents
    While Not acc.parent.Name = "Desktop"
        Dim child As STD_Automation_Accesibility
        Dim index As Long
        index = 0
        For Each child In acc.parent.children
            index = index + 1
            If child.hwnd = acc.hwnd And child.Role = acc.Role Then
                Exit For
            End If
        Next
        
        'Add index to stack
        col.Add index
                
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
    Dim acc As STD_Automation_Accesibility
    For Each acc In Me.children
        sChildren = sChildren & "," & acc.toJSON()
    Next
    sChildren = Mid(sChildren, 2)
    
    'Convert to json
    toJSON = "{" & _
        "name:""" & jsonEscape(Me.Name) & """," & _
        "desc:""" & jsonEscape(Me.Description) & """," & _
        "value:""" & jsonEscape(Me.Value) & """," & _
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

Friend Sub SetOAcc(ByRef acc As Object)
    Set oAcc = acc
End Sub

Public Sub AssertFalse()
    Debug.Assert False
End Sub

Private Sub Class_Initialize()
    Call SetupDicts
End Sub


