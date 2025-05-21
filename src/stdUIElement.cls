VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdUIElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Enum EUIElementEvent
  uiElementEventChange
  uiElementEventMove
  uiElementEventResize
  uiElementEventMouseClick
  uiElementEventMouseDblClick
  uiElementEventMouseFocus
  uiElementEventMouseMove
  uiElementEventMouseDown
  uiElementEventMouseUp
  uiElementEventTabSelected
  uiElementEventAddChild
  uiElementEventKeyDown
  uiElementEventKeyUp
  uiElementEventKeyPress
End Enum
Public Enum EUIElementType
  'Group controls
  uiUserForm
  uiFrame
  uiMultiPage

  'Display controls
  uiLabel
  uiImage
  uiTabStrip

  'Edit Controls
  uiTextBox
  uiCombobox
  uiListBox
  uiCheckBox
  uiOptionButton
  uiToggleButton
  uiScrollbar
  uiSpinButton

  'Executors
  uiCommandButton
End Enum
Private Type TThis
  messageRouter As stdICallable
  Control As MSForms.Control
  iType As EUIElementType
  id As String
  Left As Double
  Width As Double
End Type

Private this As TThis
Private WithEvents pUserForm       As MSForms.UserForm
Private WithEvents pFrame          As MSForms.Frame
Private WithEvents pMultiPage      As MSForms.MultiPage
Private WithEvents pLabel          As MSForms.Label
Private WithEvents pImage          As MSForms.Image
Private WithEvents pTabStrip       As MSForms.TabStrip
Private WithEvents pTextBox        As MSForms.TextBox
Private WithEvents pComboBox       As MSForms.ComboBox
Private WithEvents pListBox        As MSForms.ListBox
Private WithEvents pCheckBox       As MSForms.CheckBox
Private WithEvents pOptionButton   As MSForms.OptionButton
Private WithEvents pToggleButton   As MSForms.ToggleButton
Private WithEvents pScrollBar      As MSForms.ScrollBar
Private WithEvents pSpinButton     As MSForms.SpinButton
Private WithEvents pCommandButton  As MSForms.CommandButton


'CaptionSupport: pUserForm, pFrame, pLabel,pCheckBox, pOptionButton, pToggleButton, pCommandButton
'ValueSupport: pMultipage, pOptionButton, pTabStrip, pTextBox, pCombobox, pListBox, pCheckBox, pOptionButton, pToggleButton, pScrollbar, pSpinButton
'ValueEsque: Image=>Picture


'Creating a UI Element from an existing control
'@param ctrl - The control to create the UI Element from
'@param callback as stdICallable<(el: stdUIElement, event: EUIElementEvent, params: Variant<Array<Variant>>)=>void> - The callback object to use for event handling
'@returns - The UI Element
Public Function CreateFromControl(ByVal ctrl As Control, callback As stdICallable) As stdUIElement
  Set CreateFromControl = New stdUIElement
  if not callback is nothing then Call BindGlobals(callback)
  Call CreateFromControl.protInit(ctrl, callback)
End Function

'Creating a UI Element from a type
'@param oControls - The controls collection to add the control to
'@param iCtrlType - The type of control to create
'@param sName - The name of the control
'@param Caption - The caption of the control
'@param callback as stdICallable<(el: stdUIElement, event: EUIElementEvent, params: Variant<Array<Variant>>)=>void> - The callback object to use for event handling
'@param fLeft - The left position of the control
'@param fTop - The top position of the control
'@param fWidth - The width of the control
'@param fHeight - The height of the control
'@returns - The UI Element
Public Function CreateFromType(ByVal oControls As MSForms.controls, ByVal iCtrlType As EUIElementType, Optional ByVal sName As String, Optional ByVal Caption As String, Optional ByVal Value As Variant, Optional ByVal callback As stdICallable, Optional ByVal fLeft As Double = 0, Optional ByVal fTop As Double = 0, Optional ByVal fWidth As Double, Optional ByVal fHeight As Double) As stdUIElement
  Set CreateFromType = New stdUIElement
  Dim ctrl As Control
  Select Case iCtrlType
    Case uiUserForm:      Set ctrl = oControls.Add("Forms.Form.1", sName)
    Case uiFrame:         Set ctrl = oControls.Add("Forms.Frame.1", sName)
    Case uiMultiPage:     Set ctrl = oControls.Add("Forms.MultiPage.1", sName)
    Case uiLabel:         Set ctrl = oControls.Add("Forms.Label.1", sName)
    Case uiImage:         Set ctrl = oControls.Add("Forms.Image.1", sName)
    Case uiTabStrip:      Set ctrl = oControls.Add("Forms.TabStrip.1", sName)
    Case uiTextBox:       Set ctrl = oControls.Add("Forms.TextBox.1", sName)
    Case uiCombobox:      Set ctrl = oControls.Add("Forms.ComboBox.1", sName)
    Case uiListBox:       Set ctrl = oControls.Add("Forms.ListBox.1", sName)
    Case uiCheckBox:      Set ctrl = oControls.Add("Forms.CheckBox.1", sName)
    Case uiOptionButton:  Set ctrl = oControls.Add("Forms.OptionButton.1", sName)
    Case uiToggleButton:  Set ctrl = oControls.Add("Forms.ToggleButton.1", sName)
    Case uiScrollbar:     Set ctrl = oControls.Add("Forms.ScrollBar.1", sName)
    Case uiSpinButton:    Set ctrl = oControls.Add("Forms.SpinButton.1", sName)
    Case uiCommandButton: Set ctrl = oControls.Add("Forms.CommandButton.1", sName)
  End Select
  If sName <> "" Then ctrl.name = sName
  if not callback is nothing then Call BindGlobals(callback)
  Call CreateFromType.protInit(ctrl, callback)
  If not isMissing(Value) then
    If IsObject(Value) Then
      Set CreateFromType.Value = Value
    Else
      CreateFromType.Value = Value
    End If
  end if
  CreateFromType.Caption = Caption
  CreateFromType.Left = fLeft
  CreateFromType.Top = fTop
  If fWidth <> 0 Then CreateFromType.Width = fWidth
  If fHeight <> 0 Then CreateFromType.Height = fHeight
End Function

'Obtain type of input control
'@returns - The type of input control
Public Property Get uiType() As EUIElementType
  uiType = this.iType
End Property

'Obtain control object
'@returns - The control object
Public Property Get Control() As Control
  Set Control = this.Control
End Property

'Obtain ui object raw item
'@returns - The ui object raw item / event sink
Public Property Get uiObject() As Object
  Select Case this.iType
    Case uiUserForm:       Set uiObject = pUserForm
    Case uiFrame:          Set uiObject = pFrame
    Case uiMultiPage:      Set uiObject = pMultiPage
    Case uiLabel:          Set uiObject = pLabel
    Case uiImage:          Set uiObject = pImage
    Case uiTabStrip:       Set uiObject = pTabStrip
    Case uiTextBox:        Set uiObject = pTextBox
    Case uiCombobox:       Set uiObject = pComboBox
    Case uiListBox:        Set uiObject = pListBox
    Case uiCheckBox:       Set uiObject = pCheckBox
    Case uiOptionButton:   Set uiObject = pOptionButton
    Case uiToggleButton:   Set uiObject = pToggleButton
    Case uiScrollbar:      Set uiObject = pScrollBar
    Case uiSpinButton:     Set uiObject = pSpinButton
    Case uiCommandButton:  Set uiObject = pCommandButton
  End Select
End Property

'Get/Set Value
'@returns - The value of the control
Public Property Get Value() As Variant
  Select Case this.iType
    Case uiTextBox, uiCombobox, uiListBox, uiCheckBox, uiOptionButton, uiToggleButton, uiScrollbar, uiSpinButton
      Value = uiObject.Value
    Case uiImage
      Set Value = uiObject.Picture
  End Select
End Property
Public Property Let Value(v As Variant)
  Select Case this.iType
    Case uiTextBox, uiCombobox, uiListBox, uiCheckBox, uiOptionButton, uiToggleButton, uiScrollbar, uiSpinButton
      uiObject.Value = v
  End Select
  Call callEventCB(ObjRef(Me), uiElementEventChange, Array(v))
End Property
Public Property Set Value(v As Variant)
  Select Case this.iType
    Case uiImage
      Set uiObject.Picture = v
  End Select
  Call callEventCB(ObjRef(Me), uiElementEventChange, Array(v))
End Property


'Get/Set Caption - The visible text of a control
'@returns - The caption of the control
Public Property Get Caption() As String
  Select Case this.iType
    Case uiUserForm, uiFrame, uiLabel, uiCheckBox, uiOptionButton, uiToggleButton, uiCommandButton
      Caption = This.Control.Caption
  End Select
End Property
Public Property Let Caption(v As String)
  Select Case this.iType
    Case uiUserForm, uiFrame, uiLabel, uiCheckBox, uiOptionButton, uiToggleButton, uiCommandButton
      This.Control.Caption = v
  End Select
  Call callEventCB(ObjRef(Me), uiElementEventChange, Array(v))
End Property

'Get/Set Position Left
'@returns - The left position of the control
Public Property Get Left() As Double
  Left = this.Left
End Property
Public Property Let Left(fLeft As Double)
  this.Left = fLeft
  Select Case this.iType
    Case uiTextBox, uiCombobox, uiListBox, uiOptionButton, uiToggleButton, uiScrollbar, uiSpinButton, uiLabel
      uiObject.Left = this.Left
    Case uiCheckBox
      uiObject.Left = this.Left + iif(this.Width <= 12, 0, this.Width / 2 - 6)
  End Select
  Call callEventCB(ObjRef(Me), uiElementEventMove, Array(this.Left, this.Control.Top))
End Property

'Get/Set Position Top
'@returns - The top position of the control
Public Property Get Top() As Double
  Top = this.Control.Top
End Property
Public Property Let Top(fTop As Double)
  this.Control.Top = fTop
  Call callEventCB(ObjRef(Me), uiElementEventMove, Array(this.Control.Left, this.Control.Top))
End Property

'Get/Set Size Width
'@returns - The width of the control
Public Property Get Width() As Double
  Width = this.Width
End Property
Public Property Let Width(fWidth As Double)
  this.Width = fWidth
  Left = this.Left
  uiObject.Width = fWidth
  Call callEventCB(ObjRef(Me), uiElementEventResize, Array(this.Control.Width, this.Control.Height))
End Property

'Get/Set Size Height
'@returns - The height of the control
Public Property Get Height() As Double
  Height = this.Control.Height
End Property
Public Property Let Height(fHeight As Double)
  this.Control.Height = fHeight
  Call callEventCB(ObjRef(Me), uiElementEventResize, Array(this.Control.Width, this.Control.Height))
End Property

'Getter for id
'@returns - The id of the control. This is a generated GUID
Public Property Get id() As String
  id = this.id
End Property


'*********************
'* PROTECTED METHODS *
'*********************

'Initialise input control
'@protected
Friend Sub protInit(ctrl As Control, callback As stdICallable)
  Set this.messageRouter = callback
  Set this.Control = ctrl
  this.id = getGUID()
  If TypeOf ctrl Is MSForms.UserForm Then
    Set pUserForm = ctrl
    this.iType = EUIElementType.uiUserForm
    Exit Sub
  End If
  If TypeOf ctrl Is MSForms.Frame Then
    Set pFrame = ctrl
    this.iType = EUIElementType.uiFrame
    Exit Sub
  End If
  If TypeOf ctrl Is MSForms.MultiPage Then
    Set pMultiPage = ctrl
    this.iType = EUIElementType.uiMultiPage
    Exit Sub
  End If
  If TypeOf ctrl Is MSForms.Label Then
    Set pLabel = ctrl
    this.iType = EUIElementType.uiLabel
    Exit Sub
  End If
  If TypeOf ctrl Is MSForms.Image Then
    Set pImage = ctrl
    this.iType = EUIElementType.uiImage
    Exit Sub
  End If
  If TypeOf ctrl Is MSForms.TabStrip Then
    Set pTabStrip = ctrl
    this.iType = EUIElementType.uiTabStrip
    Exit Sub
  End If
  If TypeOf ctrl Is MSForms.TextBox Then
    Set pTextBox = ctrl
    this.iType = EUIElementType.uiTextBox
    Exit Sub
  End If
  If TypeOf ctrl Is MSForms.ComboBox Then
    Set pComboBox = ctrl
    this.iType = EUIElementType.uiCombobox
    Exit Sub
  End If
  If TypeOf ctrl Is MSForms.ListBox Then
    Set pListBox = ctrl
    this.iType = EUIElementType.uiListBox
    Exit Sub
  End If
  If TypeOf ctrl Is MSForms.CheckBox Then
    Set pCheckBox = ctrl
    this.iType = EUIElementType.uiCheckBox
    Exit Sub
  End If
  If TypeOf ctrl Is MSForms.OptionButton Then
    Set pOptionButton = ctrl
    this.iType = EUIElementType.uiOptionButton
    Exit Sub
  End If
  If TypeOf ctrl Is MSForms.ToggleButton Then
    Set pToggleButton = ctrl
    this.iType = EUIElementType.uiToggleButton
    Exit Sub
  End If
  If TypeOf ctrl Is MSForms.ScrollBar Then
    Set pScrollBar = ctrl
    this.iType = EUIElementType.uiScrollbar
    Exit Sub
  End If
  If TypeOf ctrl Is MSForms.SpinButton Then
    Set pSpinButton = ctrl
    this.iType = EUIElementType.uiSpinButton
    Exit Sub
  End If
  If TypeOf ctrl Is MSForms.CommandButton Then
    Set pCommandButton = ctrl
    this.iType = EUIElementType.uiCommandButton
    Exit Sub
  End If
End Sub

'Get a GUID
'@returns - A new GUID
Private Function getGUID() As String
  getGUID = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx"
  getGUID = Replace(getGUID, "y", Hex(Rnd() And &H3 Or &H8))
  Dim i As Long: For i = 1 To 30
    getGUID = Replace(getGUID, "x", Hex$(CLng(Rnd() * 15.9999)), 1, 1)
  Next
End Function



'Bind globals to the callback object
Private Sub BindGlobals(ByVal query as stdICallable)
  Dim bSuccess as boolean
  Call query.SendMessage("bindGlobal", bSuccess, Array("EUIElementEvent", stdUIElement.protUIElementEventLookup()))
End Sub

'Get a dictionary of event names and their corresponding event handler
'@protected
'@returns - A dictionary of event names and their corresponding event handler
Friend Function protUIElementEventLookup() as Object
  static obj as Object
  if obj is nothing then
    set obj = CreateObject("Scripting.Dictionary")
    obj.CompareMode = vbTextCompare
    obj.add "uiElementEventChange"       , uiElementEventChange
    obj.add "uiElementEventMove"         , uiElementEventMove
    obj.add "uiElementEventResize"       , uiElementEventResize
    obj.add "uiElementEventMouseClick"   , uiElementEventMouseClick
    obj.add "uiElementEventMouseDblClick", uiElementEventMouseDblClick
    obj.add "uiElementEventMouseFocus"   , uiElementEventMouseFocus
    obj.add "uiElementEventMouseMove"    , uiElementEventMouseMove
    obj.add "uiElementEventMouseDown"    , uiElementEventMouseDown
    obj.add "uiElementEventMouseUp"      , uiElementEventMouseUp
    obj.add "uiElementEventTabSelected"  , uiElementEventTabSelected
    obj.add "uiElementEventAddChild"     , uiElementEventAddChild
    obj.add "uiElementEventKeyDown"      , uiElementEventKeyDown
    obj.add "uiElementEventKeyUp"        , uiElementEventKeyUp
    obj.add "uiElementEventKeyPress"     , uiElementEventKeyPress
  end if
  set protUIElementEventLookup = obj
End Function

'Return a reference to an object
'@param obj - The object to return a reference to
'@returns - A reference to the object
Private Function ObjRef(ByVal obj As Object) As Object
  Set ObjRef = obj
End Function









'TODO: Consider calculating and passing X and Y coords of mouse press event
'@remark The following controls are missing the click event: pScrollBar, pSpinButton
Private Sub pUserForm_Click():      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseClick): End Sub
Private Sub pFrame_Click():         Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseClick): End Sub
Private Sub pLabel_Click():         Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseClick): End Sub
Private Sub pComboBox_Click():      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseClick): End Sub
Private Sub pListBox_Click():       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseClick): End Sub
Private Sub pCheckBox_Click():      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseClick): End Sub
Private Sub pOptionButton_Click():  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseClick): End Sub
Private Sub pToggleButton_Click():  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseClick): End Sub
Private Sub pCommandButton_Click(): Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseClick): End Sub
Private Sub pMultiPage_Click(ByVal Index As Long)
  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseClick)
  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventTabSelected, Array(Index))
End Sub
Private Sub pTabStrip_Click(ByVal Index As Long)
  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseClick)
  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventTabSelected, Array(Index))
End Sub

'Parent object events
Private Sub pUserForm_AddControl(ByVal Control As MSForms.Control):                       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventAddChild, Array(Control, 0)): End Sub
Private Sub pFrame_AddControl(ByVal Control As MSForms.Control):                          Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventAddChild, Array(Control, 0)): End Sub
Private Sub pMultiPage_AddControl(ByVal Index As Long, ByVal Control As MSForms.Control): Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventAddChild, Array(Control, Index)): End Sub


'Change events
Private Sub pMultiPage_Change():    Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventChange, Array(uiObject.Value)): End Sub
Private Sub pTabStrip_Change():     Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventChange, Array(uiObject.Value)): End Sub
Private Sub pCheckBox_Change():     Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventChange, Array(uiObject.Value)): End Sub
Private Sub pComboBox_Change():     Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventChange, Array(uiObject.Value)): End Sub
Private Sub pListBox_Change():      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventChange, Array(uiObject.Value)): End Sub
Private Sub pOptionButton_Change(): Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventChange, Array(uiObject.Value)): End Sub
Private Sub pScrollBar_Change():    Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventChange, Array(uiObject.Value)): End Sub
Private Sub pSpinButton_Change():   Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventChange, Array(uiObject.Value)): End Sub
Private Sub pTextBox_Change():      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventChange, Array(uiObject.Value)): End Sub
Private Sub pToggleButton_Change(): Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventChange, Array(uiObject.Value)): End Sub

'*******************
'* KEYBOARD EVENTS *
'*******************

Private Sub pToggleButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyDown, Array(KeyCode, Shift)): End Sub
Private Sub pUserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyDown, Array(KeyCode, Shift)): End Sub
Private Sub pCheckBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyDown, Array(KeyCode, Shift)): End Sub
Private Sub pComboBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyDown, Array(KeyCode, Shift)): End Sub
Private Sub pCommandButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer): Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyDown, Array(KeyCode, Shift)): End Sub
Private Sub pFrame_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):         Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyDown, Array(KeyCode, Shift)): End Sub
Private Sub pListBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyDown, Array(KeyCode, Shift)): End Sub
Private Sub pMultiPage_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):     Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyDown, Array(KeyCode, Shift)): End Sub
Private Sub pOptionButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyDown, Array(KeyCode, Shift)): End Sub
Private Sub pScrollBar_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):     Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyDown, Array(KeyCode, Shift)): End Sub
Private Sub pSpinButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):    Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyDown, Array(KeyCode, Shift)): End Sub
Private Sub pTabStrip_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyDown, Array(KeyCode, Shift)): End Sub
Private Sub pTextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyDown, Array(KeyCode, Shift)): End Sub

Private Sub pToggleButton_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger):  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyPress, Array(KeyAscii)): End Sub
Private Sub pUserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger):      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyPress, Array(KeyAscii)): End Sub
Private Sub pCheckBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger):      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyPress, Array(KeyAscii)): End Sub
Private Sub pComboBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger):      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyPress, Array(KeyAscii)): End Sub
Private Sub pCommandButton_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger): Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyPress, Array(KeyAscii)): End Sub
Private Sub pFrame_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger):         Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyPress, Array(KeyAscii)): End Sub
Private Sub pListBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger):       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyPress, Array(KeyAscii)): End Sub
Private Sub pMultiPage_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger):     Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyPress, Array(KeyAscii)): End Sub
Private Sub pOptionButton_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger):  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyPress, Array(KeyAscii)): End Sub
Private Sub pScrollBar_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger):     Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyPress, Array(KeyAscii)): End Sub
Private Sub pSpinButton_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger):    Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyPress, Array(KeyAscii)): End Sub
Private Sub pTabStrip_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger):      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyPress, Array(KeyAscii)): End Sub
Private Sub pTextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger):       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyPress, Array(KeyAscii)): End Sub

Private Sub pToggleButton_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyUp, Array(KeyCode, Shift)): End Sub
Private Sub pUserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyUp, Array(KeyCode, Shift)): End Sub
Private Sub pCheckBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyUp, Array(KeyCode, Shift)): End Sub
Private Sub pComboBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyUp, Array(KeyCode, Shift)): End Sub
Private Sub pCommandButton_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer): Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyUp, Array(KeyCode, Shift)): End Sub
Private Sub pFrame_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):         Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyUp, Array(KeyCode, Shift)): End Sub
Private Sub pListBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyUp, Array(KeyCode, Shift)): End Sub
Private Sub pMultiPage_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):     Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyUp, Array(KeyCode, Shift)): End Sub
Private Sub pOptionButton_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyUp, Array(KeyCode, Shift)): End Sub
Private Sub pScrollBar_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):     Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyUp, Array(KeyCode, Shift)): End Sub
Private Sub pSpinButton_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):    Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyUp, Array(KeyCode, Shift)): End Sub
Private Sub pTabStrip_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyUp, Array(KeyCode, Shift)): End Sub
Private Sub pTextBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyUp, Array(KeyCode, Shift)): End Sub

'****************
'* MOUSE EVENTS *
'****************

Private Sub pUserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):                       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDown, Array(Button, Shift, X, Y)): End Sub
Private Sub pCheckBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):                       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDown, Array(Button, Shift, X, Y)): End Sub
Private Sub pComboBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):                       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDown, Array(Button, Shift, X, Y)): End Sub
Private Sub pCommandButton_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):                  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDown, Array(Button, Shift, X, Y)): End Sub
Private Sub pFrame_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):                          Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDown, Array(Button, Shift, X, Y)): End Sub
Private Sub pLabel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):                          Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDown, Array(Button, Shift, X, Y)): End Sub
Private Sub pListBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):                        Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDown, Array(Button, Shift, X, Y)): End Sub
Private Sub pMultiPage_MouseDown(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDown, Array(Button, Shift, X, Y)): End Sub
Private Sub pOptionButton_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):                   Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDown, Array(Button, Shift, X, Y)): End Sub
Private Sub pTabStrip_MouseDown(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDown, Array(Button, Shift, X, Y)): End Sub
Private Sub pToggleButton_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):                   Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDown, Array(Button, Shift, X, Y)): End Sub
Private Sub pImage_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseClick) 'There is no MouseClick event for Image, so we raise this here too
  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDown, Array(Button, Shift, X, Y))
End Sub
Private Sub pTextBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseClick) 'There is no MouseClick event for TextBox, so we raise this here too
  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDown, Array(Button, Shift, X, Y))
End Sub

Private Sub pUserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):                       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseMove, Array(Button, Shift, X, Y)): End Sub
Private Sub pCheckBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):                       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseMove, Array(Button, Shift, X, Y)): End Sub
Private Sub pComboBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):                       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseMove, Array(Button, Shift, X, Y)): End Sub
Private Sub pCommandButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):                  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseMove, Array(Button, Shift, X, Y)): End Sub
Private Sub pFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):                          Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseMove, Array(Button, Shift, X, Y)): End Sub
Private Sub pImage_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):                          Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseMove, Array(Button, Shift, X, Y)): End Sub
Private Sub pLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):                          Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseMove, Array(Button, Shift, X, Y)): End Sub
Private Sub pListBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):                        Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseMove, Array(Button, Shift, X, Y)): End Sub
Private Sub pMultiPage_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseMove, Array(Button, Shift, X, Y)): End Sub
Private Sub pOptionButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):                   Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseMove, Array(Button, Shift, X, Y)): End Sub
Private Sub pTabStrip_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseMove, Array(Button, Shift, X, Y)): End Sub
Private Sub pTextBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):                        Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseMove, Array(Button, Shift, X, Y)): End Sub
Private Sub pToggleButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):                   Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseMove, Array(Button, Shift, X, Y)): End Sub

Private Sub pUserForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):                       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseUp, Array(Button, Shift, X, Y)): End Sub
Private Sub pCheckBox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):                       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseUp, Array(Button, Shift, X, Y)): End Sub
Private Sub pComboBox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):                       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseUp, Array(Button, Shift, X, Y)): End Sub
Private Sub pCommandButton_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):                  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseUp, Array(Button, Shift, X, Y)): End Sub
Private Sub pFrame_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):                          Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseUp, Array(Button, Shift, X, Y)): End Sub
Private Sub pImage_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):                          Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseUp, Array(Button, Shift, X, Y)): End Sub
Private Sub pLabel_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):                          Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseUp, Array(Button, Shift, X, Y)): End Sub
Private Sub pListBox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):                        Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseUp, Array(Button, Shift, X, Y)): End Sub
Private Sub pMultiPage_MouseUp(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseUp, Array(Button, Shift, X, Y)): End Sub
Private Sub pOptionButton_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):                   Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseUp, Array(Button, Shift, X, Y)): End Sub
Private Sub pTabStrip_MouseUp(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseUp, Array(Button, Shift, X, Y)): End Sub
Private Sub pTextBox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):                        Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseUp, Array(Button, Shift, X, Y)): End Sub
Private Sub pToggleButton_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single):                   Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseUp, Array(Button, Shift, X, Y)): End Sub

Private Sub pCheckBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean):                        Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDblClick, Array(Cancel, 0)): End Sub
Private Sub pComboBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean):                        Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDblClick, Array(Cancel, 0)): End Sub
Private Sub pCommandButton_DblClick(ByVal Cancel As MSForms.ReturnBoolean):                   Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDblClick, Array(Cancel, 0)): End Sub
Private Sub pFrame_DblClick(ByVal Cancel As MSForms.ReturnBoolean):                           Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDblClick, Array(Cancel, 0)): End Sub
Private Sub pLabel_DblClick(ByVal Cancel As MSForms.ReturnBoolean):                           Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDblClick, Array(Cancel, 0)): End Sub
Private Sub pListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean):                         Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDblClick, Array(Cancel, 0)): End Sub
Private Sub pOptionButton_DblClick(ByVal Cancel As MSForms.ReturnBoolean):                    Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDblClick, Array(Cancel, 0)): End Sub
Private Sub pTextBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean):                         Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDblClick, Array(Cancel, 0)): End Sub
Private Sub pToggleButton_DblClick(ByVal Cancel As MSForms.ReturnBoolean):                    Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDblClick, Array(Cancel, 0)): End Sub
Private Sub pUserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean):                        Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDblClick, Array(Cancel, 0)): End Sub
Private Sub pMultiPage_DblClick(ByVal Index As Long, ByVal Cancel As MSForms.ReturnBoolean):  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDblClick, Array(Cancel, Index)): End Sub
Private Sub pTabStrip_DblClick(ByVal Index As Long, ByVal Cancel As MSForms.ReturnBoolean):   Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDblClick, Array(Cancel, Index)): End Sub
'TODO: SpinButton_DblClick using MouseDown
'TODO: ScrollBar_DblClick using MouseDown
'TODO: Image_DblClick using MouseDown

Private Sub callEventCB(ByVal MeRef As stdUIElement, ByVal iMsg As EUIElementEvent, Optional ByVal args As Variant = Empty)
  If IsEmpty(args) Then args = Array()
  If Not this.messageRouter Is Nothing Then
    Call this.messageRouter.Run(MeRef, iMsg, args)
  End If
End Sub