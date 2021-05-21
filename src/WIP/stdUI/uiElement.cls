VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "uiElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Enum EUIElementMessage
  uiElementEventChange
  uiElementEventMove
  uiElementEventResize
  uiElementEventMouseClick
  uiElementEventMouseFocus
  uiElementEventMouseMove
  uiElementEventTabSelected
  uiElementEventAddChild
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
  uiComboBox
  uiListBox
  uiCheckBox
  uiOptionButton
  uiToggleButton
  uiScrollBar
  uiSpinButton

  'Executors
  uiCommandButton
End Enum

Private pParent  as uiIMessagable
Private pControl as MSForms.Control

Private WithEvents pUserForm       as MSForms.UserForm
Private WithEvents pFrame          as MSForms.Frame
Private WithEvents pMultiPage      as MSForms.MultiPage
Private WithEvents pLabel          as MSForms.Label
Private WithEvents pImage          as MSForms.Image
Private WithEvents pTabStrip       as MSForms.TabStrip
Private WithEvents pTextBox        as MSForms.TextBox
Private WithEvents pComboBox       as MSForms.ComboBox
Private WithEvents pListBox        as MSForms.ListBox
Private WithEvents pCheckBox       as MSForms.CheckBox
Private WithEvents pOptionButton   as MSForms.OptionButton
Private WithEvents pToggleButton   as MSForms.ToggleButton
Private WithEvents pScrollBar      as MSForms.ScrollBar
Private WithEvents pSpinButton     as MSForms.SpinButton
Private WithEvents pCommandButton  as MSForms.CommandButton

Private pType as EUIElementType
Private pID as String

'Creating a input component from 
Public Function CreateFromControl(ByVal parent as uiIMessagable, ByVal ctrl as Control) as uiElement
  set CreateFromControl = new uiElement
  Call CreateFromControl.protInit(parent,ctrl)
End Function

Public Function CreateFromType(ByVal parent As uiIMessagable, ByVal oControls As MSForms.controls, ByVal iCtrlType As EUIElementType, Optional ByVal sName As String, Optional ByVal sCaption as string, Optional ByVal fLeft As Double = 0, Optional ByVal fTop As Double = 0, Optional ByVal fWidth As Double, Optional ByVal fHeight As Double) As uiElement
  set CreateFromType = new uiElement
  Dim ctrl as Control
  select case iCtrlType
    Case uiUserForm:      set ctrl = oControls.Add("Forms.Form.1",          sName)
    Case uiFrame:         set ctrl = oControls.Add("Forms.Frame.1",         sName)
    Case uiMultiPage:     set ctrl = oControls.Add("Forms.MultiPage.1",     sName)
    Case uiLabel:         set ctrl = oControls.Add("Forms.Label.1",         sName)
    Case uiImage:         set ctrl = oControls.Add("Forms.Image.1",         sName)
    Case uiTabStrip:      set ctrl = oControls.Add("Forms.TabStrip.1",      sName)
    Case uiTextBox:       set ctrl = oControls.Add("Forms.TextBox.1",       sName)
    Case uiComboBox:      set ctrl = oControls.Add("Forms.ComboBox.1",      sName)
    Case uiListBox:       set ctrl = oControls.Add("Forms.ListBox.1",       sName)
    Case uiCheckBox:      set ctrl = oControls.Add("Forms.CheckBox.1",      sName)
    Case uiOptionButton:  set ctrl = oControls.Add("Forms.OptionButton.1",  sName)
    Case uiToggleButton:  set ctrl = oControls.Add("Forms.ToggleButton.1",  sName)
    Case uiScrollBar:     set ctrl = oControls.Add("Forms.ScrollBar.1",     sName)
    Case uiSpinButton:    set ctrl = oControls.Add("Forms.SpinButton.1",    sName)
    Case uiCommandButton: set ctrl = oControls.Add("Forms.CommandButton.1", sName)
  end select
  if sName <> "" then ctrl.name = sName
  if sCaption <> "" then ctrl.caption = sCaption
  ctrl.Left = fLeft
  ctrl.Top  = fTop
  if fWidth <> 0  then ctrl.Width  = fWidth
  if fHeight <> 0 then ctrl.Height = fHeight
  Call CreateFromType.protInit(parent,ctrl)
End Function

'Obtain type of input control
Public Property Get uiType() as EUIElementType
  uiType = pType
End Property

'Obtain control object
Public Property Get Control() as Control
  set Control = pControl
End Property

'Obtain ui object raw item
Public Property Get uiObject() as Object
  select case pType
    Case uiUserForm:       set uiObject = pUserForm
    Case uiFrame:          set uiObject = pFrame
    Case uiMultiPage:      set uiObject = pMultiPage
    Case uiLabel:          set uiObject = pLabel
    Case uiImage:          set uiObject = pImage
    Case uiTabStrip:       set uiObject = pTabStrip
    Case uiTextBox:        set uiObject = pTextBox
    Case uiComboBox:       set uiObject = pComboBox
    Case uiListBox:        set uiObject = pListBox
    Case uiCheckBox:       set uiObject = pCheckBox
    Case uiOptionButton:   set uiObject = pOptionButton
    Case uiToggleButton:   set uiObject = pToggleButton
    Case uiScrollBar:      set uiObject = pScrollBar
    Case uiSpinButton:     set uiObject = pSpinButton
    Case uiCommandButton:  set uiObject = pCommandButton
  end select
End Property

'Get/Set Value
Public Property Get Value() as Variant
  select case pType
    Case uiTextBox, uiComboBox, uiListBox, uiCheckBox, uiOptionButton, uiToggleButton, uiScrollBar, uiSpinButton    Value = uiObject.Value
      Value = uiObject.value
    Case uiLabel
      Value = uiObject.caption
  end select
End Property
Public Property Let Value(v as Variant)
  uiObject.Value = v
End Property

'Get/Set Position Left
Public Property Get Left() as Double
  Left = pControl.left
End Property
Public Property Let Left(fLeft as Double)
  pControl.Left = fLeft
  Call pParent.Message(Me, uiElementEventMove, pControl.Left, pControl.Top)
End Property

'Get/Set Position Top
Public Property Get Top() as Double
  Top = pControl.Top
End Property
Public Property Let Top(fTop as Double)
  pControl.Top = fTop
  Call pParent.Message(Me, uiElementEventMove, pControl.Left, pControl.Top)
End Property

'Get/Set Size Width
Public Property Get Width() as Double
  Width = pControl.Width
End Property
Public Property Let Width(fWidth as Double)
  pControl.Width = fWidth
  Call pParent.Message(Me, uiElementEventResize, pControl.Width, pControl.Height)
End Property

'Get/Set Size Height
Public Property Get Height() as Double
  Height = pControl.Height
End Property
Public Property Let Height(fHeight as Double)
  pControl.Height = fHeight
  Call pParent.Message(Me, uiElementEventResize, pControl.Width, pControl.Height)
End Property

'Getter for id
Public Property Get id() as string
  id=pID
End Property


'*********************
'* PROTECTED METHODS *
'*********************

'Initialise input control
'@protected
Friend Sub protInit(parent as uiIMessagable, ctrl as Control)
  set pParent = parent
  set pControl = ctrl
  pID = getGUID()
  if TypeOf ctrl is MSForms.UserForm then
    set pUserForm = ctrl
    pType = EUIElementType.uiUserForm
    Exit Sub
  end if
  if TypeOf ctrl is MSForms.Frame then
    set pFrame = ctrl
    pType = EUIElementType.uiFrame
    Exit Sub
  end if
  if TypeOf ctrl is MSForms.MultiPage then
    set pMultiPage = ctrl
    pType = EUIElementType.uiMultiPage
    Exit Sub
  end if
  if TypeOf ctrl is MSForms.Label then
    set pLabel = ctrl
    pType = EUIElementType.uiLabel
    Exit Sub
  end if
  if TypeOf ctrl is MSForms.Image then
    set pImage = ctrl
    pType = EUIElementType.uiImage
    Exit Sub
  end if
  if TypeOf ctrl is MSForms.TabStrip then
    set pTabStrip = ctrl
    pType = EUIElementType.uiTabStrip
    Exit Sub
  end if
  if TypeOf ctrl is MSForms.TextBox then
    set pTextBox = ctrl
    pType = EUIElementType.uiTextBox
    Exit Sub
  end if
  if TypeOf ctrl is MSForms.ComboBox then
    set pComboBox = ctrl
    pType = EUIElementType.uiComboBox
    Exit Sub
  end if
  if TypeOf ctrl is MSForms.ListBox then
    set pListBox = ctrl
    pType = EUIElementType.uiListBox
    Exit Sub
  end if
  if TypeOf ctrl is MSForms.CheckBox then
    set pCheckBox = ctrl
    pType = EUIElementType.uiCheckBox
    Exit Sub
  end if
  if TypeOf ctrl is MSForms.OptionButton then
    set pOptionButton = ctrl
    pType = EUIElementType.uiOptionButton
    Exit Sub
  end if
  if TypeOf ctrl is MSForms.ToggleButton then
    set pToggleButton = ctrl
    pType = EUIElementType.uiToggleButton
    Exit Sub
  end if
  if TypeOf ctrl is MSForms.ScrollBar then
    set pScrollBar = ctrl
    pType = EUIElementType.uiScrollBar
    Exit Sub
  end if
  if TypeOf ctrl is MSForms.SpinButton then
    set pSpinButton = ctrl
    pType = EUIElementType.uiSpinButton
    Exit Sub
  end if
  if TypeOf ctrl is MSForms.CommandButton then
    set pCommandButton = ctrl
    pType = EUIElementType.uiCommandButton
    Exit Sub
  end if
End Sub

Private Function getGUID() As String
  getGUID = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx"
  getGUID = Replace(getGUID, "y", Hex(Rnd() And &H3 Or &H8))
  Dim i As Long: For i = 1 To 30
    getGUID = Replace(getGUID, "x", Hex$(CLng(Rnd() * 15.9999)), 1, 1)
  Next
End Function






'TODO: Consider calculating and passing X and Y coords of mouse press event
'@remark The following controls are missing the click event: pScrollBar, pSpinButton
Private Sub pUserForm_Click():      Call pParent.Message(Me, EUIElementMessage.uiElementEventMouseClick ) : End Sub
Private Sub pFrame_Click():         Call pParent.Message(Me, EUIElementMessage.uiElementEventMouseClick ) : End Sub
Private Sub pLabel_Click():         Call pParent.Message(Me, EUIElementMessage.uiElementEventMouseClick ) : End Sub
Private Sub pComboBox_Click():      Call pParent.Message(Me, EUIElementMessage.uiElementEventMouseClick ) : End Sub
Private Sub pListBox_Click():       Call pParent.Message(Me, EUIElementMessage.uiElementEventMouseClick ) : End Sub
Private Sub pCheckBox_Click():      Call pParent.Message(Me, EUIElementMessage.uiElementEventMouseClick ) : End Sub
Private Sub pOptionButton_Click():  Call pParent.Message(Me, EUIElementMessage.uiElementEventMouseClick ) : End Sub
Private Sub pToggleButton_Click():  Call pParent.Message(Me, EUIElementMessage.uiElementEventMouseClick ) : End Sub
Private Sub pCommandButton_Click(): Call pParent.Message(Me, EUIElementMessage.uiElementEventMouseClick ) : End Sub
Private Sub pMultiPage_Click(ByVal index As Long)
  Call pParent.Message(Me, EUIElementMessage.uiElementEventMouseClick)
  Call pParent.Message(Me, EUIElementMessage.uiElementEventTabSelected, index)
End Sub
Private Sub pTabStrip_Click(ByVal Index As Long)
  Call pParent.Message(Me, EUIElementMessage.uiElementEventMouseClick )
  Call pParent.Message(Me, EUIElementMessage.uiElementEventTabSelected, index)
End Sub
Private Sub pImage_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call pParent.Message(Me, EUIElementMessage.uiElementEventMouseClick, X, Y)
End Sub
Private Sub pTextBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call pParent.Message(Me, EUIElementMessage.uiElementEventMouseClick, X, Y)
End Sub

'Parent object events
Private Sub pUserForm_AddControl(ByVal Control As MSForms.Control): Call pParent.Message(Me, EUIElementMessage.uiElementEventAddChild, control, 0) : End Sub
Private Sub pFrame_AddControl(ByVal Control As MSForms.Control): Call pParent.Message(Me, EUIElementMessage.uiElementEventAddChild, control, 0) : End Sub
Private Sub pMultiPage_AddControl(ByVal Index As Long, ByVal Control As MSForms.Control): Call pParent.Message(Me, EUIElementMessage.uiElementEventAddChild, control, index) End Sub
Private Sub pMultiPage_AddControl(ByVal Control As MSForms.Control): Call pParent.Message(Me, EUIElementMessage.uiElementEventAddChild, control, 0) : End Sub

'Change events
Private Sub pMultiPage_Change():    Call pParent.Message(Me, EUIElementMessage.uiElementEventChange, uiObject.value)  : End Sub
Private Sub pTabStrip_Change():     Call pParent.Message(Me, EUIElementMessage.uiElementEventChange, uiObject.value)  : End Sub
Private Sub pCheckBox_Change():     Call pParent.Message(Me, EUIElementMessage.uiElementEventChange, uiObject.value ) : End Sub
Private Sub pComboBox_Change():     Call pParent.Message(Me, EUIElementMessage.uiElementEventChange, uiObject.value ) : End Sub
Private Sub pListBox_Change():      Call pParent.Message(Me, EUIElementMessage.uiElementEventChange, uiObject.value ) : End Sub
Private Sub pOptionButton_Change(): Call pParent.Message(Me, EUIElementMessage.uiElementEventChange, uiObject.value ) : End Sub
Private Sub pScrollBar_Change():    Call pParent.Message(Me, EUIElementMessage.uiElementEventChange, uiObject.value ) : End Sub
Private Sub pSpinButton_Change():   Call pParent.Message(Me, EUIElementMessage.uiElementEventChange, uiObject.value ) : End Sub
Private Sub pTextBox_Change():      Call pParent.Message(Me, EUIElementMessage.uiElementEventChange, uiObject.value ) : End Sub
Private Sub pToggleButton_Change(): Call pParent.Message(Me, EUIElementMessage.uiElementEventChange, uiObject.value ) : End Sub

'MouseMove events