VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "uiInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Enum EUIInputMessage
  uiInputEventChange
  uiInputEventMove
  uiInputEventResize
End Enum
Public Enum EUIInputType
  uiTextBox
  uiComboBox
  uiListBox
  uiCheckBox
  uiOptionButton
  uiToggleButton
  uiScrollBar
  uiSpinButton
End Enum

Private pParent  as uiIMessagable
Private pControl as MSForms.Control
Private WithEvents pTextBox       as MSForms.TextBox
Private WithEvents pComboBox      as MSForms.ComboBox
Private WithEvents pListBox       as MSForms.ListBox
Private WithEvents pCheckBox      as MSForms.CheckBox
Private WithEvents pOptionButton  as MSForms.OptionButton
Private WithEvents pToggleButton  as MSForms.ToggleButton
Private WithEvents pScrollBar     as MSForms.ScrollBar
Private WithEvents pSpinButton    as MSForms.SpinButton
Private pType as EUIInputType
Private pID as String

'Creating a input component from 
Public Function CreateFromControl(ByVal parent as uiIMessagable, ByVal ctrl as Control) as uiInput
  set CreateFromControl = new uiInput
  Call CreateFromControl.protInit(parent,ctrl)
End Function

Public Function CreateFromType(ByVal parent As uiIMessagable, ByVal oControls As MSForms.controls, ByVal iCtrlType As EUIInputType, Optional ByVal sName As Variant, Optional ByVal fLeft As Double = 0, Optional ByVal fTop As Double = 0, Optional ByVal fWidth As Double = 0, Optional ByVal fHeight As Double = 0) As uiInput
  set CreateFromType = new uiInput
  Dim ctrl as Control
  select case iCtrlType
    Case uiTextBox:      set ctrl = oControls.Add("Forms.TextBox.1",      sName)
    Case uiComboBox:     set ctrl = oControls.Add("Forms.ComboBox.1",     sName)
    Case uiListBox:      set ctrl = oControls.Add("Forms.ListBox.1",      sName)
    Case uiCheckBox:     set ctrl = oControls.Add("Forms.CheckBox.1",     sName)
    Case uiOptionButton: set ctrl = oControls.Add("Forms.OptionButton.1", sName)
    Case uiToggleButton: set ctrl = oControls.Add("Forms.ToggleButton.1", sName)
    Case uiScrollBar:    set ctrl = oControls.Add("Forms.ScrollBar.1",    sName)
    Case uiSpinButton:   set ctrl = oControls.Add("Forms.SpinButton.1",   sName)
  end select
  ctrl.Left = fLeft
  ctrl.Top  = fTop
  if fWidth <> 0  then ctrl.Width  = fWidth
  if fHeight <> 0 then ctrl.Height = fHeight
  Call CreateFromType.protInit(parent,ctrl)
End Function

'Obtain type of input control
Public Property Get uiType() as EUIInputType
  uiType = pType
End Property

'Obtain control object
Public Property Get Control() as Control
  set Control = pControl
End Property

'Obtain ui object raw item
Public Property Get uiObject() as Object
  select case pType
    Case uiTextBox:      set uiObject = pTextBox 
    Case uiComboBox:     set uiObject = pComboBox 
    Case uiListBox:      set uiObject = pListBox 
    Case uiCheckBox:     set uiObject = pCheckBox 
    Case uiOptionButton: set uiObject = pOptionButton 
    Case uiToggleButton: set uiObject = pToggleButton 
    Case uiScrollBar:    set uiObject = pScrollBar 
    Case uiSpinButton:   set uiObject = pSpinButton 
  end select
End Property

'Get/Set Value
Public Property Get Value() as Variant
  Value = uiObject.Value
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
  Call pParent.Message(Me, uiInputEventMove, pControl.Left, pControl.Top)
End Property

'Get/Set Position Top
Public Property Get Top() as Double
  Top = pControl.Top
End Property
Public Property Let Top(fTop as Double)
  pControl.Top = fTop
  Call pParent.Message(Me, uiInputEventMove, pControl.Left, pControl.Top)
End Property

'Get/Set Size Width
Public Property Get Width() as Double
  Width = pControl.Width
End Property
Public Property Let Width(fWidth as Double)
  pControl.Width = fWidth
  Call pParent.Message(Me, uiInputEventResize, pControl.Width, pControl.Height)
End Property

'Get/Set Size Height
Public Property Get Height() as Double
  Height = pControl.Height
End Property
Public Property Let Height(fHeight as Double)
  pControl.Height = fHeight
  Call pParent.Message(Me, uiInputEventResize, pControl.Width, pControl.Height)
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
  if TypeOf ctrl is MSForms.TextBox then
    set pTextBox = ctrl
    pType = EUIInputType.uiTextBox
    Exit Sub
  end if
  if TypeOf ctrl is MSForms.CheckBox then
    set pCheckBox = ctrl
    pType = EUIInputType.uiCheckBox
    Exit Sub
  end if
  if TypeOf ctrl is MSForms.ComboBox then
    set pComboBox = ctrl
    pType = EUIInputType.uiComboBox
    Exit Sub
  end if
  if TypeOf ctrl is MSForms.ListBox then
    set pListBox = ctrl
    pType = EUIInputType.uiListBox
    Exit Sub
  end if
  if TypeOf ctrl is MSForms.OptionButton then
    set pOptionButton = ctrl
    pType = EUIInputType.uiOptionButton
    Exit Sub
  end if
  if TypeOf ctrl is MSForms.ToggleButton then
    set pToggleButton = ctrl
    pType = EUIInputType.uiToggleButton
    Exit Sub
  end if
  if TypeOf ctrl is MSForms.ScrollBar then
    set pScrollBar = ctrl
    pType = EUIInputType.uiScrollBar
    Exit Sub
  end if
  if TypeOf ctrl is MSForms.SpinButton then
    set pSpinButton = ctrl
    pType = EUIInputType.uiSpinButton
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

Private Sub pCheckBox_Change():     Call pParent.Message(Me, EUIInputMessage.uiInputEventChange, uiObject.value ) : End Sub
Private Sub pComboBox_Change():     Call pParent.Message(Me, EUIInputMessage.uiInputEventChange, uiObject.value ) : End Sub
Private Sub pListBox_Change():      Call pParent.Message(Me, EUIInputMessage.uiInputEventChange, uiObject.value ) : End Sub
Private Sub pOptionButton_Change(): Call pParent.Message(Me, EUIInputMessage.uiInputEventChange, uiObject.value ) : End Sub
Private Sub pScrollBar_Change():    Call pParent.Message(Me, EUIInputMessage.uiInputEventChange, uiObject.value ) : End Sub
Private Sub pSpinButton_Change():   Call pParent.Message(Me, EUIInputMessage.uiInputEventChange, uiObject.value ) : End Sub
Private Sub pTextBox_Change():      Call pParent.Message(Me, EUIInputMessage.uiInputEventChange, uiObject.value ) : End Sub
Private Sub pToggleButton_Change(): Call pParent.Message(Me, EUIInputMessage.uiInputEventChange, uiObject.value ) : End Sub
