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
'@constructor
'@param ctrl - The control to create the UI Element from
'@param callback as stdICallable<(el: stdUIElement, event: EUIElementEvent, params: Variant<Array<Variant>>)=>void> - The callback object to use for event handling
'@returns - The UI Element
Public Function CreateFromControl(ByVal ctrl As Control, callback As stdICallable) As stdUIElement
  With stdError.getSentry("stdUIElement#CreateFromControl", "ctrl", ctrl, "callback", callback)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Set CreateFromControl = New stdUIElement
      If Not callback Is Nothing Then Call BindGlobals(callback)
      Call CreateFromControl.protInit(ctrl, callback)
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Creating a UI Element from a type
'@constructor
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
Public Function CreateFromType(ByVal oControls As MSForms.Controls, ByVal iCtrlType As EUIElementType, Optional ByVal sName As String, Optional ByVal Caption As String, Optional ByVal Value As Variant, Optional ByVal callback As stdICallable, Optional ByVal fLeft As Double = 0, Optional ByVal fTop As Double = 0, Optional ByVal fWidth As Double, Optional ByVal fHeight As Double) As stdUIElement
  With stdError.getSentry("stdUIElement#CreateFromType", "oControls", oControls, "iCtrlType", iCtrlType, "sName", sName, "Caption", Caption, "Value", Value, "callback", callback, "fLeft", fLeft, "fTop", fTop, "fWidth", fWidth, "fHeight", fHeight)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
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
      If Not callback Is Nothing Then Call BindGlobals(callback)
      Call CreateFromType.protInit(ctrl, callback)
      If Not IsMissing(Value) Then
        If IsObject(Value) Then
          Set CreateFromType.Value = Value
        Else
          CreateFromType.Value = Value
        End If
      End If
      CreateFromType.Caption = Caption
      CreateFromType.Left = fLeft
      CreateFromType.Top = fTop
      If fWidth <> 0 Then CreateFromType.Width = fWidth
      If fHeight <> 0 Then CreateFromType.Height = fHeight
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Obtain type of input control
'@returns - The type of input control
Public Property Get uiType() As EUIElementType
  With stdError.getSentry("stdUIElement#uiType[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      uiType = this.iType
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Obtain control object
'@returns - The control object
Public Property Get Control() As Control
  With stdError.getSentry("stdUIElement#Control[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Set Control = this.Control
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Obtain ui object raw item
'@returns - The ui object raw item / event sink
Public Property Get uiObject() As Object
  With stdError.getSentry("stdUIElement#uiObject[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
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
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Get/Set Value
'@returns - The value of the control
Public Property Get Value() As Variant
  With stdError.getSentry("stdUIElement#Value[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Select Case this.iType
        Case uiTextBox, uiCombobox, uiListBox, uiCheckBox, uiOptionButton, uiToggleButton, uiScrollbar, uiSpinButton
          Value = uiObject.Value
        Case uiImage
          Set Value = uiObject.Picture
      End Select
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property
Public Property Let Value(v As Variant)
  With stdError.getSentry("stdUIElement#Value[Let]", "v", v)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Select Case this.iType
        Case uiTextBox, uiCombobox, uiListBox, uiCheckBox, uiOptionButton, uiToggleButton, uiScrollbar, uiSpinButton
          uiObject.Value = v
      End Select
      Call callEventCB(ObjRef(Me), uiElementEventChange, Array(v))
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property
Public Property Set Value(v As Variant)
  With stdError.getSentry("stdUIElement#Value[Set]", "v", v)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Select Case this.iType
        Case uiImage
          Set uiObject.Picture = v
      End Select
      Call callEventCB(ObjRef(Me), uiElementEventChange, Array(v))
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property


'Get/Set Caption - The visible text of a control
'@returns - The caption of the control
Public Property Get Caption() As String
  With stdError.getSentry("stdUIElement#Caption[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Select Case this.iType
        Case uiUserForm, uiFrame, uiLabel, uiCheckBox, uiOptionButton, uiToggleButton, uiCommandButton
          Caption = this.Control.Caption
      End Select
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property
Public Property Let Caption(v As String)
  With stdError.getSentry("stdUIElement#Caption[Let]", "v", v)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Select Case this.iType
        Case uiUserForm, uiFrame, uiLabel, uiCheckBox, uiOptionButton, uiToggleButton, uiCommandButton
          this.Control.Caption = v
      End Select
      Call callEventCB(ObjRef(Me), uiElementEventChange, Array(v))
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Get/Set Position Left
'@returns - The left position of the control
Public Property Get Left() As Double
  With stdError.getSentry("stdUIElement#Left[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Left = this.Left
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property
Public Property Let Left(fLeft As Double)
  With stdError.getSentry("stdUIElement#Left[Let]", "fLeft", fLeft)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      this.Left = fLeft
      Select Case this.iType
        Case uiTextBox, uiCombobox, uiListBox, uiOptionButton, uiToggleButton, uiScrollbar, uiSpinButton, uiLabel
          uiObject.Left = this.Left
        Case uiCheckBox
          uiObject.Left = this.Left + IIf(this.Width <= 12, 0, this.Width / 2 - 6)
      End Select
      Call callEventCB(ObjRef(Me), uiElementEventMove, Array(this.Left, this.Control.Top))
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Get/Set Position Top
'@returns - The top position of the control
Public Property Get Top() As Double
  With stdError.getSentry("stdUIElement#Top[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Top = this.Control.Top
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property
Public Property Let Top(fTop As Double)
  With stdError.getSentry("stdUIElement#Top[Let]", "fTop", fTop)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      this.Control.Top = fTop
      Call callEventCB(ObjRef(Me), uiElementEventMove, Array(this.Control.Left, this.Control.Top))
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Get/Set Size Width
'@returns - The width of the control
Public Property Get Width() As Double
  With stdError.getSentry("stdUIElement#Width[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Width = this.Width
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property
Public Property Let Width(fWidth As Double)
  With stdError.getSentry("stdUIElement#Width[Let]", "fWidth", fWidth)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      this.Width = fWidth
      Left = this.Left
      uiObject.Width = fWidth
      Call callEventCB(ObjRef(Me), uiElementEventResize, Array(this.Control.Width, this.Control.Height))
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Get/Set Size Height
'@returns - The height of the control
Public Property Get Height() As Double
  With stdError.getSentry("stdUIElement#Height[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Height = this.Control.Height
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property
Public Property Let Height(fHeight As Double)
  With stdError.getSentry("stdUIElement#Height[Let]", "fHeight", fHeight)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      this.Control.Height = fHeight
      Call callEventCB(ObjRef(Me), uiElementEventResize, Array(this.Control.Width, this.Control.Height))
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Getter for id
'@returns - The id of the control. This is a generated GUID
Public Property Get id() As String
  With stdError.getSentry("stdUIElement#id[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      id = this.id
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property


'*********************
'* PROTECTED METHODS *
'*********************

'Initialise input control
'@protected
Friend Sub protInit(ctrl As Control, callback As stdICallable)
  With stdError.getSentry("stdUIElement#protInit", "ctrl", ctrl, "callback", callback)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
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
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Get a GUID
'@returns - A new GUID
Private Function getGUID() As String
  With stdError.getSentry("stdUIElement#getGUID")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      getGUID = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx"
      getGUID = Replace(getGUID, "y", Hex(Rnd() And &H3 Or &H8))
      Dim i As Long: For i = 1 To 30
        getGUID = Replace(getGUID, "x", Hex$(CLng(Rnd() * 15.9999)), 1, 1)
      Next
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function



'Bind globals to the callback object
Private Sub BindGlobals(ByVal query As stdICallable)
  With stdError.getSentry("stdUIElement#BindGlobals", "query", query)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim bSuccess As Boolean
      Call query.SendMessage("bindGlobal", bSuccess, Array("EUIElementEvent", stdUIElement.protUIElementEventLookup()))
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Get a dictionary of event names and their corresponding event handler
'@protected
'@returns - A dictionary of event names and their corresponding event handler
Friend Function protUIElementEventLookup() As Object
  With stdError.getSentry("stdUIElement#protUIElementEventLookup")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Static obj As Object
      If obj Is Nothing Then
        Set obj = CreateObject("Scripting.Dictionary")
        obj.CompareMode = vbTextCompare
        obj.add "uiElementEventChange", uiElementEventChange
        obj.add "uiElementEventMove", uiElementEventMove
        obj.add "uiElementEventResize", uiElementEventResize
        obj.add "uiElementEventMouseClick", uiElementEventMouseClick
        obj.add "uiElementEventMouseDblClick", uiElementEventMouseDblClick
        obj.add "uiElementEventMouseFocus", uiElementEventMouseFocus
        obj.add "uiElementEventMouseMove", uiElementEventMouseMove
        obj.add "uiElementEventMouseDown", uiElementEventMouseDown
        obj.add "uiElementEventMouseUp", uiElementEventMouseUp
        obj.add "uiElementEventTabSelected", uiElementEventTabSelected
        obj.add "uiElementEventAddChild", uiElementEventAddChild
        obj.add "uiElementEventKeyDown", uiElementEventKeyDown
        obj.add "uiElementEventKeyUp", uiElementEventKeyUp
        obj.add "uiElementEventKeyPress", uiElementEventKeyPress
      End If
      Set protUIElementEventLookup = obj
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Return a reference to an object
'@param obj - The object to return a reference to
'@returns - A reference to the object
Private Function ObjRef(ByVal obj As Object) As Object
  With stdError.getSentry("stdUIElement#ObjRef", "obj", obj)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Set ObjRef = obj
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function









'TODO: Consider calculating and passing X and Y coords of mouse press event
'@remark The following controls are missing the click event: pScrollBar, pSpinButton
Private Sub pUserForm_Click()
  With stdError.getSentry("stdUIElement#pUserForm_Click")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseClick): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pFrame_Click()
  With stdError.getSentry("stdUIElement#pFrame_Click")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :         Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseClick): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pLabel_Click()
  With stdError.getSentry("stdUIElement#pLabel_Click")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :         Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseClick): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pComboBox_Click()
  With stdError.getSentry("stdUIElement#pComboBox_Click")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseClick): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pListBox_Click()
  With stdError.getSentry("stdUIElement#pListBox_Click")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseClick): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pCheckBox_Click()
  With stdError.getSentry("stdUIElement#pCheckBox_Click")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseClick): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pOptionButton_Click()
  With stdError.getSentry("stdUIElement#pOptionButton_Click")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseClick): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pToggleButton_Click()
  With stdError.getSentry("stdUIElement#pToggleButton_Click")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseClick): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pCommandButton_Click()
  With stdError.getSentry("stdUIElement#pCommandButton_Click")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    : Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseClick): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pMultiPage_Click(ByRef Index As Long)
  With stdError.getSentry("stdUIElement#pMultiPage_Click", "Index", Index)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseClick)
      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventTabSelected, Array(Index))
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pTabStrip_Click(ByRef Index As Long)
  With stdError.getSentry("stdUIElement#pTabStrip_Click", "Index", Index)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseClick)
      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventTabSelected, Array(Index))
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Parent object events
Private Sub pUserForm_AddControl(ByVal Control As MSForms.Control)
  With stdError.getSentry("stdUIElement#pUserForm_AddControl", "Control", Control)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventAddChild, Array(Control, 0)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pFrame_AddControl(ByVal Control As MSForms.Control)
  With stdError.getSentry("stdUIElement#pFrame_AddControl", "Control", Control)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                          Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventAddChild, Array(Control, 0)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pMultiPage_AddControl(ByVal Index As Long, ByVal Control As MSForms.Control)
  With stdError.getSentry("stdUIElement#pMultiPage_AddControl", "Index", Index, "Control", Control)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    : Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventAddChild, Array(Control, Index)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub


'Change events
Private Sub pMultiPage_Change()
  With stdError.getSentry("stdUIElement#pMultiPage_Change")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :    Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventChange, Array(uiObject.Value)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pTabStrip_Change()
  With stdError.getSentry("stdUIElement#pTabStrip_Change")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :     Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventChange, Array(uiObject.Value)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pCheckBox_Change()
  With stdError.getSentry("stdUIElement#pCheckBox_Change")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :     Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventChange, Array(uiObject.Value)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pComboBox_Change()
  With stdError.getSentry("stdUIElement#pComboBox_Change")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :     Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventChange, Array(uiObject.Value)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pListBox_Change()
  With stdError.getSentry("stdUIElement#pListBox_Change")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventChange, Array(uiObject.Value)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pOptionButton_Change()
  With stdError.getSentry("stdUIElement#pOptionButton_Change")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    : Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventChange, Array(uiObject.Value)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pScrollBar_Change()
  With stdError.getSentry("stdUIElement#pScrollBar_Change")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :    Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventChange, Array(uiObject.Value)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pSpinButton_Change()
  With stdError.getSentry("stdUIElement#pSpinButton_Change")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :   Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventChange, Array(uiObject.Value)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pTextBox_Change()
  With stdError.getSentry("stdUIElement#pTextBox_Change")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventChange, Array(uiObject.Value)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pToggleButton_Change()
  With stdError.getSentry("stdUIElement#pToggleButton_Change")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    : Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventChange, Array(uiObject.Value)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'*******************
'* KEYBOARD EVENTS *
'*******************

Private Sub pToggleButton_KeyDown(ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
  With stdError.getSentry("stdUIElement#pToggleButton_KeyDown", "KeyCode", KeyCode, "Shift", Shift)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyDown, Array(KeyCode, Shift)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pUserForm_KeyDown(ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
  With stdError.getSentry("stdUIElement#pUserForm_KeyDown", "KeyCode", KeyCode, "Shift", Shift)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyDown, Array(KeyCode, Shift)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pCheckBox_KeyDown(ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
  With stdError.getSentry("stdUIElement#pCheckBox_KeyDown", "KeyCode", KeyCode, "Shift", Shift)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyDown, Array(KeyCode, Shift)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pComboBox_KeyDown(ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
  With stdError.getSentry("stdUIElement#pComboBox_KeyDown", "KeyCode", KeyCode, "Shift", Shift)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyDown, Array(KeyCode, Shift)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pCommandButton_KeyDown(ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
  With stdError.getSentry("stdUIElement#pCommandButton_KeyDown", "KeyCode", KeyCode, "Shift", Shift)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    : Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyDown, Array(KeyCode, Shift)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pFrame_KeyDown(ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
  With stdError.getSentry("stdUIElement#pFrame_KeyDown", "KeyCode", KeyCode, "Shift", Shift)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :         Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyDown, Array(KeyCode, Shift)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pListBox_KeyDown(ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
  With stdError.getSentry("stdUIElement#pListBox_KeyDown", "KeyCode", KeyCode, "Shift", Shift)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyDown, Array(KeyCode, Shift)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pMultiPage_KeyDown(ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
  With stdError.getSentry("stdUIElement#pMultiPage_KeyDown", "KeyCode", KeyCode, "Shift", Shift)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :     Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyDown, Array(KeyCode, Shift)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pOptionButton_KeyDown(ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
  With stdError.getSentry("stdUIElement#pOptionButton_KeyDown", "KeyCode", KeyCode, "Shift", Shift)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyDown, Array(KeyCode, Shift)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pScrollBar_KeyDown(ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
  With stdError.getSentry("stdUIElement#pScrollBar_KeyDown", "KeyCode", KeyCode, "Shift", Shift)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :     Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyDown, Array(KeyCode, Shift)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pSpinButton_KeyDown(ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
  With stdError.getSentry("stdUIElement#pSpinButton_KeyDown", "KeyCode", KeyCode, "Shift", Shift)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :    Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyDown, Array(KeyCode, Shift)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pTabStrip_KeyDown(ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
  With stdError.getSentry("stdUIElement#pTabStrip_KeyDown", "KeyCode", KeyCode, "Shift", Shift)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyDown, Array(KeyCode, Shift)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pTextBox_KeyDown(ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
  With stdError.getSentry("stdUIElement#pTextBox_KeyDown", "KeyCode", KeyCode, "Shift", Shift)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyDown, Array(KeyCode, Shift)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

Private Sub pToggleButton_KeyPress(ByRef KeyAscii As MSForms.ReturnInteger)
  With stdError.getSentry("stdUIElement#pToggleButton_KeyPress", "KeyAscii", KeyAscii)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyPress, Array(KeyAscii)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pUserForm_KeyPress(ByRef KeyAscii As MSForms.ReturnInteger)
  With stdError.getSentry("stdUIElement#pUserForm_KeyPress", "KeyAscii", KeyAscii)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyPress, Array(KeyAscii)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pCheckBox_KeyPress(ByRef KeyAscii As MSForms.ReturnInteger)
  With stdError.getSentry("stdUIElement#pCheckBox_KeyPress", "KeyAscii", KeyAscii)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyPress, Array(KeyAscii)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pComboBox_KeyPress(ByRef KeyAscii As MSForms.ReturnInteger)
  With stdError.getSentry("stdUIElement#pComboBox_KeyPress", "KeyAscii", KeyAscii)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyPress, Array(KeyAscii)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pCommandButton_KeyPress(ByRef KeyAscii As MSForms.ReturnInteger)
  With stdError.getSentry("stdUIElement#pCommandButton_KeyPress", "KeyAscii", KeyAscii)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    : Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyPress, Array(KeyAscii)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pFrame_KeyPress(ByRef KeyAscii As MSForms.ReturnInteger)
  With stdError.getSentry("stdUIElement#pFrame_KeyPress", "KeyAscii", KeyAscii)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :         Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyPress, Array(KeyAscii)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pListBox_KeyPress(ByRef KeyAscii As MSForms.ReturnInteger)
  With stdError.getSentry("stdUIElement#pListBox_KeyPress", "KeyAscii", KeyAscii)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyPress, Array(KeyAscii)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pMultiPage_KeyPress(ByRef KeyAscii As MSForms.ReturnInteger)
  With stdError.getSentry("stdUIElement#pMultiPage_KeyPress", "KeyAscii", KeyAscii)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :     Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyPress, Array(KeyAscii)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pOptionButton_KeyPress(ByRef KeyAscii As MSForms.ReturnInteger)
  With stdError.getSentry("stdUIElement#pOptionButton_KeyPress", "KeyAscii", KeyAscii)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyPress, Array(KeyAscii)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pScrollBar_KeyPress(ByRef KeyAscii As MSForms.ReturnInteger)
  With stdError.getSentry("stdUIElement#pScrollBar_KeyPress", "KeyAscii", KeyAscii)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :     Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyPress, Array(KeyAscii)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pSpinButton_KeyPress(ByRef KeyAscii As MSForms.ReturnInteger)
  With stdError.getSentry("stdUIElement#pSpinButton_KeyPress", "KeyAscii", KeyAscii)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :    Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyPress, Array(KeyAscii)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pTabStrip_KeyPress(ByRef KeyAscii As MSForms.ReturnInteger)
  With stdError.getSentry("stdUIElement#pTabStrip_KeyPress", "KeyAscii", KeyAscii)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyPress, Array(KeyAscii)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pTextBox_KeyPress(ByRef KeyAscii As MSForms.ReturnInteger)
  With stdError.getSentry("stdUIElement#pTextBox_KeyPress", "KeyAscii", KeyAscii)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyPress, Array(KeyAscii)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

Private Sub pToggleButton_KeyUp(ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
  With stdError.getSentry("stdUIElement#pToggleButton_KeyUp", "KeyCode", KeyCode, "Shift", Shift)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyUp, Array(KeyCode, Shift)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pUserForm_KeyUp(ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
  With stdError.getSentry("stdUIElement#pUserForm_KeyUp", "KeyCode", KeyCode, "Shift", Shift)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyUp, Array(KeyCode, Shift)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pCheckBox_KeyUp(ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
  With stdError.getSentry("stdUIElement#pCheckBox_KeyUp", "KeyCode", KeyCode, "Shift", Shift)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyUp, Array(KeyCode, Shift)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pComboBox_KeyUp(ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
  With stdError.getSentry("stdUIElement#pComboBox_KeyUp", "KeyCode", KeyCode, "Shift", Shift)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyUp, Array(KeyCode, Shift)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pCommandButton_KeyUp(ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
  With stdError.getSentry("stdUIElement#pCommandButton_KeyUp", "KeyCode", KeyCode, "Shift", Shift)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    : Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyUp, Array(KeyCode, Shift)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pFrame_KeyUp(ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
  With stdError.getSentry("stdUIElement#pFrame_KeyUp", "KeyCode", KeyCode, "Shift", Shift)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :         Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyUp, Array(KeyCode, Shift)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pListBox_KeyUp(ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
  With stdError.getSentry("stdUIElement#pListBox_KeyUp", "KeyCode", KeyCode, "Shift", Shift)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyUp, Array(KeyCode, Shift)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pMultiPage_KeyUp(ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
  With stdError.getSentry("stdUIElement#pMultiPage_KeyUp", "KeyCode", KeyCode, "Shift", Shift)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :     Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyUp, Array(KeyCode, Shift)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pOptionButton_KeyUp(ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
  With stdError.getSentry("stdUIElement#pOptionButton_KeyUp", "KeyCode", KeyCode, "Shift", Shift)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyUp, Array(KeyCode, Shift)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pScrollBar_KeyUp(ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
  With stdError.getSentry("stdUIElement#pScrollBar_KeyUp", "KeyCode", KeyCode, "Shift", Shift)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :     Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyUp, Array(KeyCode, Shift)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pSpinButton_KeyUp(ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
  With stdError.getSentry("stdUIElement#pSpinButton_KeyUp", "KeyCode", KeyCode, "Shift", Shift)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :    Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyUp, Array(KeyCode, Shift)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pTabStrip_KeyUp(ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
  With stdError.getSentry("stdUIElement#pTabStrip_KeyUp", "KeyCode", KeyCode, "Shift", Shift)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyUp, Array(KeyCode, Shift)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pTextBox_KeyUp(ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
  With stdError.getSentry("stdUIElement#pTextBox_KeyUp", "KeyCode", KeyCode, "Shift", Shift)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventKeyUp, Array(KeyCode, Shift)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'****************
'* MOUSE EVENTS *
'****************

Private Sub pUserForm_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pUserForm_MouseDown", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDown, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pCheckBox_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pCheckBox_MouseDown", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDown, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pComboBox_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pComboBox_MouseDown", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDown, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pCommandButton_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pCommandButton_MouseDown", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDown, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pFrame_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pFrame_MouseDown", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                          Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDown, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pLabel_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pLabel_MouseDown", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                          Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDown, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pListBox_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pListBox_MouseDown", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                        Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDown, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pMultiPage_MouseDown(ByRef Index As Long, ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pMultiPage_MouseDown", "Index", Index, "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    : Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDown, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pOptionButton_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pOptionButton_MouseDown", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                   Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDown, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pTabStrip_MouseDown(ByRef Index As Long, ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pTabStrip_MouseDown", "Index", Index, "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDown, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pToggleButton_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pToggleButton_MouseDown", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                   Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDown, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pImage_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pImage_MouseDown", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseClick) 'There is no MouseClick event for Image, so we raise this here too
      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDown, Array(Button, Shift, X, Y))
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pTextBox_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pTextBox_MouseDown", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseClick) 'There is no MouseClick event for TextBox, so we raise this here too
      Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDown, Array(Button, Shift, X, Y))
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

Private Sub pUserForm_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pUserForm_MouseMove", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseMove, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pCheckBox_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pCheckBox_MouseMove", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseMove, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pComboBox_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pComboBox_MouseMove", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseMove, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pCommandButton_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pCommandButton_MouseMove", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseMove, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pFrame_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pFrame_MouseMove", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                          Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseMove, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pImage_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pImage_MouseMove", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                          Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseMove, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pLabel_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pLabel_MouseMove", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                          Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseMove, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pListBox_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pListBox_MouseMove", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                        Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseMove, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pMultiPage_MouseMove(ByRef Index As Long, ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pMultiPage_MouseMove", "Index", Index, "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    : Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseMove, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pOptionButton_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pOptionButton_MouseMove", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                   Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseMove, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pTabStrip_MouseMove(ByRef Index As Long, ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pTabStrip_MouseMove", "Index", Index, "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseMove, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pTextBox_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pTextBox_MouseMove", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                        Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseMove, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pToggleButton_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pToggleButton_MouseMove", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                   Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseMove, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

Private Sub pUserForm_MouseUp(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pUserForm_MouseUp", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseUp, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pCheckBox_MouseUp(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pCheckBox_MouseUp", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseUp, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pComboBox_MouseUp(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pComboBox_MouseUp", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                       Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseUp, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pCommandButton_MouseUp(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pCommandButton_MouseUp", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseUp, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pFrame_MouseUp(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pFrame_MouseUp", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                          Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseUp, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pImage_MouseUp(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pImage_MouseUp", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                          Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseUp, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pLabel_MouseUp(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pLabel_MouseUp", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                          Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseUp, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pListBox_MouseUp(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pListBox_MouseUp", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                        Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseUp, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pMultiPage_MouseUp(ByRef Index As Long, ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pMultiPage_MouseUp", "Index", Index, "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    : Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseUp, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pOptionButton_MouseUp(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pOptionButton_MouseUp", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                   Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseUp, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pTabStrip_MouseUp(ByRef Index As Long, ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pTabStrip_MouseUp", "Index", Index, "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseUp, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pTextBox_MouseUp(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pTextBox_MouseUp", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                        Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseUp, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pToggleButton_MouseUp(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  With stdError.getSentry("stdUIElement#pToggleButton_MouseUp", "Button", Button, "Shift", Shift, "X", X, "Y", Y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                   Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseUp, Array(Button, Shift, X, Y)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

Private Sub pCheckBox_DblClick(ByRef Cancel As MSForms.ReturnBoolean)
  With stdError.getSentry("stdUIElement#pCheckBox_DblClick", "Cancel", Cancel)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                        Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDblClick, Array(Cancel, 0)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pComboBox_DblClick(ByRef Cancel As MSForms.ReturnBoolean)
  With stdError.getSentry("stdUIElement#pComboBox_DblClick", "Cancel", Cancel)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                        Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDblClick, Array(Cancel, 0)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pCommandButton_DblClick(ByRef Cancel As MSForms.ReturnBoolean)
  With stdError.getSentry("stdUIElement#pCommandButton_DblClick", "Cancel", Cancel)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                   Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDblClick, Array(Cancel, 0)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pFrame_DblClick(ByRef Cancel As MSForms.ReturnBoolean)
  With stdError.getSentry("stdUIElement#pFrame_DblClick", "Cancel", Cancel)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                           Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDblClick, Array(Cancel, 0)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pLabel_DblClick(ByRef Cancel As MSForms.ReturnBoolean)
  With stdError.getSentry("stdUIElement#pLabel_DblClick", "Cancel", Cancel)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                           Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDblClick, Array(Cancel, 0)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pListBox_DblClick(ByRef Cancel As MSForms.ReturnBoolean)
  With stdError.getSentry("stdUIElement#pListBox_DblClick", "Cancel", Cancel)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                         Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDblClick, Array(Cancel, 0)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pOptionButton_DblClick(ByRef Cancel As MSForms.ReturnBoolean)
  With stdError.getSentry("stdUIElement#pOptionButton_DblClick", "Cancel", Cancel)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                    Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDblClick, Array(Cancel, 0)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pTextBox_DblClick(ByRef Cancel As MSForms.ReturnBoolean)
  With stdError.getSentry("stdUIElement#pTextBox_DblClick", "Cancel", Cancel)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                         Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDblClick, Array(Cancel, 0)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pToggleButton_DblClick(ByRef Cancel As MSForms.ReturnBoolean)
  With stdError.getSentry("stdUIElement#pToggleButton_DblClick", "Cancel", Cancel)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                    Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDblClick, Array(Cancel, 0)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pUserForm_DblClick(ByRef Cancel As MSForms.ReturnBoolean)
  With stdError.getSentry("stdUIElement#pUserForm_DblClick", "Cancel", Cancel)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :                        Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDblClick, Array(Cancel, 0)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pMultiPage_DblClick(ByRef Index As Long, ByRef Cancel As MSForms.ReturnBoolean)
  With stdError.getSentry("stdUIElement#pMultiPage_DblClick", "Index", Index, "Cancel", Cancel)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :  Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDblClick, Array(Cancel, Index)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
Private Sub pTabStrip_DblClick(ByRef Index As Long, ByRef Cancel As MSForms.ReturnBoolean)
  With stdError.getSentry("stdUIElement#pTabStrip_DblClick", "Index", Index, "Cancel", Cancel)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    :   Call callEventCB(ObjRef(Me), EUIElementEvent.uiElementEventMouseDblClick, Array(Cancel, Index)): 
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub
'TODO: SpinButton_DblClick using MouseDown
'TODO: ScrollBar_DblClick using MouseDown
'TODO: Image_DblClick using MouseDown

Private Sub callEventCB(ByVal MeRef As stdUIElement, ByVal iMsg As EUIElementEvent, Optional ByVal args As Variant = Empty)
  With stdError.getSentry("stdUIElement#callEventCB", "MeRef", MeRef, "iMsg", iMsg, "args", args)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If IsEmpty(args) Then args = Array()
      If Not this.messageRouter Is Nothing Then
        Call this.messageRouter.Run(MeRef, iMsg, args)
      End If
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub


Private Sub Err_Raise(ByVal number as Long, Optional ByVal source as string = "", Optional ByVal description as string = "")
  Call stdError.Raise(description)
End Sub
