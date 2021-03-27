VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdUserform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Userform wrapper class
'
'Key motivations:
'* Easily obtain x,y, width and height of userform.
'* No standard components for userforms (there are buttons, input boxes etc. but advanced controls are not given (or have dependencies))
'* Inability to modularise userforms (something I want to address)
'* Opening up API behavior e.g. `uf.resizable = true`

'Structure:
'* Compatable userforms must implement a specific interface (not sure if required?)
'* Standard components e.g. stdUFButton, stdUFToggleButton, stdUFCombobox, stdUFDropDown, stdUFLabel, ...
'* Custom components stdUFTimeline, stdUFTreeView, stdUFTreeGrid, stdUFShellView, stdUFGridView, ...

'Example:
'Dim uf as stdUserform: set uf = stdUserform.Create(MyUserform)
'uf.resizable = true
'
'
'Dim cmpWrapper as stdUserform: set cmpWrapper = uf.EmbedComponent(AnotherUserform.Create())
'Dim cmpInternal as AnotherUserform: set cmpInternal = cmpWrapper.object
'cmpWrapper.x = 10
'cmpWrapper.y = 10
'cmpWrapper.width = 300
'cmpWrapper.height = 100
'cmpInternal.somePropertyName = 5


