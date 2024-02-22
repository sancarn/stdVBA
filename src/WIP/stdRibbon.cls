VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdRibbon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
#If VBA7 Then
    Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDest As LongPtr, ByVal lpSrc As LongPtr, ByVal cbCopy As Long)
#Else
    Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDest As Long, ByVal lpSrc As Long, ByVal cbCopy As Long)
#End If
'Dependencies: stdSettings

Private pRibbon As IRibbonUI
Private pSettings As StdSettings

Private Sub CheckSettings()
  'Instantiate pSettings if not done so already
  if pSettings is nothing then Set pSettings = StdSettings.Create()
End Sub
Private Sub CheckRibbon()
  'If ribbon is gone, restore from pointer
  If pRibbon Is Nothing Then Call CopyMemory(VarPtr(pRibbon), CLng(pSettings.system("RibbonXPtr").Value), 8)
End Sub

Sub Invalidate()
  'Requirements:
  Call CheckSettings
  Call CheckRibbon
  
  'Invalidate ribbon
  Call pRibbon.Invalidate
End Sub
Sub InvalidateControl(ByVal sCtrlId As String)
  'Requirements:
  Call CheckSettings
  Call CheckRibbon

  'Invalidate ribbon control by id
  Call pRibbon.InvalidateControl(sCtrlId)
End Sub

Sub Ribbon_Load(ByRef Rib As IRibbonUI)
  'Requirements:
  Call CheckSettings
  
  'Silently store pointer to ribbon on sheet, such that we can restore the ribbon if vba state loss occurs
  pSettings.system("RibbonXPtr").Value = ObjPtr(Rib)
  
  'Set pRibbon to ribbon object
  Set pRibbon = Rib
End Sub

Sub Tab_GetVisible(control As IRibbonControl, ByRef Visible)
  'Requirements:
  Call CheckSettings
  
  'Set visibility based on user's role
  Visible = CBool(pSettings.user("Role") = "Admin")
End Sub

Sub ToggleButton_getPressed(control As IRibbonControl, ByRef Pressed)
  'Requirements:
  Call CheckSettings
  
  'Set pressed based on pSettings
  Pressed = CBool(pSettings.user("ToggleValue"))
End Sub

Sub ToggleButton_Clicked(control As IRibbonControl, ByRef Pressed)
  'Requirements:
  Call CheckSettings
  
  'Set pressed based on pSettings
  pSettings.user("ToggleValue").Value = Pressed
End Sub

Sub Button_Clicked(control As IRibbonControl)
  Debug.Print "HFRR_Reports_Dashboard"
End Sub

