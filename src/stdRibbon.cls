VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdRibbon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Enum ERibbonState
  RibbonFullScreenMode
  RibbonShowTabsOnly
  RibbonAlwaysShowRibbon
  RibbonNotVisible
  ApplicationDisplayFullScreenMode
End Enum

'Get/Set the ribbon's state. The ribbon itself comes with the following options:
'* &Full-screen mode
'* Show &tabs only
'* &Always show Ribbon
'The ribbon can also be in a state where it isn't visible at all
'And a state where the application is in full screen mode, thus hiding the ribbon
'This Get/Set property helps cleanly reset the state to a necessary one
Public Property Get State() As ERibbonState
  'In this mode state is reversable
  If Application.DisplayFullScreen Then
    State = ApplicationDisplayFullScreenMode
  ElseIf Not Application.CommandBars("Ribbon").Visible Then
    State = RibbonNotVisible
  Else
    Select Case Application.CommandBars("Ribbon").Height
      Case Is < 50
        State = RibbonFullScreenMode
      Case Is < 100
        State = RibbonShowTabsOnly
      Case Else
        State = RibbonAlwaysShowRibbon
    End Select
  End If
End Property
Public Property Let State(v As ERibbonState)
  If State = v Then Exit Property
  
  'Reset the state to `RibbonAlwaysShowRibbon`
  Select Case State
    Case RibbonAlwaysShowRibbon
      'Do nothing
    Case ApplicationDisplayFullScreenMode
      Application.DisplayFullScreen = False
    Case RibbonNotVisible
      Call Application.ExecuteExcel4Macro("Show.ToolBar(""Ribbon"",true)")
    Case RibbonFullScreenMode
      Call Application.CommandBars.ExecuteMso("HideRibbon")
    Case RibbonShowTabsOnly
      Call Application.CommandBars.ExecuteMso("MinimizeRibbon")
  End Select
  DoEvents
  If State <> RibbonAlwaysShowRibbon Then State = RibbonAlwaysShowRibbon
  
  'Change state from `RibbonAlwaysShowRibbon` to `v`
  Select Case v
    Case RibbonAlwaysShowRibbon
      Exit Property
    Case ApplicationDisplayFullScreenMode
      Application.DisplayFullScreen = True
    Case RibbonNotVisible
      Call Application.ExecuteExcel4Macro("Show.ToolBar(""Ribbon"",false)")
    Case RibbonFullScreenMode
      Call Application.CommandBars.ExecuteMso("HideRibbon")
    Case RibbonShowTabsOnly
      Call Application.CommandBars.ExecuteMso("MinimizeRibbon")
  End Select
End Property
