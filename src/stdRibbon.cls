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
  RibbonNullState
  RibbonFullScreenMode
  RibbonShowTabsOnly
  RibbonAlwaysShowRibbon
  RibbonNotVisible
  RibbonProtectedViewMode
  ExcelApplicationDisplayFullScreenMode
  WordViewFullScreenMode
  WordViewReadMode
  
End Enum

'Get/Set the ribbon's state. The ribbon itself comes with the following options:
'* &Full-screen mode
'* Show &tabs only
'* &Always show Ribbon
'The ribbon can also be in a state where it isn't visible at all
'And a state where the application is in full screen mode, thus hiding the ribbon
'This Get/Set property helps cleanly reset the state to a necessary one
Public Property Get State() As ERibbonState
  Dim appProtectedView as Boolean: appProtectedView = Not Application.ActiveProtectedViewWindow Is Nothing
  Dim oApp as Object: Set oApp = Application 'Latebound to avoid compile errors

  Dim xlDisplayFullScreen as Boolean
  Dim wdReadMode as Boolean
  Dim wdViewFullScreen as Boolean
  Dim wdProtectedViewMode as Boolean
  select case Application.Name
    case "Microsoft Excel"
      xlDisplayFullScreen = oApp.DisplayFullScreen
    case "Microsoft Word"
      If Not appProtectedView Then 
        wdViewFullScreen = oApp.ActiveWindow.View.FullScreen
        wdReadMode = oApp.ActiveWindow.View.ReadingLayout
      end if
  end select

  'In this mode state is reversable
  select case true
    case appProtectedView
      State = RibbonProtectedViewMode
    case xlDisplayFullScreen
      State = ExcelApplicationDisplayFullScreenMode
    case wdViewFullScreen
      State = WordViewFullScreenMode
    case wdReadMode
      State = WordViewReadMode
    case Not Application.CommandBars("Ribbon").Visible
      State = RibbonNotVisible
    case Application.CommandBars.GetPressedMso("MinimizeRibbon") 
      State = RibbonShowTabsOnly
    Case Application.CommandBars.GetPressedMso("HideRibbon")
      State = RibbonFullScreenMode
    Case Else  
      State = RibbonAlwaysShowRibbon
  End Select
End Property
Public Property Let State(v As ERibbonState)
  If State = v Then Exit Property
  Dim oApp as Object: Set oApp = Application 'Latebound to avoid compile errors

  'Reset the state to `RibbonAlwaysShowRibbon`
  Select Case State
    Case RibbonAlwaysShowRibbon
      'Do nothing
    case RibbonProtectedViewMode
      Call Application.ActiveProtectedViewWindow.Edit
    case WordViewFullScreenMode
      if Application.Name <> "Microsoft Word" then Err.Raise 1, "stdRibbon#State[Let]", "WordViewFullScreenMode found in non Word Application. This should never happen."
      oApp.ActiveWindow.View.FullScreen = False
    case WordViewReadMode
      if Application.Name <> "Microsoft Word" then Err.Raise 1, "stdRibbon#State[Let]", "WordViewReadMode found in non Word Application. This should never happen."
      oApp.ActiveWindow.View.ReadingLayout = False
    Case ExcelApplicationDisplayFullScreenMode
      if Application.Name <> "Microsoft Excel" then Err.Raise 1, "stdRibbon#State[Let]", "ExcelApplicationDisplayFullScreenMode found in non Excel Application. This should never happen."
      oApp.DisplayFullScreen = False
    Case RibbonNotVisible
      if Application.Name <> "Microsoft Excel" then Err.Raise 1, "stdRibbon#State[Let]", "RibbonNotVisible found in non Excel Application. Mechanism for this state is unknown at this time. Please contact the stdVBA library authors and discuss."
      Call oApp.ExecuteExcel4Macro("Show.ToolBar(""Ribbon"",true)")
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
    Case RibbonProtectedViewMode
      Err.Raise 1, "stdRibbon#State[Let]", "WordProtectedViewMode can only be entered by opening a document in Protected View"
    case WordViewFullScreenMode
      if Application.Name <> "Microsoft Word" then Err.Raise 1, "stdRibbon#State[Let]", "WordViewReadMode is only supported in Word"
      oApp.ActiveWindow.View.FullScreen = true
    case WordViewReadMode
      if Application.Name <> "Microsoft Word" then Err.Raise 1, "stdRibbon#State[Let]", "WordViewReadMode is only supported in Word"
      oApp.ActiveWindow.View.ReadingLayout = true
    Case ExcelApplicationDisplayFullScreenMode
      if Application.Name <> "Microsoft Excel" then Err.Raise 1, "stdRibbon#State[Let]", "ExcelApplicationDisplayFullScreenMode is only supported in Excel"
      oApp.DisplayFullScreen = true
    Case RibbonNotVisible
      if Application.Name <> "Microsoft Excel" then Err.Raise 1, "stdRibbon#State[Let]", "RibbonNotVisible is only supported in Excel currently. A mechanism for Word is yet to be discovered"
      Call oApp.ExecuteExcel4Macro("Show.ToolBar(""Ribbon"",false)")
    Case RibbonFullScreenMode
      Call Application.CommandBars.ExecuteMso("HideRibbon")
    Case RibbonShowTabsOnly
      Call Application.CommandBars.ExecuteMso("MinimizeRibbon")
  End Select
End Property
