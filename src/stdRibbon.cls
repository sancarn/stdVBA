VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdRibbon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

#Const isWindows = Mac = 0

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
  With stdError.getSentry("stdRibbon#State[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim appProtectedView as Boolean: appProtectedView = Not Application.ActiveProtectedViewWindow Is Nothing
      Dim oApp as Object: Set oApp = Application 'Latebound to avoid compile errors
    
    
      Dim xlDisplayFullScreen as Boolean
      Dim wdReadMode as Boolean
      Dim wdViewFullScreen as Boolean
      Dim wdProtectedViewMode as Boolean
      Dim genRibbonNotVisible as Boolean
      Dim genRibbonHidden as Boolean
      Dim genRibbonMinimized as Boolean
    
      
      select case Application.Name
        case "Microsoft Excel"
          #If isWindows then
            xlDisplayFullScreen = oApp.DisplayFullScreen
            genRibbonNotVisible = Not Application.CommandBars("Ribbon").Visible
            genRibbonHidden = Application.CommandBars.GetPressedMso("HideRibbon")
            genRibbonMinimized = Application.CommandBars.GetPressedMso("MinimizeRibbon")
          #Else
            'None of these are supported on Mac
            'Note: Application.DisplayFullScreen is supported on Mac, but doesn't hide the ribbon!
          #End If
        case "Microsoft Word"
          #If isWindows then
            If Not appProtectedView Then 
              wdViewFullScreen = oApp.ActiveWindow.View.FullScreen
              wdReadMode = oApp.ActiveWindow.View.ReadingLayout
            end if
            genRibbonNotVisible = Not Application.CommandBars("Ribbon").Visible
            genRibbonHidden = Application.CommandBars.GetPressedMso("HideRibbon")
            genRibbonMinimized = Application.CommandBars.GetPressedMso("MinimizeRibbon")
          #Else
            'Only Ribbon Minimized is supported on Mac
            genRibbonMinimized = Application.CommandBars.GetPressedMso("MinimizeRibbon")
          #End If
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
        case genRibbonNotVisible
          State = RibbonNotVisible
        case genRibbonMinimized
          State = RibbonShowTabsOnly
        Case genRibbonHidden
          State = RibbonFullScreenMode
        Case Else  
          State = RibbonAlwaysShowRibbon
      End Select
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property
Public Property Let State(v As ERibbonState)
  With stdError.getSentry("stdRibbon#State[Let]", "v", v)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim iCurrentState as ERibbonState: iCurrentState = State
      If iCurrentState = v Then Exit Property
      Dim oApp as Object: Set oApp = Application 'Latebound to avoid compile errors
    
      'Reset the state to `RibbonAlwaysShowRibbon`
      Select Case iCurrentState
        Case RibbonAlwaysShowRibbon
          'Do nothing
        
        'Else progress toward `RibbonAlwaysShowRibbon`
        case RibbonProtectedViewMode
          Call Application.ActiveProtectedViewWindow.Edit
        case WordViewFullScreenMode
          if Application.Name <> "Microsoft Word" then Err_Raise 1, "stdRibbon#State[Let]", "WordViewFullScreenMode found in non Word Application. This should never happen."
          oApp.ActiveWindow.View.FullScreen = False
        case WordViewReadMode
          if Application.Name <> "Microsoft Word" then Err_Raise 1, "stdRibbon#State[Let]", "WordViewReadMode found in non Word Application. This should never happen."
          oApp.ActiveWindow.View.ReadingLayout = False
        Case ExcelApplicationDisplayFullScreenMode
          if Application.Name <> "Microsoft Excel" then Err_Raise 1, "stdRibbon#State[Let]", "ExcelApplicationDisplayFullScreenMode found in non Excel Application. This should never happen."
          oApp.DisplayFullScreen = False
        Case RibbonNotVisible
          if Application.Name <> "Microsoft Excel" then Err_Raise 1, "stdRibbon#State[Let]", "RibbonNotVisible found in non Excel Application. Mechanism for this state is unknown at this time. Please contact the stdVBA library authors and discuss."
          Call oApp.ExecuteExcel4Macro("Show.ToolBar(""Ribbon"",true)")
        Case RibbonFullScreenMode
          Call Application.CommandBars.ExecuteMso("HideRibbon")
        Case RibbonShowTabsOnly
          #If isWindows then
            Call Application.CommandBars.ExecuteMso("MinimizeRibbon")
          #Else
            'On Mac, it seems MinimizeRibbon is asynchronous, so we need to wait for it to finish
            Dim macMinimizeRibbon as Boolean
            macMinimizeRibbon = Application.CommandBars.GetPressedMso("MinimizeRibbon")
            Call Application.CommandBars.ExecuteMso("MinimizeRibbon")
            Do While macMinimizeRibbon = Application.CommandBars.GetPressedMso("MinimizeRibbon")
              'For some god forsaken reason, DoEvents does not work here... Need to use macscript delays.
              Call MacScript("delay 0")
            Loop
          #End If
      End Select
      DoEvents
      
      'Check if state has changed and if not error.
      If (State = iCurrentState) And (iCurrentState <> RibbonAlwaysShowRibbon) then Err_Raise 1, "stdRibbon#State[Let]", "Ribbon state has not changed away from " & v
      
      'Recursively change state to `RibbonAlwaysShowRibbon`
      If iCurrentState <> RibbonAlwaysShowRibbon Then State = RibbonAlwaysShowRibbon
      
      'Change state from `RibbonAlwaysShowRibbon` to `v`
      Select Case v
        Case RibbonAlwaysShowRibbon
          Exit Property
        Case RibbonProtectedViewMode
          Err_Raise 1, "stdRibbon#State[Let]", "WordProtectedViewMode can only be entered by opening a document in Protected View"
        case WordViewFullScreenMode
          if Application.Name <> "Microsoft Word" then Err_Raise 1, "stdRibbon#State[Let]", "WordViewReadMode is only supported in Word"
          oApp.ActiveWindow.View.FullScreen = true
        case WordViewReadMode
          if Application.Name <> "Microsoft Word" then Err_Raise 1, "stdRibbon#State[Let]", "WordViewReadMode is only supported in Word"
          oApp.ActiveWindow.View.ReadingLayout = true
        Case ExcelApplicationDisplayFullScreenMode
          if Application.Name <> "Microsoft Excel" then Err_Raise 1, "stdRibbon#State[Let]", "ExcelApplicationDisplayFullScreenMode is only supported in Excel"
          oApp.DisplayFullScreen = true
        Case RibbonNotVisible
          if Application.Name <> "Microsoft Excel" then Err_Raise 1, "stdRibbon#State[Let]", "RibbonNotVisible is only supported in Excel currently. A mechanism for Word is yet to be discovered"
          Call oApp.ExecuteExcel4Macro("Show.ToolBar(""Ribbon"",false)")
        Case RibbonFullScreenMode
          Call Application.CommandBars.ExecuteMso("HideRibbon")
        Case RibbonShowTabsOnly
          #If isWindows then
            Call Application.CommandBars.ExecuteMso("MinimizeRibbon")
          #Else
            'On Mac, it seems MinimizeRibbon is asynchronous, so we need to wait for it to finish
            macMinimizeRibbon = Application.CommandBars.GetPressedMso("MinimizeRibbon")
            Call Application.CommandBars.ExecuteMso("MinimizeRibbon")
            Do While macMinimizeRibbon = Application.CommandBars.GetPressedMso("MinimizeRibbon")
              'For some god forsaken reason, DoEvents does not work here... Need to use macscript delays.
              Call MacScript("delay 0")
            Loop
          #End If
      End Select
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property



Private Sub Err_Raise(ByVal number as Long, Optional ByVal source as string = "", Optional ByVal description as string = "")
  Call stdError.Raise(description)
End Sub
