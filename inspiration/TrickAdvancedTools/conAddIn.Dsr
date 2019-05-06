VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} conAddIn 
   ClientHeight    =   11895
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   16380
   _ExtentX        =   28893
   _ExtentY        =   20981
   _Version        =   393216
   Description     =   $"conAddIn.dsx":0000
   DisplayName     =   "Trick Advanced Tools"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Command Line / Startup"
   LoadBehavior    =   5
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "conAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' // conAddIn.dsr - connection handler
' // © Krivous Anatoly Anatolevich (The trick), 2016

Option Explicit

Private mcbMenu         As Office.CommandBarControl     ' // Menu item
Private mConfig         As frmAddIn                     ' // Settings window
Private mIsHookInit     As Boolean                      ' // Determines if hijacking has been initialized

Private WithEvents mnuHandler   As CommandBarEvents     ' // Menu event handler
Attribute mnuHandler.VB_VarHelpID = -1
Private WithEvents mProjEvents  As VBProjectsEvents     ' // Project's events
Attribute mProjEvents.VB_VarHelpID = -1

' // Hide form
Sub Hide()
    
    On Error Resume Next
    mConfig.Hide
   
End Sub

' // Show form
Sub Show()
  
    On Error Resume Next
    
    If pVBInstance Is Nothing Then Exit Sub
    If pVBInstance.ActiveVBProject Is Nothing Then Exit Sub
    
    If mConfig Is Nothing Then
        Set mConfig = New frmAddIn
    End If

    mConfig.Show vbModal
     
    UpdateProject pVBInstance.ActiveVBProject
    
End Sub

' // Initialize hooks
Private Function InitializeHooks() As Boolean
    Dim hVba    As Long

    If mIsHookInit Then
        
        InitializeHooks = True
        Exit Function
        
    End If
    
    ' // Get address of the function
    hVba = GetModuleHandle("vba6")
    If hVba = 0 Then Exit Function
    
    pfnTipCompileProject = GetProcAddress(hVba, "TipCompileProject")
    If pfnTipCompileProject = 0 Then Exit Function
    
    ' // Hook
    If Not HookFunction(pfnTipCompileProject, AddressOf TipCompileProject_user) Then
        Exit Function
    End If
    
    pfnTipCompileProjectFull = GetProcAddress(hVba, "TipCompileProjectFull")
    If pfnTipCompileProjectFull = 0 Then Exit Function
    
    ' // Hook
    If Not HookFunction(pfnTipCompileProjectFull, AddressOf TipCompileProjectFull_user) Then
        Exit Function
    End If
    
    pfnTipMakeExe2 = GetProcAddress(hVba, "TipMakeExe2")
    If pfnTipMakeExe2 = 0 Then Exit Function
    
    ' // Hook
    If Not HookFunction(pfnTipMakeExe2, AddressOf TipMakeExe2_user) Then
        Exit Function
    End If
    
    pfnTipFinishExe2 = GetProcAddress(hVba, "TipFinishExe2")
    If pfnTipFinishExe2 = 0 Then Exit Function
    
    ' // Hook
    If Not HookFunction(pfnTipFinishExe2, AddressOf TipFinishExe2_user) Then
        Exit Function
    End If
    
    mIsHookInit = True
    InitializeHooks = True
    
End Function

Private Sub UninitializeHooks()

    If Not mIsHookInit Then Exit Sub
    
    UnhookFunction pfnTipCompileProject
    UnhookFunction pfnTipCompileProjectFull
    UnhookFunction pfnTipMakeExe2
    UnhookFunction pfnTipFinishExe2
    
    mIsHookInit = False
    
End Sub

' // Update checking information for current project
Private Sub UpdateProject( _
            ByVal pVBProject As VBProject)
    Dim lFlags  As Long
    
    On Error GoTo exit_proc
    
    If pVBProject Is Nothing Then Exit Sub
    
    lFlags = pVBInstance.ActiveVBProject.ReadProperty("TAT", "GlobalChecking")
    
    If lFlags And 1 Then
        modCheckingOptions.IntegerOverflowCheck = False
    Else
        modCheckingOptions.IntegerOverflowCheck = True
    End If
    
    If lFlags And 2 Then
        modCheckingOptions.FloatingPointCheck = False
    Else
        modCheckingOptions.FloatingPointCheck = True
    End If
    
    If lFlags And 4 Then
        modCheckingOptions.ArrayBoundsCheck = False
    Else
        modCheckingOptions.ArrayBoundsCheck = True
    End If
    
exit_proc:
    
End Sub

' // Connection is established
Private Sub AddinInstance_OnConnection( _
            ByVal Application As Object, _
            ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, _
            ByVal AddInInst As Object, _
            ByRef custom() As Variant)
            
    On Error GoTo error_handler
    
    Set pVBInstance = Application
    Set mProjEvents = pVBInstance.Events.VBProjectsEvents
    
    If Not InitializeHooks() Then
        Err.Raise ERR_UNABLE_TO_HOOK_FUNCTIONS
    End If
    
    If Not modCheckingOptions.Initialize() Then
        Err.Raise ERR_UNABLE_TO_INITIALIZE_CHECKING
    End If
        
    UpdateProject pVBInstance.ActiveVBProject
    
    Set mcbMenu = AddToAddInCommandBar("Trick Advanced Tools")
    Set mnuHandler = pVBInstance.Events.CommandBarEvents(mcbMenu)

    Exit Sub
    
error_handler:
    
    ErrorLog "Connect::AddinInstance_OnConnection"
    
End Sub

' // Exit add-in
Private Sub AddinInstance_OnDisconnection( _
            ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, _
            ByRef custom() As Variant)
            
    On Error Resume Next
    
    UninitializeHooks
    modCheckingOptions.UnInitialize
    
    mcbMenu.Delete
    Unload mConfig
    Set mConfig = Nothing

End Sub

Private Sub mnuHandler_Click( _
            ByVal CommandBarControl As Object, _
            ByRef handled As Boolean, _
            ByRef CancelDefault As Boolean)
    Me.Show
End Sub

' // Add the item to menu "Add-ins"
Private Function AddToAddInCommandBar( _
                 ByRef sCaption As String) As Office.CommandBarControl
                 
    Dim cbMenuCommandBar As Office.CommandBarControl
    Dim cbMenu           As Object
  
    On Error GoTo error_handler
    
    Set cbMenu = pVBInstance.CommandBars("Add-Ins")
    
    If cbMenu Is Nothing Then
        Exit Function
    End If
    
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    cbMenuCommandBar.Caption = sCaption
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
error_handler:
    
    ErrorLog "Connect::AddToAddInCommandBar"
    
End Function

' // An project has been activated
Private Sub mProjEvents_ItemActivated( _
            ByVal VBProject As VBIDE.VBProject)
    UpdateProject VBProject
End Sub

Private Sub mProjEvents_ItemAdded( _
            ByVal VBProject As VBIDE.VBProject)
    UpdateProject VBProject
End Sub
