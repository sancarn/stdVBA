VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Module greatly inspired by Cristian Buse's SafeTimers:
'  https://github.com/cristianbuse/Excel-VBA-SafeTimers
'
'Inspirations used:
'  * Use of remote workbook
'  * Use of VBOM (requires user-enabled trust)
'
'Comparrison of SaveTimers-Native and stdTimer:
'  * stdTimer will not raise the Tick event while displaying Msgbox or InputBox, however SafeTimers will
'  * stdTimer will raise events while Edit mode is "Edit" or "Point", SafeTimers-Native do not.
'  * stdTimer uses COM events (specifically WorksheetChange), SafeTimers-Native uses direct COM calls as a communication mechanism
'  * stdTimer can be bound to stdICallable callback OR via event interop. SafeTimers-Native can only be hooked into with Module macro names
'
'Limitations:
'  * Neither set of timers raise events while debugging code.
'  * Neither set of timers will be called exactly when it needs to. There will be a little bit of error which may have to be corrected.
'
'Architecture:
'  When a timer is initialised it:
'    * Requires existing VBOM access; never changes trust settings
'    * Creates an owned isolated Excel application and adds a workbook to this new application.
'    * Adds Timer code to this workbook and schedules this code to run with `xlApp.OnTime Now(), "MainLoop"`
'  Every time the timer loops
'    * A change to the remote workbook's worksheet is made
'    * In the original workbook we register to receive these events, and receive a `xlRemoteWorkbook_SheetChange` event.
'    * Here we raise the event, and call the callback if supplied.
'  When this class is terminated
'    * It will also tell the remote workbook to close.


'Set DEBUG_MODE to true if you want to see the remote workbook
#Const DEBUG_MODE = False

'Called every time the timer loops
Public Event Tick()

Private WithEvents xlRemoteWorkbook As Workbook
Attribute xlRemoteWorkbook.VB_VarHelpID = -1
Private pCallback As Object 'stdICallable
Private xlOwnedApplication As Application
Private pStopped As Boolean

'Create an instance of the timer
'@constructor
'@param iMilliseconds - Number of milliseconds to raise Tick Event between
'@param callback as stdICallable<()=>void> - Callback object. Must contain a Run() method.
'@remarks - Requires user-enabled AccessVBOM; never modifies trust settings.
'@returns - The frequency isn't guaranteed but it will be as close as humanly possible.
Public Function Create(ByVal iMilliseconds As Long, Optional ByVal callback As Object = Nothing) As stdTimer
  With stdError.getSentry("stdTimer#Create", "iMilliseconds", iMilliseconds, "callback", callback)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Set Create = New stdTimer
      Call Create.protInit(iMilliseconds, callback)
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Initialise module
'@protected
'@param iTimerResolution - Number of milliseconds to raise Tick Event between
'@param callback as stdICallable<()=>void> -  Optional Listener
Public Sub protInit(ByVal iTimerResolution As Long, callback As Object)
  With stdError.getSentry("stdTimer#protInit", "iTimerResolution", iTimerResolution, "callback", callback)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If Not xlOwnedApplication Is Nothing Or pStopped Then Err_Raise 5, "stdTimer", "Timer already initialised or stopped"
      If iTimerResolution < 1 Then Err_Raise 5, "stdTimer", "Timer resolution must be positive"
      On Error GoTo Failed
      Set pCallback = callback
      Set xlRemoteWorkbook = createRemoteWorkbook()
      Call injectCode(xlRemoteWorkbook, "TimerMain", getTimerCode(iTimerResolution))
      Call xlRemoteWorkbook.Application.Run("TimerMain.StartTimer")
      Exit Sub
    Failed:
      Dim errorNumber As Long, errorSource As String, errorDescription As String
      errorNumber = Err.Number: errorSource = Err.Source: errorDescription = Err.Description
      On Error Resume Next
      If Not xlRemoteWorkbook Is Nothing Then xlRemoteWorkbook.Close False
      If Not xlOwnedApplication Is Nothing Then
        If xlOwnedApplication.Workbooks.Count = 0 Then xlOwnedApplication.Quit
      End If
      Set xlRemoteWorkbook = Nothing
      Set xlOwnedApplication = Nothing
      Set pCallback = Nothing
      pStopped = True
      On Error GoTo stdErrorWrapper_ErrorOccurred
      Err_Raise errorNumber, errorSource, errorDescription
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Create and return a remote workbook, close equivalent of generating a new thread, however these threads will be in sync (due to usage of COM events)
'@returns - Remote Workbook object.
Private Function createRemoteWorkbook(Optional ByVal withWorkbook As Boolean = True) As Workbook
  With stdError.getSentry("stdTimer#createRemoteWorkbook", "withWorkbook", withWorkbook)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      'Do not create app with `new Application`. If this is used VBOM Extensibility will revert upon creation.
      Set xlOwnedApplication = CreateObject("Excel.Application")
      With xlOwnedApplication
        #If Not DEBUG_MODE Then
          .Visible = False
          .ScreenUpdating = False
        #Else
          .Visible = True
        #End If
        .PrintCommunication = False
        .DisplayAlerts = False
      End With
      Set createRemoteWorkbook = xlOwnedApplication.Workbooks.Add
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Inject VBA code into a workbook
'@param wb          - The workbook to inject code into
'@param sModuleName - The module name to create
'@param sCode       - The code to insert into the module
'@param iModuleType - The type of code module to create (Module=1, Class=2)
Public Sub injectCode(ByVal wb As Workbook, ByVal sModuleName As String, ByRef sCode As String, Optional ByVal iModuleType As Long = 1)
  With stdError.getSentry("stdTimer#injectCode", "wb", wb, "sModuleName", sModuleName, "sCode", sCode, "iModuleType", iModuleType)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      With wb.VBProject.VBComponents.Add(iModuleType).CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .AddFromString sCode
        .Parent.Name = sModuleName
      End With
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Obtain code for a simple timer
'@param iTimerResolution - Number of milliseconds to raise Tick Event between
'@returns - VBA code to insert into a module
Private Function getTimerCode(ByVal iTimerResolution As Long) As String
  With stdError.getSentry("stdTimer#getTimerCode", "iTimerResolution", iTimerResolution)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim s As String: s = "'Code generated by stdTimer"
      s = s & vbCrLf & "#If Mac Then"
      s = s & vbCrLf & "  #If VBA7 Then"
      s = s & vbCrLf & "    Private Declare PtrSafe Sub USleep Lib ""/usr/lib/libc.dylib"" Alias ""usleep"" (ByVal dwMicroseconds As Long)"
      s = s & vbCrLf & "  #Else"
      s = s & vbCrLf & "    Private Declare Sub USleep Lib ""/usr/lib/libc.dylib"" Alias ""usleep"" (ByVal dwMicroseconds As Long)"
      s = s & vbCrLf & "  #End If"
      s = s & vbCrLf & "#Else 'Windows"
      s = s & vbCrLf & "  #If VBA7 Then"
      s = s & vbCrLf & "    Public Declare PtrSafe Sub Sleep Lib ""kernel32"" (ByVal dwMilliseconds As Long)"
      s = s & vbCrLf & "  #Else"
      s = s & vbCrLf & "    Public Declare  Sub Sleep Lib ""kernel32"" Alias ""Sleep"" (ByVal dwMilliseconds As Long)"
      s = s & vbCrLf & "  #End If"
      s = s & vbCrLf & "#End If"
      s = s & vbCrLf & "Private bRunning As Boolean"
      s = s & vbCrLf & "Private bInLoop As Boolean"
      s = s & vbCrLf & "Private nextTick As Date"
      s = s & vbCrLf & "Private timerMacro As String"
      s = s & vbCrLf & "Sub MainLoop()"
      s = s & vbCrLf & "  bInLoop = True"
      s = s & vbCrLf & "  On Error GoTo Finished"
      s = s & vbCrLf & "  While bRunning"
      s = s & vbCrLf & "    ThisWorkbook.Sheets(1).Range(""A1"").Value = ""Tick"""
      s = s & vbCrLf & "#If Mac Then"
      s = s & vbCrLf & "    USleep " & CDbl(iTimerResolution) * 1000
      s = s & vbCrLf & "#Else"
      s = s & vbCrLf & "    Sleep " & iTimerResolution
      s = s & vbCrLf & "#End If"
      s = s & vbCrLf & "    DoEvents"
      s = s & vbCrLf & "  Wend"
      s = s & vbCrLf & "Finished:"
      s = s & vbCrLf & "  bRunning = False"
      s = s & vbCrLf & "  bInLoop = False"
      s = s & vbCrLf & "  Application.OnTime Now(), ""'"" & ThisWorkbook.Name & ""'!TimerMain.Cleanup"""
      s = s & vbCrLf & "
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub"
  s = s & vbCrLf & "Sub StartTimer()"
  s = s & vbCrLf & "  bRunning = True"
  s = s & vbCrLf & "  nextTick = Now()"
  s = s & vbCrLf & "  timerMacro = ""'"" & ThisWorkbook.Name & ""'!TimerMain.MainLoop"""
  s = s & vbCrLf & "  Application.OnTime nextTick, timerMacro"
  s = s & vbCrLf & "End Sub"
  s = s & vbCrLf & "Sub StopTimer()"
  s = s & vbCrLf & "  bRunning = False"
  s = s & vbCrLf & "  If bInLoop Then Exit Sub"
  s = s & vbCrLf & "  On Error Resume Next"
  s = s & vbCrLf & "  Application.OnTime nextTick, timerMacro, , False"
  s = s & vbCrLf & "  Application.OnTime Now(), ""'"" & ThisWorkbook.Name & ""'!TimerMain.Cleanup"""
  s = s & vbCrLf & "End Sub"
  s = s & vbCrLf & "Public Sub Cleanup()"
  s = s & vbCrLf & "  On Error Resume Next"
  s = s & vbCrLf & "  If Application.Workbooks.Count = 1 Then"
  s = s & vbCrLf & "    ThisWorkbook.Saved = True"
  s = s & vbCrLf & "    Application.Quit"
  s = s & vbCrLf & "  Else"
  s = s & vbCrLf & "    ThisWorkbook.Close False"
  s = s & vbCrLf & "  End If"
  s = s & vbCrLf & "End Sub"
  getTimerCode = s
End Function

'Stop and release the timer. Safe to call repeatedly or from Tick.
'@remarks - Active remote loops unwind before closing their workbook and host.
Public Sub StopTimer()
  With stdError.getSentry("stdTimer#StopTimer")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If pStopped Then Exit Sub
      pStopped = True
      On Error Resume Next
      If Not xlOwnedApplication Is Nothing Then xlOwnedApplication.Run "TimerMain.StopTimer"
      Set xlRemoteWorkbook = Nothing
      Set xlOwnedApplication = Nothing
      Set pCallback = Nothing
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

Private Sub Class_Terminate()
  With stdError.getSentry("stdTimer#Class_Terminate")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      StopTimer
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'SheetChange event hook
'@param Sh     - The sheet that changed
'@param Target - The range that changed
Private Sub xlRemoteWorkbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
  With stdError.getSentry("stdTimer#xlRemoteWorkbook_SheetChange", "Sh", Sh, "Target", Target)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If pStopped Then Exit Sub
      Dim keepAlive As stdTimer: Set keepAlive = Me
      RaiseEvent Tick
      If Not pStopped And Not pCallback Is Nothing Then pCallback.Run
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub



Private Sub Err_Raise(ByVal number as Long, Optional ByVal source as string = "", Optional ByVal description as string = "")
  Call stdError.Raise(description)
End Sub
