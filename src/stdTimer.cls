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
'  * Use of VBOM and simplified registry setting
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
'    * Ensures VBOM is enabled
'    * Creates a `new Application` object and adds a workbook to this new application.
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

'Create an instance of the timer
'@param iMilliseconds - Number of milliseconds to raise Tick Event between
'@param callback as stdICallable<()=>void> - Callback object. Must contain a Run() method.
'@returns - The frequency isn't guaranteed but it will be as close as humanly possible.
Public Function Create(ByVal iMilliseconds As Long, Optional ByVal callback As Object = Nothing) As stdTimer
  Set Create = New stdTimer
  Call Create.protInit(iMilliseconds, callback)
End Function

'Initialise module
'@protected
'@param iTimerResolution - Number of milliseconds to raise Tick Event between
'@param callback as stdICallable<()=>void> -  Optional Listener
Public Sub protInit(ByVal iTimerResolution As Long, callback As Object)
  Set pCallback = callback
  If Not getVBOM Then Call setVBOM(True)
  Set xlRemoteWorkbook = createRemoteWorkbook()
  If Not getVBOM Then Call setVBOM(False)
  Call injectCode(xlRemoteWorkbook, "TimerMain", getTimerCode(iTimerResolution))
  Call xlRemoteWorkbook.Application.Run("TimerMain.StartTimer")
End Sub

'Create and return a remote workbook, close equivalent of generating a new thread, however these threads will be in sync (due to usage of COM events)
'@returns - Remote Workbook object.
Private Function createRemoteWorkbook(Optional ByVal withWorkbook As Boolean = True) As Workbook
  Dim app As Application: Set app = New Application
  With app
    #If Not DEBUG_MODE Then
      .Visible = False
      .ScreenUpdating = False
    #Else
      .Visible = True
    #End If
    .PrintCommunication = False
    .DisplayAlerts = False
  End With
  Set createRemoteWorkbook = app.Workbooks.Add
End Function

'Inject VBA code into a workbook
'@param wb          - The workbook to inject code into
'@param sModuleName - The module name to create
'@param sCode       - The code to insert into the module
'@param iModuleType - The type of code module to create (Module=1, Class=2)
Public Sub injectCode(ByVal wb As Workbook, ByVal sModuleName As String, ByRef sCode As String, Optional ByVal iModuleType As Long = 1)
  With wb.VBProject.VBComponents.Add(iModuleType).CodeModule
    If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
    .AddFromString sCode
    .Parent.Name = sModuleName
  End With
End Sub

'Check if VB Object model is extensible
'@returns - `True` if VBE Is Extensible (i.e. checked within the IDE, rather than the registry)
Private Function getVBOM() As Boolean
  On Error Resume Next
  getVBOM = Not Application.VBE.ActiveVBProject Is Nothing
End Function

'Check whether a registry key exists
'@param newValue - `True` to enable VBOM, False otherwise
'@returns - True if key exists, False otherwise
'@author [Cristian Buse](https://github.com/cristianbuse)
Private Function setVBOM(ByVal newValue As Boolean) As Boolean
  Dim i As Long: i = IIf(newValue, 1, 0)
  #If Mac Then
    Shell "defaults write com.microsoft.Excel AccessVBOM -int " & i
    setVBOM = True
  #Else
    Dim rKey As String
    rKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & Application.Version _
          & "\" & Replace(Application.Name, "Microsoft ", vbNullString) _
          & "\Security\AccessVBOM"
    On Error Resume Next
    CreateObject("WScript.Shell").RegWrite rKey, i, "REG_DWORD"
    setVBOM = (Err.Number = 0)
    On Error GoTo 0
  #End If
End Function

'Obtain code for a simple timer
'@param iTimerResolution - Number of milliseconds to raise Tick Event between
'@returns - VBA code to insert into a module
Private Function getTimerCode(ByVal iTimerResolution As Long) As String
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
  s = s & vbCrLf & "Private bRunning as boolean"
  s = s & vbCrLf & "Sub MainLoop()"
  s = s & vbCrLf & "  While bRunning"
  s = s & vbCrLf & "    Sheet1.Range(""A1"").Value = ""Tick"""
  s = s & vbCrLf & "    sleep " & iTimerResolution
  s = s & vbCrLf & "    DoEvents"
  s = s & vbCrLf & "  Wend"
  s = s & vbCrLf & "  ThisWorkbook.Close false" 'cleanup
  s = s & vbCrLf & "End Sub"
  s = s & vbCrLf & "Sub StartTimer()"
  s = s & vbCrLf & "  bRunning = true"
  s = s & vbCrLf & "  Application.OnTime now(), ""MainLoop""" 'initialise main loop asynchronously
  s = s & vbCrLf & "End Sub"
  s = s & vbCrLf & "Sub StopTimer()"
  s = s & vbCrLf & "  bRunning = false"
  s = s & vbCrLf & "End Sub"
  getTimerCode = s
End Function

'On terminate, stop timer. This will cause workbook to close
Private Sub Class_Terminate()
  Call xlRemoteWorkbook.Application.Run("TimerMain.StopTimer")
  Set xlRemoteWorkbook = Nothing
End Sub

'SheetChange event hook
'@param Sh     - The sheet that changed
'@param Target - The range that changed
Private Sub xlRemoteWorkbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
  RaiseEvent Tick
  If Not pCallback Is Nothing Then pCallback.Run
End Sub
