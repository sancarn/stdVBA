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
'  * stdTimer will raise events while Edit mode is "Edit" or "Point" (i.e. while editing formulae), SafeTimers-Native do not.
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
'    * Adds Timer code to this workbook
'    * New timers are added to this remote instance via the `AddTimer` method which also returns a GUID (identifier for timer).
'    * After the first timer is added we schedule the main code to run with `xlApp.OnTime Now(), "MainLoop"`
'  Every time the timer loops
'    * Each added timer is checked whether it has been exceeded, if it does exceed:
'      * ID is written to cell A1 (i.e. A change to the remote workbook)
'      * In the original workbook we register to receive these events, and receive a `xlRemoteWorkbook_SheetChange` event.
'      * If ID matches ID which caused change, we raise the event, and call the callback if supplied.
'  When this class is terminated
'    * It will also tell the remote workbook to close.
'  If state is lossed:
'    * State can be recovered using timer ID and a pointer saved in the registry pointing to the remote workbook.
'    * If state is not re-initialised within 5 minutes minimum (timer frequency maximum), the timer will be deactivated.
'    * If all timers are deactivated, and no new timers are added the remote workbook is closed, and the remote application quit.
'    * Note: To recover state automatically you can use a state loss detector, like https://github.com/cristianbuse/VBA-StateLossCallback 

'TODO: Max should be 2 ^ 27 as this will enable 24 hour timers at least.
'TODO: Maybe use VBA.Timer() instead of Sleep(10) as this will be more accurate, even if it isn't super accurate. Mind that
'      this returns the number of seconds elapsed since midnight using a single-precision floating point value. Multiplying by 
'      1000 will give ms since midnight. So mind that this may reset to 0.
'      Ultimately, (OldTimer-Timer())*1000
 

#If Mac Then
    #If VBA7 Then
        Private Declare PtrSafe Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As LongPtr) As LongPtr
    #Else
        Private Declare Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As Long) As Long
    #End If
#Else 'Windows
    #If VBA7 Then
        Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    #Else
        Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    #End If
#End If
#If VBA7 = 0 Then
    Private Enum LongPtr
        [_]
    End Enum
#End If
#If Win64 Then
    Const PTR_SIZE As Long = 8
#Else
    Const PTR_SIZE As Long = 4
#End If

'Set DEBUG_MODE to true if you want to see the remote workbook
#Const DEBUG_MODE = False

'Called every time the timer loops
Public Event Tick()

Private WithEvents xlRemoteWorkbook As Workbook
Private pCallback As Object 'stdICallable
Private pID as string

'Create an instance of the timer
'@param {Long} Number of milliseconds to raise Tick Event between
'@param {stdICallable=Nothing} Callback object. Must contain a Run() method.
'@returns {stdTimer} The frequency isn't guaranteed but it will be as close as humanly possible.
Public Function Create(ByVal iMilliseconds As Long, Optional ByVal callback As Object = Nothing) As stdTimer
  Set Create = New stdTimer
  Call Create.protInit(iMilliseconds, callback)
End Function

'Create an instance of the timer
'@param {Long} Number of milliseconds to raise Tick Event between
'@param {String} ID of existing callback
'@param {stdICallable=Nothing} Callback object. Must contain a Run() method.
'@returns {stdTimer} The frequency isn't guaranteed but it will be as close as humanly possible.
Public Function CreateFromExisting(ByVal iMilliseconds as long, ByVal sExistingID as string, optional byval callback as object = nothing) as stdTimer
  set Create = new stdTimer
  Call Create.protInit(iMilliseconds, callback, sExistingID)
End Function

'Initialise module
'@param {Long} Number of milliseconds to raise Tick Event between
'@param {stdICallable} Optional Listener
Public Sub protInit(ByVal iTimerResolution As Long, callback As Object, optional byval sID as string = vbNullString)
  set xlRemoteWorkbook = stdTimer.getRemoteWorkbookInstance()
  Set pCallback = callback
  if sID = vbNullString then
    pID = xlRemoteWorkbook.Application.Run("TimerMain.AddTimer", iTimerResolution)
  else
    pID = sID
  end if
End Sub

'Obtain application instance
'@returns {Application} Returns a remote instance of application
Public Function getRemoteWorkbookInstance() as Workbook
  static wb as Workbook
  if wb is nothing then
    'check whether instance is still active
    Dim sExistingInstance as string: sExistingInstance = GetSetting("stdTimer","instance")
    if sExistingInstance = "0" then
      If Not getVBOM Then Call setVBOM(True)
      Set wb = createRemoteWorkbook()
      If Not getVBOM Then Call setVBOM(False)
      Call injectCode(wb, "TimerMain", getTimerCode(iTimerResolution))
      Call wb.Application.Run("TimerMain.StartTimer")
      Call SaveSetting("stdTimer","instance", CStr(ObjPtr(wb)))
    else
      set wb = derefObject(CCLngPtr(sExistingInstance))
    end if
  end if
  set getRemoteWorkbookInstance = wb
End Function

Public Property Get ID as string
  if stdTimer is me then
    Err.Raise 1, "stdTimer", "No timer for predefined instance"
  else
    ID = pID
  end if
End Property

'Dereference an object from pointer
Private Function derefObject(ByVal ptr as longptr) as  Object
  Dim result As Object
  CopyMemory result, ByVal varPtr(memAddress), LenB(memAddress)
  Set derefObject = result                             'Set and add ref
  'Zero memory
  Dim b() as Byte, Redim b(lenb(memAddress))
  CopyMemory result, ByVal varPtr(b(0)), LenB(memAddress)
End Function

'Obtain a pointer from string
'@param {String} Pointer as string
'@returns {LongPtr} Casted longptr
Private Function CCLngPtr(ByVal ptr as string) as LongPtr
  #if VBA7 then
    CCLngPtr = CCLngPtr(ptr)
  #else
    CCLngPtr = CLng(ptr)
  #end if
End Function

'Create and return a remote workbook, close equivalent of generating a new thread, however these threads will be in sync (due to usage of COM events)
'@returns {Workbook} Remote Workbook object.
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
'@param {Workbook} The workbook to inject code into
'@param {string} The module name to create
'@param {ByRef string} The code to insert into the module
'@param {Long} The type of code module to create (Module=1, Class=2)
Public Sub injectCode(ByVal wb As Workbook, ByVal sModuleName As String, ByRef sCode As String, Optional ByVal iModuleType As Long = 1)
  With wb.VBProject.VBComponents.Add(iModuleType).CodeModule
    If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
    .AddFromString sCode
    .Parent.Name = sModuleName
  End With
End Sub

'Check if VB Object model is extensible
'@returns {Boolean} True if VBE Is Extensible (i.e. checked within the IDE, rather than the registry)
Private Function getVBOM() As Boolean
  On Error Resume Next
  getVBOM = Not Application.VBE.ActiveVBProject Is Nothing
End Function

'Check whether a registry key exists
'@param {Boolean} True to enable VBOM, False otherwise
'@returns {Boolean} True if key exists, False otherwise
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
  s = s & vbCrLf & "    Sheets(1).Range(""A1"").Value = ""Tick""" 'Sheets(1) due to issue #89
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
Private Sub xlRemoteWorkbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
  if pID = target.value then
    RaiseEvent Tick
    If Not pCallback Is Nothing Then pCallback.Run
    Call SaveSetting("stdTimer", "last_" & pID, CStr(Now()))
  end if
End Sub
