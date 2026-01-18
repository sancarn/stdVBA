VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@module
'@description A class which implements a wrapper around a process. Use this class to
'query process information, manipulate them and to forcefully terminate a process.
'@example Simple example of creating a process:
'```vb
'Sub startNewExcel()
'  Dim p As stdProcess: Set p = stdProcess.Create(Application.path & "\EXCEL.EXE /x", vbNormalFocus)
'  Debug.Print p.name
'  Debug.Print p.path
'  Debug.Print p.isRunning
'  Stop
'  p.forceQuit
'  p.waitTilClose
'  Debug.Print p.TimeQuit
'End Sub
'```
'@example Find a process by it's name:
'```vb
'Sub main()
'  Dim p as stdProcess: set p = stdProcess.CreateFromQuery(stdLambda.Create("$1.name = ""notepad.exe"""))
'  Debug.Print p.processId
'End Sub
'```

#if VBA7 then
  Private Declare PtrSafe Function CreateToolhelp32Snapshot Lib "Kernel32.dll" ( _
    ByVal dwFlags As Integer, _
    ByVal th32ProcessID As Integer _
  ) As Long

  Private Declare PtrSafe Function Process32First Lib "Kernel32.dll" ( _
    ByVal hSnapshot As LongPtr, _
    ByRef lppe As PROCESSENTRY32 _
  ) As Boolean
  Private Declare PtrSafe Function Process32Next Lib "Kernel32.dll" ( _
    ByVal hSnapshot As LongPtr, _
    ByRef lppe As PROCESSENTRY32 _
  ) As Boolean

  Private Declare PtrSafe Function CloseHandle Lib "Kernel32.dll" ( _
    ByVal hObject As LongPtr _
  ) As Long

  Private Declare PtrSafe Function QueryFullProcessImageNameA Lib "Kernel32.dll" ( _
    ByVal hProcess As LongPtr, _
    ByVal dwFlags As Long, _
    ByVal lpExeName As String, _
    ByRef lpdwSize As Long _
  ) As Boolean

  Private Declare PtrSafe Function OpenProcess Lib "Kernel32.dll" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Boolean, _
    ByVal dwProcessId As Long _
  ) As LongPtr

  Private Declare PtrSafe Function TerminateProcess Lib "Kernel32.dll" ( _
    ByVal hProcess As LongPtr, _
    ByVal uExitCode As Long _
  ) As Boolean

  Private Declare PtrSafe Function GetProcessVersion Lib "Kernel32.dll" ( _
    ByVal pID As Long _
  ) As Long

  Private Declare PtrSafe Function IsProcessCritical Lib "Kernel32.dll" ( _
    ByVal hProcess As LongPtr, _
    ByRef critical As Boolean _
  ) As Boolean

  Private Declare PtrSafe Function GetProcessTimes Lib "Kernel32.dll" ( _
    ByVal hProcess As LongPtr, _
    ByRef lpCreationTime As FileTime, _
    ByRef lpExitTime As FileTime, _
    ByRef lpKernelTime As FileTime, _
    ByRef lpUserTime As FileTime _
  ) As Boolean

  Private Declare PtrSafe Function GetExitCodeProcess Lib "Kernel32.dll" ( _
    ByVal hProcess As LongPtr, _
    ByRef lpExitCode As Long _
  ) As Boolean

  Private Declare PtrSafe Function GetPriorityClass Lib "Kernel32.dll" ( _
    ByVal hProcess As LongPtr _
  ) As Long

  Private Declare PtrSafe Function FileTimeToSystemTime Lib "Kernel32.dll" ( _
    ByRef lpFileTime As FileTime, _
    ByRef lpSystemTime As SystemTime _
  ) As Boolean
#else
  Private Enum LongPtr
    [_]
  End Enum

  Private Declare Function CreateToolhelp32Snapshot Lib "Kernel32.dll" ( _
    ByVal dwFlags As Integer, _
    ByVal th32ProcessID As Integer _
  ) As Long

  Private Declare Function Process32First Lib "Kernel32.dll" ( _
    ByVal hSnapshot As LongPtr, _
    ByRef lppe As PROCESSENTRY32 _
  ) As Boolean
  Private Declare Function Process32Next Lib "Kernel32.dll" ( _
    ByVal hSnapshot As LongPtr, _
    ByRef lppe As PROCESSENTRY32 _
  ) As Boolean

  Private Declare Function CloseHandle Lib "Kernel32.dll" ( _
    ByVal hObject As LongPtr _
  ) As Long

  Private Declare Function QueryFullProcessImageNameA Lib "Kernel32.dll" ( _
    ByVal hProcess As LongPtr, _
    ByVal dwFlags As Long, _
    ByVal lpExeName As String, _
    ByRef lpdwSize As Long _
  ) As Boolean

  Private Declare Function OpenProcess Lib "Kernel32.dll" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Boolean, _
    ByVal dwProcessId As Long _
  ) As LongPtr

  Private Declare Function TerminateProcess Lib "Kernel32.dll" ( _
    ByVal hProcess As LongPtr, _
    ByVal uExitCode As Long _
  ) As Boolean

  Private Declare Function GetProcessVersion Lib "Kernel32.dll" ( _
    ByVal pID As Long _
  ) As Long

  Private Declare Function IsProcessCritical Lib "Kernel32.dll" ( _
    ByVal hProcess As LongPtr, _
    ByRef critical As Boolean _
  ) As Boolean

  Private Declare Function GetProcessTimes Lib "Kernel32.dll" ( _
    ByVal hProcess As LongPtr, _
    ByRef lpCreationTime As FileTime, _
    ByRef lpExitTime As FileTime, _
    ByRef lpKernelTime As FileTime, _
    ByRef lpUserTime As FileTime _
  ) As Boolean

  Private Declare Function GetExitCodeProcess Lib "Kernel32.dll" ( _
    ByVal hProcess As LongPtr, _
    ByRef lpExitCode As Long _
  ) As Boolean

  Private Declare Function GetPriorityClass Lib "Kernel32.dll" ( _
    ByVal hProcess As LongPtr _
  ) As Long

  Private Declare Function FileTimeToSystemTime Lib "Kernel32.dll" ( _
    ByRef lpFileTime As FileTime, _
    ByRef lpSystemTime As SystemTime _
  ) As Boolean
#end if

'The priority of the process, use #Priority to get one of these values
Public Enum EProcessPriority
  ABOVE_NORMAL_PRIORITY_CLASS = &H8000
  BELOW_NORMAL_PRIORITY_CLASS = &H4000
  HIGH_PRIORITY_CLASS = &H80
  IDLE_PRIORITY_CLASS = &H40
  NORMAL_PRIORITY_CLASS = &H20
  REALTIME_PRIORITY_CLASS = &H100
End Enum

'EProcessAccess is an enum 
'This is used by OpenProcess and ultimately protProcessHandleCreate(...). You generally shouldn't need this enum.
'@protected
Public Enum EProcessAccess
  PROCESS_CREATE_PROCESS = &H80
  PROCESS_CREATE_THREAD = &H2
  PROCESS_DUP_HANDLE = &H40
  PROCESS_QUERY_INFORMATION = &H400
  PROCESS_QUERY_LIMITED_INFORMATION = &H1000
  PROCESS_SET_INFORMATION = &H200
  PROCESS_SET_QUOTA = &H100
  PROCESS_SUSPEND_RESUME = &H800
  PROCESS_TERMINATE = &H1
  PROCESS_VM_OPERATION = &H8
  PROCESS_VM_READ = &H10
  PROCESS_VM_WRITE = &H20
  SYNCHRONIZE = &H100000
End Enum


Private Enum TH32CS
  TH32CS_INHERIT = &H80000000
  TH32CS_SNAPHEAPLIST = &H1
  TH32CS_SNAPMODULE = &H8
  TH32CS_SNAPMODULE32 = &H10
  TH32CS_SNAPPROCESS = &H2
  TH32CS_SNAPTHREAD = &H4
End Enum

Private Type FileTime
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Private Type SystemTime
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Private Const MAX_PATH = 260
Private Type PROCESSENTRY32
  dwSize As Long
  cntUsage As Long
  th32ProcessID As Long
  th32DefaultHeapID As LongPtr
  th32ModuleID As Long
  cntThreads As Long
  th32ParentProcessID As Long
  pcPriClassBase As Long
  dwFlags As Long
  szexeFile As String * MAX_PATH
End Type

Private Type TThis
  ProcessID as Long
  ModuleID as Long
  Name as String
  Path as String
  ProcessHandle as LongPtr
  QueryInfoHandle as LongPtr
End Type
Private This as TThis

'Launches a process and creates a stdProcess object for it
'@constructor
'@param sCMD     - The command to create the process
'@param winStyle - The style to show the window
'@returns        - The launched process
Public Function Create(ByVal sCmd As String, Optional ByVal winStyle As VbAppWinStyle = VbAppWinStyle.vbHide) As stdProcess
  With stdError.getSentry("stdProcess#Create", "sCmd", sCmd, "winStyle", winStyle)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Set Create = New stdProcess
      Call Create.protInitFromProcessId(Shell(sCmd, winStyle))
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Creates a process from a given process id
'@constructor
'@param pID - the process id
'@returns   - the queried process
Public Function CreateFromProcessId(ByVal pID As Long) As stdProcess
  With stdError.getSentry("stdProcess#CreateFromProcessId", "pID", pID)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Set CreateFromProcessId = New stdProcess
      Call CreateFromProcessId.protInitFromProcessId(pID)
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Obtains a the first process which matches the query given
'@constructor
'@param query - The query to search processes for
'@returns     - the queried process
Public Function CreateFromQuery(ByVal query As stdICallable) As stdProcess
  With stdError.getSentry("stdProcess#CreateFromQuery", "query", query)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim pEntry As PROCESSENTRY32: pEntry.dwSize = LenB(pEntry)
      Dim snapshot As LongPtr: snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, ByVal 0&)
      Dim continueSearch As Boolean: continueSearch = Process32First(snapshot, pEntry)
      While continueSearch
        If pEntry.th32ProcessID <> 0 Then
            Dim p As stdProcess: Set p = getProcessFromEntry(pEntry)
            If query.Run(p) Then
              Set CreateFromQuery = p
              Call CloseHandle(snapshot)
              Exit Function
            End If
        End If
        
        'Next process
        continueSearch = Process32Next(snapshot, pEntry)
      Wend
      Call CloseHandle(snapshot)
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Obtains a collection of processes which match the query given
'@constructor
'@param query - The query to search processes for
'@returns Collection<stdProcess> - the queried processes
Public Function CreateManyFromQuery(ByVal query As stdICallable) As Collection
  With stdError.getSentry("stdProcess#CreateManyFromQuery", "query", query)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim pEntry As PROCESSENTRY32: pEntry.dwSize = LenB(pEntry)
      Dim snapshot As LongPtr: snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, ByVal 0&)
      Dim continueSearch As Boolean: continueSearch = Process32First(snapshot, pEntry)
      Dim ret As Collection: Set ret = New Collection
      While continueSearch
        If pEntry.th32ProcessID <> 0 Then
          Dim p As stdProcess: Set p = getProcessFromEntry(pEntry)
          If query.Run(p) Then
            Call ret.add(p)
          End If
        End If
        
        'Next process
        continueSearch = Process32Next(snapshot, pEntry)
      Wend
      Call CloseHandle(snapshot)
      Set CreateManyFromQuery = ret
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Obtains a collection of all processes
'@constructor
'@returns Collection<stdProcess> - All processes
Public Function CreateAll() As Collection
  With stdError.getSentry("stdProcess#CreateAll")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim pEntry As PROCESSENTRY32: pEntry.dwSize = LenB(pEntry)
      Dim snapshot As LongPtr: snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, ByVal 0&)
      Dim continueSearch As Boolean: continueSearch = Process32First(snapshot, pEntry)
      Dim ret As Collection: Set ret = New Collection
      While continueSearch
        'Add proc to return collection
        Call ret.add(getProcessFromEntry(pEntry))
    
        'Next process
        continueSearch = Process32Next(snapshot, pEntry)
      Wend
      Call CloseHandle(snapshot)
      Set CreateAll = ret
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'================================================================================================
'=      PROTECTED CONSTRUCTORS / DESTRUCTORS
'================================================================================================

'Initialises a process object from process id and other data. Optional args are generated if not supplied
'@protected
'@constructor
'@param argID       - The process id
'@param argName     - The process name
'@param argPath     - The process path
'@param argModuleID - The process module id
Friend Sub protInitFromProcessId(ByVal argID As Long, Optional ByVal argName As String = "", Optional ByVal argPath As String = "", Optional ByVal argModuleID As Long = 0)
  With stdError.getSentry("stdProcess#protInitFromProcessId", "argID", argID, "argName", argName, "argPath", argPath, "argModuleID", argModuleID)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      This.ProcessId = argID
      This.Name = argName
      This.Path = argPath
      'If argModuleID = 0 Then This.ModuleId = getModuleIDByPid(argID)
      If Len(argName) = 0 Then This.Name = getProcessNameById(argID)
      If Len(argPath) = 0 Then This.Path = getProcessImageName(argID)
      This.QueryInfoHandle = OpenProcess(PROCESS_QUERY_INFORMATION, False, argID)
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'@destructor
Private Sub Class_Terminate()
  With stdError.getSentry("stdProcess#Class_Terminate")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Call CloseHandle(This.QueryInfoHandle)
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'================================================================================================
'=      PUBLIC PROPERTIES
'================================================================================================

'Obtain the process id
'@returns  - the process id
Public Property Get id() As Long
  With stdError.getSentry("stdProcess#id[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      id = This.ProcessId
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'TODO:
''Obtain the process module id
''@returns {Long} - the process module id
'Public Property Get moduleID() As Long
'  moduleID = This.ModuleId
'End Property

'TODO: SessionID
'TODO: ThreadID

'Obtain the process name
'@returns - the process name
Public Property Get name() As String
  With stdError.getSentry("stdProcess#name[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      name = This.Name
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Obtain the process path
'@returns - the process path
Public Property Get path() As String
  With stdError.getSentry("stdProcess#path[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      path = This.Path
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Returns the Winmgmts object representing the process
'@returns Object<SWbemObjectEx> - The Winmgmts object representing the process
Public Property Get Winmgmt() As Object
  With stdError.getSentry("stdProcess#Winmgmt[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Static cache As Object
      If cache Is Nothing Then Set cache = GetObject("WINMGMTS:\\.\ROOT\cimv2:Win32_Process.Handle=" & This.ProcessId)
      Set Winmgmt = cache
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Attempts to return the command line of the application. Please note [the command line may have changed](https://devblogs.microsoft.com/oldnewthing/20091125-00/?p=15923).
'@returns - The string which originally contained the command line arguments passed to this application.
Public Property Get CommandLine() As String
  With stdError.getSentry("stdProcess#CommandLine[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If Not Winmgmt Is Nothing Then CommandLine = Winmgmt.CommandLine
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Returns whether the process is still running or not
'@returns - `true` if the process is still running, `false` otherwise
Public Property Get isRunning() As Boolean
  With stdError.getSentry("stdProcess#isRunning[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      isRunning = GetProcessVersion(This.ProcessId) > 0
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Returns whether the process is critical or not
'@returns - `true` if the process is critical, `false` otherwise
Public Property Get isCritical() As Boolean
  With stdError.getSentry("stdProcess#isCritical[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      'Note: IsProcessCritical can return a weird boolean where `bool` and `Not bool` both result in `True`, which is nonsense...
      'for this reason we explicitely cast to a long here...
      If CLng(IsProcessCritical(This.QueryInfoHandle, isCritical)) = 0 Then
        Err_Raise 1, "stdProcess.isCritical", "Cannot get critical status of process. Error code 0x" & Hex(Err.LastDllError)
      End If
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Get the process' priority
'@returns - the process' priority
Public Property Get Priority() As EProcessPriority
  With stdError.getSentry("stdProcess#Priority[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Priority = GetPriorityClass(This.QueryInfoHandle)
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property


'Get the datetime representing the time the process was started/created
'@returns - The datetime representing the time the process was started/created
'@TODO: Currently returns in UTC. Convert to standard time.
Public Property Get TimeCreated() As Date
  With stdError.getSentry("stdProcess#TimeCreated[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim t1 As FileTime, t2 As FileTime, t3 As FileTime, t4 As FileTime
      if This.QueryInfoHandle = 0 then exit property
      If GetProcessTimes(This.QueryInfoHandle, t1, t2, t3, t4) Then
        TimeCreated = TimeFromFileTime(t1)
      Else
        Call Err_Raise(1, "stdProcess.TimeCreated", "Cannot get creation time of process. Error code 0x" & Hex(Err.LastDllError))
      End If
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Get the datetime representing the time the process was quit
'@returns - The datetime representing the time the process was quit
'@TODO: Currently returns in UTC. Convert to standard time.
Public Property Get TimeQuit() As Date
  With stdError.getSentry("stdProcess#TimeQuit[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim t1 As FileTime, t2 As FileTime, t3 As FileTime, t4 As FileTime
      if This.QueryInfoHandle = 0 then exit property
      If Not isRunning Then
        If GetProcessTimes(This.QueryInfoHandle, t1, t2, t3, t4) Then
          TimeQuit = TimeFromFileTime(t2)
        Else
          Call Err_Raise(1, "stdProcess.TimeQuit", "Cannot get creation time of process. Error code 0x" & Hex(Err.LastDllError))
        End If
      Else
        Call Err_Raise(1, "stdProcess.TimeQuit", "Cannot obtain the Exit time for a process which is still running.")
      End If
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Get the amount of time that the process has executed in kernel mode
'@returns - The datetime representing the kernel time
'@TODO: Currently returns in UTC. Convert to standard time.
Public Property Get TimeKernel() As Date
  With stdError.getSentry("stdProcess#TimeKernel[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim t1 As FileTime, t2 As FileTime, t3 As FileTime, t4 As FileTime
      if This.QueryInfoHandle = 0 then exit property
      If GetProcessTimes(This.QueryInfoHandle, t1, t2, t3, t4) Then
        TimeKernel = TimeFromFileTime(t3)
      Else
        Call Err_Raise(1, "stdProcess.TimeKernel", "Cannot get creation time of process. Error code 0x" & Hex(Err.LastDllError))
      End If
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Get the amount of time that the process has executed in user mode
'@returns - The datetime representing the user time
'@TODO: Currently returns in UTC. Convert to standard time.
Public Property Get TimeUser() As Date
  With stdError.getSentry("stdProcess#TimeUser[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim t1 As FileTime, t2 As FileTime, t3 As FileTime, t4 As FileTime
      if This.QueryInfoHandle = 0 then exit property
      If GetProcessTimes(This.QueryInfoHandle, t1, t2, t3, t4) Then
        TimeUser = TimeFromFileTime(t4)
      Else
        Call Err_Raise(1, "stdProcess.TimeUser", "Cannot get creation time of process. Error code 0x" & Hex(Err.LastDllError))
      End If
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Get the exit code of this process. Note, an exit code is only ever received if the process has ended. Check isRunning prior to calling this function.
'@returns - the exit code provided at runtime
Public Property Get ExitCode() As Long
  With stdError.getSentry("stdProcess#ExitCode[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      'Note: GetExitCodeProcess can return a weird boolean where `bool` and `Not bool` both result in `True`, which is nonsense...
      'for this reason we explicitely cast to a long here...
      If CLng(GetExitCodeProcess(This.QueryInfoHandle, ExitCode)) = 0 Then
        Err_Raise Err.LastDllError, "stdProcess.ExitCode()", "Cannot get exit code of process. Error code 0x" & Hex(Err.LastDllError)
      End If
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'================================================================================================
'=      PUBLIC METHODS
'================================================================================================

'Wait till the process closes
'@param ExitCode - The exit code to use when terminating the process
Public Sub forceQuit(Optional ByVal ExitCode As Long = 0)
  With stdError.getSentry("stdProcess#forceQuit", "ExitCode", ExitCode)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Call protProcessHandleCreate(PROCESS_TERMINATE)
      If This.ProcessHandle = 0 Then Exit Sub
      
      'Note: TerminateProcess can return a weird boolean where `bool` and `Not bool` both result in `True`, which is nonsense...
      'for this reason we explicitely cast to a long here...
      If CLng(TerminateProcess(This.ProcessHandle, ExitCode)) = 0 Then
        Err_Raise Err.LastDllError, "stdProcess#ForceQuit()", "Cannot terminate process. Error code 0x" & Hex(Err.LastDllError)
      End If
      Call protProcessHandleRelease
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Wait till the process closes
Public Sub waitTilClose()
  With stdError.getSentry("stdProcess#waitTilClose")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      While isRunning
        DoEvents
      Wend
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub







'================================================================================================
'=      PROTECTED METHODS
'================================================================================================

'Get the process handle
'@protected
'@returns - the open process handle
#if VBA7 then
Friend Property Get protProcessHandle() As LongPtr
#else
  Friend Property Get protProcessHandle() As Long
#end if
  With stdError.getSentry("stdProcess#protProcessHandle[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      protProcessHandle = This.ProcessHandle
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Call to OpenProcess() to save a handle in the class. Typically this is not required, thus this method is listed as protected.
'@protected
Friend Sub protProcessHandleCreate(ByVal access As EProcessAccess)
  With stdError.getSentry("stdProcess#protProcessHandleCreate", "access", access)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If This.ProcessHandle = 0 Then
        This.ProcessHandle = OpenProcess(access, 0, This.ProcessId)
      Else
        Err_Raise 1, "", "Process already open."
      End If
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Call to CloseProcess() to close process handle
'@protected
Friend Sub protProcessHandleRelease()
  With stdError.getSentry("stdProcess#protProcessHandleRelease")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Call CloseHandle(This.ProcessHandle)
      This.ProcessHandle = 0
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'================================================================================================
'=      HELPERS
'================================================================================================
'Obtain a stdProcess object from a PROCESSENTRY32 entry
'@param entry - The PROCESSENTRY32 entry to convert to a dictionary.
'@returns - The stdProcess object representing the process
Private Function getProcessFromEntry(entry As PROCESSENTRY32) As stdProcess
  With stdError.getSentry("stdProcess#getProcessFromEntry")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim ret As stdProcess
      Set ret = New stdProcess
      
      With entry
        Dim sName As String: sName = Left(.szexeFile, InStr(1, .szexeFile, vbNullChar) - 1)
        Dim sPath As String: sPath = getProcessImageName(.th32ProcessID)
        Call ret.protInitFromProcessId(.th32ProcessID, sName, sPath, .th32ModuleID)
      End With
      
      Set getProcessFromEntry = ret
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Obtain the Process name from the ProcessID
'@param lProcessID - The process id
'@returns          - The process name
Private Function getProcessNameById(ByVal lProcessID As Long) As String
  With stdError.getSentry("stdProcess#getProcessNameById", "lProcessID", lProcessID)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim entry As PROCESSENTRY32: entry = getCachedProcessEntry(lProcessID)
      getProcessNameById = Left(entry.szexeFile, InStr(1, entry.szexeFile, vbNullChar) - 1)
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Obtain the ModuleID from the ProcessID
'@param lProcessID - The process id
'@returns          - The module id
'@TODO: This method does not currently work as PROCESSENTRY32.th32ModuleID has been deprecated and always returns 0
Private Function getModuleIDByPid(ByVal lProcessID As Long) As Long
  With stdError.getSentry("stdProcess#getModuleIDByPid", "lProcessID", lProcessID)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim entry As PROCESSENTRY32: entry = getCachedProcessEntry(lProcessID)
      getModuleIDByPid = entry.th32ModuleID
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Obtain the PROCESSENTRY32 struct from the ProcessID
'@param lProcessID - The process id
'@param Override   - If `true`, the cache will be ignored and the process will be re-queried, otherwise the cache will be used.
'@returns          - The PROCESSENTRY32 representing the process' data
Private Function getCachedProcessEntry(ByVal lProcessID As Long, Optional ByVal Override As Boolean = False) As PROCESSENTRY32
  With stdError.getSentry("stdProcess#getCachedProcessEntry", "lProcessID", lProcessID, "Override", Override)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Static pEntry As PROCESSENTRY32
      Static stcProcessID As Long
      If stcProcessID <> lProcessID Or Override Then
        stcProcessID = lProcessID
        pEntry.dwSize = LenB(pEntry)
    
        Dim snapshot As LongPtr: snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, ByVal 0&)
        Dim continueSearch As Boolean: continueSearch = Process32First(snapshot, pEntry)
        While continueSearch
          If pEntry.th32ProcessID = lProcessID Then
            continueSearch = False
          Else
            continueSearch = Process32Next(snapshot, pEntry)
          End If
        Wend
        Call CloseHandle(snapshot)
      End If
      getCachedProcessEntry = pEntry
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Obtain the full qualified path of a process from it's ProcessID
'@param lProcessID - The process id
'@returns          - The fully qualified path to the process
Private Function getProcessImageName(ByVal lProcessID As Long) As String
  With stdError.getSentry("stdProcess#getProcessImageName", "lProcessID", lProcessID)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim hProcess As LongPtr
      hProcess = OpenProcess(EProcessAccess.PROCESS_QUERY_INFORMATION, 0, lProcessID)
      If hProcess Then
        Dim sBuf As String
        sBuf = String$(MAX_PATH, Chr$(0))
    
        Dim iLen As Long
        iLen = MAX_PATH
    
        Call QueryFullProcessImageNameA(hProcess, 0, sBuf, iLen)
        sBuf = Left$(sBuf, iLen)
    
        'Close handle
        Call CloseHandle(hProcess)
      Else
        If Err.LastDllError <> 5 Then
          Debug.Print "Error in getProcessImageName of process " & lProcessID & " with code: 0x" & Hex(Err.LastDllError)
        Else
          'Unauthorised - some system processes have this issue.
        End If
      End If
      getProcessImageName = sBuf
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Convert a FileTime struct into an Excel DateTime value.
'@param ft - The filetime to get the Excel date of
'@returns  - The Excel date represented by the FileTime
'@TODO: use GetDynamicTimeZoneInformation() in combination with SystemTimeToTzSpecificLocalTimeEx() to convert to local timezone
Private Function TimeFromFileTime(ByRef ft As FileTime) As Date
  With stdError.getSentry("stdProcess#TimeFromFileTime")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim st As SystemTime
      If FileTimeToSystemTime(ft, st) Then
        With st
          TimeFromFileTime = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
        End With
      Else
        Err_Raise Err.LastDllError, "TimeFromFileTime", "Unhandled error in time conversion. Error number: " & Err.LastDllError
      End If
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function



Private Sub Err_Raise(ByVal number as Long, Optional ByVal source as string = "", Optional ByVal description as string = "")
  Call stdError.Raise(description)
End Sub
