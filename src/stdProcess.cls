VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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


Const MAX_PATH = 260
Const TH32CS_SNAPPROCESS = &H2&

Public Enum EProcessPriority
  ABOVE_NORMAL_PRIORITY_CLASS = &H8000
  BELOW_NORMAL_PRIORITY_CLASS = &H4000
  HIGH_PRIORITY_CLASS = &H80
  IDLE_PRIORITY_CLASS = &H40
  NORMAL_PRIORITY_CLASS = &H20
  REALTIME_PRIORITY_CLASS = &H100
End Enum

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

Private Type FileTime
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

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

Private pProcessId As Long
Private pModuleID As Long
Private pName As String
Private pPath As String

Private pProcessHandle As LongPtr
Private pQueryInfoHandle As LongPtr

'Launches a process and creates a stdProcess object for it
'@constructor
'@param {ByVal String} - The command to create the process
'@param {ByVal vbAppWinStyle} - The style to show the window
'@returns {stdProcess} - The launched process
Public Function Create(ByVal sCmd As String, Optional ByVal winStyle As VbAppWinStyle = VbAppWinStyle.vbHide) As stdProcess
    Set Create = New stdProcess
    Call Create.protInitFromProcessId(Shell(sCmd, winStyle))
End Function

'Creates a process from a given process id
'@constructor
'@param {ByVal Long} - the process id
'@returns {stdProcess} the queried process
Public Function CreateFromProcessId(ByVal pID As Long) As stdProcess
    Set CreateFromProcessId = New stdProcess
    Call CreateFromProcessId.protInitFromProcessId(pID)
End Function

'Obtains a the first process which matches the query given
'@constructor
'@param {stdICallable} - The query to search processes for
'@returns {stdProcess} the queried process
Public Function CreateFromQuery(ByVal query As stdICallable) As stdProcess
    Dim pEntry As PROCESSENTRY32: pEntry.dwSize = LenB(pEntry)
    Dim snapshot As LongPtr: snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, ByVal 0&)
    Dim continueSearch As Boolean: continueSearch = Process32First(snapshot, pEntry)
    While continueSearch
        Dim oDict As Object: Set oDict = getQueryDictFromEntry(pEntry)
        If query.Run(oDict) Then
            With oDict
                Set CreateFromQuery = New stdProcess
                Call CreateFromQuery.protInitFromProcessId(.item("id"), .item("name"), .item("path"), .item("moduleID"))
                Call CloseHandle(snapshot)
                Exit Function
            End With
        Else
            continueSearch = Process32Next(snapshot, pEntry)
        End If
    Wend
    Call CloseHandle(snapshot)
End Function

'Obtains a collection of processes which match the query given
'@constructor
'@param {stdICallable} - The query to search processes for
'@returns {Collection<stdProcess>}
Public Function CreateManyFromQuery(ByVal query As stdICallable) As collection
    Dim pEntry As PROCESSENTRY32: pEntry.dwSize = LenB(pEntry)
    Dim snapshot As LongPtr: snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, ByVal 0&)
    Dim continueSearch As Boolean: continueSearch = Process32First(snapshot, pEntry)
    Dim ret As collection: Set ret = New collection
    While continueSearch
        Dim oDict As Object: Set oDict = getQueryDictFromEntry(pEntry)
        If query.Run(oDict) Then
            With oDict
                Dim proc As stdProcess
                Set proc = New stdProcess
                Call proc.protInitFromProcessId(.item("id"), .item("name"), .item("path"), .item("moduleID"))
                ret.add proc
            End With
        End If

        continueSearch = Process32Next(snapshot, pEntry)
    Wend
    Call CloseHandle(snapshot)
    Set CreateManyFromQuery = ret
End Function

'Obtains a collection of processes which match the query given
'@constructor
'@returns {Collection<stdProcess>} All processes
Public Function CreateAll() As collection
    Dim pEntry As PROCESSENTRY32: pEntry.dwSize = LenB(pEntry)
    Dim snapshot As LongPtr: snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, ByVal 0&)
    Dim continueSearch As Boolean: continueSearch = Process32First(snapshot, pEntry)
    Dim ret As collection: Set ret = New collection
    While continueSearch
        Dim oDict As Object: Set oDict = getQueryDictFromEntry(pEntry)

        With oDict
            Dim proc As stdProcess
            Set proc = New stdProcess
            Call proc.protInitFromProcessId(.item("id"), .item("name"), .item("path"), .item("moduleID"))
            ret.add proc
        End With

        continueSearch = Process32Next(snapshot, pEntry)

    Wend
    Call CloseHandle(snapshot)
    Set CreateAll = ret
End Function

'================================================================================================
'=      PROTECTED CONSTRUCTORS / DESTRUCTORS
'================================================================================================

'Initialises a process object from process id and other data. Optional args are generated if not supplied
'@protected
'@constructor
'@param {ByVal argID} - The process id
'@param {ByVal String = ""} - The process name
'@param {ByVal String = ""} - The process path
'@param {ByVal Long   = 0} - The process module id
'@returns {VOID}
Public Sub protInitFromProcessId(ByVal argID As Long, Optional ByVal argName As String = "", Optional ByVal argPath As String = "", Optional ByVal argModuleID As Long = 0)
    pProcessId = argID
    pModuleID = argModuleID
    pName = argName
    pPath = argPath
    If argModuleID = 0 Then pModuleID = getModuleIDByPid(argID)
    If Len(argName) = 0 Then pName = getProcessNameById(argID)
    If Len(argPath) = 0 Then pPath = getProcessImageName(argID)
    pQueryInfoHandle = OpenProcess(PROCESS_QUERY_INFORMATION, False, argID)
End Sub

Private Sub Class_Terminate()
    Call CloseHandle(pQueryInfoHandle)
End Sub


'================================================================================================
'=      PUBLIC PROPERTIES
'================================================================================================

'Obtain the process id
'@returns {Long} - the process id
Public Property Get id() As Long
    id = pProcessId
End Property

'Obtain the process module id
'@returns {Long} - the process module id
Public Property Get moduleID() As Long
    moduleID = pModuleID
End Property

'TODO: SessionID
'TODO: ThreadID

'Obtain the process name
'@returns {Long} - the process name
Public Property Get name() As String
    name = pName
End Property

'Obtain the process path
'@returns {Long} - the process path
Public Property Get path() As String
    path = pPath
End Property

'Returns the Winmgmts object representing the process
'@returns {Object<SWbemObjectEx>}
Public Property Get Winmgmt() As Object
    Static cache As Object
    If cache Is Nothing Then Set cache = GetObject("WINMGMTS:\\.\ROOT\cimv2:Win32_Process.Handle=" & pProcessId)
    Set Winmgmt = cache
End Property

'Attempts to return the command line of the application. Please note [the command line may have changed](https://devblogs.microsoft.com/oldnewthing/20091125-00/?p=15923).
'@returns {String} The string which originally contained the command line arguments passed to this application.
Public Property Get CommandLine() As String
    If Not Winmgmt Is Nothing Then CommandLine = Winmgmt.CommandLine
End Property

'Returns whether the process is still running or not
'@returns {Boolean} True if the process is still running, false otherwise
Public Property Get isRunning() As Boolean
  isRunning = GetProcessVersion(pProcessId) > 0
End Function

'Returns whether the process is critical or not
'@returns {Boolean} True if the process is critical, false otherwise
Public Property Get isCritical() As Boolean
  'Note: IsProcessCritical can return a weird boolean where `bool` and `Not bool` both result in `True`, which is nonsense...
  'for this reason we explicitely cast to a long here...
  If CLng(IsProcessCritical(pQueryInfoHandle, isCritical)) = 0 Then
    Err.Raise 1, "stdProcess.isCritical", "Cannot get critical status of process. Error code 0x" & Hex(Err.LastDllError)
  End If
End Function

Public Property Get Priority() As EProcessPriority
  Priority = GetPriorityClass(pQueryInfoHandle)
End Property


'Get the datetime representing the time the process was started/created
'@returns {String} The datetime representing the time the process was started/created
Public Property Get TimeCreated() As String
  Dim t1 As FileTime, t2 As FileTime, t3 As FileTime, t4 As FileTime
  If GetProcessTimes(pQueryInfoHandle, t1, t2, t3, t4) Then
    TimeCreated = t1.dwHighDateTime & " " & t1.dwLowDateTime
  Else
    Call Err.Raise(1, "stdProcess.TimeCreated", "Cannot get creation time of process. Error code 0x" & Hex(Err.LastDllError))
  End If
End Property

'Get the datetime representing the time the process was quit
'@returns {String} The datetime representing the time the process was quit
Public Property Get TimeQuit() As String
  Dim t1 As FileTime, t2 As FileTime, t3 As FileTime, t4 As FileTime
  If Not isRunning Then
    If GetProcessTimes(pQueryInfoHandle, t1, t2, t3, t4) Then
      TimeQuit = t2.dwHighDateTime & " " & t2.dwLowDateTime
    Else
      Call Err.Raise(1, "stdProcess.TimeQuit", "Cannot get creation time of process. Error code 0x" & Hex(Err.LastDllError))
    End If
  Else
    Call Err.Raise(1, "stdProcess.TimeQuit", "Cannot obtain the Exit time for a process which is still running.")
  End If
End Property

'Get the amount of time that the process has executed in kernel mode
'@returns {String} The datetime representing the kernel time
Public Property Get TimeKernel() As String
  Dim t1 As FileTime, t2 As FileTime, t3 As FileTime, t4 As FileTime
  If GetProcessTimes(pQueryInfoHandle, t1, t2, t3, t4) Then
    TimeKernel = t3.dwHighDateTime & " " & t3.dwLowDateTime
  Else
    Call Err.Raise(1, "stdProcess.TimeKernel", "Cannot get creation time of process. Error code 0x" & Hex(Err.LastDllError))
  End If
End Property

'Get the amount of time that the process has executed in user mode
'@returns {String} The datetime representing the user time
Public Property Get TimeUser() As String
  Dim t1 As FileTime, t2 As FileTime, t3 As FileTime, t4 As FileTime
  If GetProcessTimes(pQueryInfoHandle, t1, t2, t3, t4) Then
    TimeUser = t4.dwHighDateTime & " " & t4.dwLowDateTime
  Else
    Call Err.Raise(1, "stdProcess.TimeUser", "Cannot get creation time of process. Error code 0x" & Hex(Err.LastDllError))
  End If
End Property


'@returns {Long} -1 if cannot get ExitCode, otherwise the exit code provided at runtime
Public Property Get ExitCode() As Long
  'Note: GetExitCodeProcess can return a weird boolean where `bool` and `Not bool` both result in `True`, which is nonsense...
  'for this reason we explicitely cast to a long here...
  If CLng(GetExitCodeProcess(pQueryInfoHandle, ExitCode)) = 0 Then
    Err.Raise Err.LastDllError, "stdProcess.ExitCode()", "Cannot get exit code of process. Error code 0x" & Hex(Err.LastDllError)
  End If
End Property

'================================================================================================
'=      PUBLIC METHODS
'================================================================================================

'Wait till the process closes
'@returns {VOID}
Public Sub forceQuit(Optional ByVal ExitCode As Long = 0)
  Call protProcessHandleCreate(PROCESS_TERMINATE)
  If pProcessHandle = 0 Then Exit Sub
  
  'Note: TerminateProcess can return a weird boolean where `bool` and `Not bool` both result in `True`, which is nonsense...
  'for this reason we explicitely cast to a long here...
  If CLng(TerminateProcess(pProcessHandle, ExitCode)) = 0 Then
    Err.Raise Err.LastDllError, "stdProcess#ForceQuit()", "Cannot terminate process. Error code 0x" & Hex(Err.LastDllError)
  End If
  Call protProcessHandleRelease
End Sub

'Wait till the process closes
'@returns {VOID}
Public Sub waitTilClose()
  While isRunning
    DoEvents
  Wend
End Sub







'================================================================================================
'=      PROTECTED METHODS
'================================================================================================

'Get the process handle
'@protected
'@returns {LongPtr} the open process handle
Public Property Get protProcessHandle() As LongPtr
  protProcessHandle = pProcessHandle
End Property

'Call to OpenProcess() to save a handle in the class. Typically this is not required, thus this method is listed as protected.
'@protected
'@returns {VOID}
Public Sub protProcessHandleCreate(ByVal access As EProcessAccess)
  If pProcessHandle = 0 Then
    pProcessHandle = OpenProcess(access, 0, pProcessId)
  Else
    Err.Raise 1, "", "Process already open."
  End If
End Sub

'Call to CloseProcess() to close process handle
'@protected
'@returns {VOID}
Public Sub protProcessHandleRelease()
  Call CloseHandle(pProcessHandle)
  pProcessHandle = 0
End Sub

'================================================================================================
'=      HELPERS
'================================================================================================
'Obtain a Dictionary object used for Querying.
'@param {PROCESSENTRY32} - The PROCESSENTRY32 entry to convert to a dictionary.
'@returns {Dictionary<string: variant>} - Dictionary of process information
Private Function getQueryDictFromEntry(entry As PROCESSENTRY32) As Object
    Dim ret As Object: Set ret = CreateObject("Scripting.Dictionary")
    ret("id") = entry.th32ProcessID
    ret("parentID") = entry.th32ParentProcessID
    ret("moduleID") = entry.th32ModuleID
    ret("name") = Left(entry.szexeFile, InStr(1, entry.szexeFile, vbNullChar) - 1)
    ret("path") = getProcessImageName(entry.th32ProcessID)
    Set getQueryDictFromEntry = ret
End Function

'Obtain the Process name from the ProcessID
'@param {ByVal Long} - The process id
'@returns {Long} - The process name
Private Function getProcessNameById(ByVal lProcessID As Long) As String
    Dim entry As PROCESSENTRY32: entry = getCachedProcessEntry(lProcessID)
    getProcessNameById = Left(entry.szexeFile, InStr(1, entry.szexeFile, vbNullChar) - 1)
End Function

'Obtain the ModuleID from the ProcessID
'@param {ByVal Long} - The process id
'@returns {Long} - The module id
Private Function getModuleIDByPid(ByVal lProcessID As Long) As Long
    Dim entry As PROCESSENTRY32: entry = getCachedProcessEntry(lProcessID)
    getModuleIDByPid = entry.th32ModuleID
End Function

'Obtain the PROCESSENTRY32 struct from the ProcessID
'@param {ByVal Long} - The process id
'@returns {PROCESSENTRY32} - The PROCESSENTRY32 representing the process' data
Private Function getCachedProcessEntry(ByVal lProcessID As Long, Optional ByVal Override As Boolean = False) As PROCESSENTRY32
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
End Function

'Obtain the full qualified path of a process from it's ProcessID
'@param {ByVal Long} - The process id
'@returns {String} - The fully qualified path to the process
Public Function getProcessImageName(ByVal lProcessID As Long) As String
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
End Function

