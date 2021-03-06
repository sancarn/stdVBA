VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare PtrSafe Function CreateToolhelp32Snapshot Lib "kernel32.dll" ( _
    ByVal dwFlags As Integer, _
    ByVal th32ProcessID As Integer _
) As Long
 
Private Declare PtrSafe Function Process32First Lib "kernel32.dll" ( _
    ByVal hSnapshot As LongPtr, _
    ByRef lppe As PROCESSENTRY32 _
) As Boolean
Private Declare PtrSafe Function Process32Next Lib "kernel32.dll" ( _
    ByVal hSnapshot As LongPtr, _
    ByRef lppe As PROCESSENTRY32 _
) As Boolean

Private Declare PtrSafe Function CloseHandle Lib "kernel32.dll" ( _ 
    ByVal hObject As LongPtr _ 
) As Long

Private Declare PtrSafe Function QueryFullProcessImageNameA Lib "Kernel32.dll" ( _ 
    ByVal hProcess As LongPtr, _ 
    ByVal dwFlags As Long, _ 
    ByVal lpExeName As String, _ 
    ByRef lpdwSize As Long _ 
) As Boolean

Private Declare PtrSafe Function OpenProcess Lib "Kernel32.dll" ( _ 
  ByVal dwDesiredAccess as Long, _ 
  ByVal bInheritHandle as Boolean, _ 
  ByVal dwProcessId as Long _ 
) as LongPtr

Private Declare PtrSafe Function TerminateProcess Lib "Kernel32.dll" ( _ 
  ByVal hProcess as LongPtr, _ 
  ByVal uExitCode as Long, _  
) as Boolean

Private Declare PtrSafe Function GetProcessVersion Lib "Kernel32.dll" ( _ 
  ByVal pid as long _ 
) as Long

Private Declare PtrSafe Function IsProcessCritical Lib "Kernel32.dll" ( _ 
  ByVal hProcess as LongPtr, _ 
  ByRef critical as Boolean _ 
) as Boolean

Private Declare PtrSafe Function GetProcessTimes Lib "Kernel32.dll" ( _ 
  ByVal hProcess as LongPtr, _ 
  ByRef lpCreationTime as FileTime, _ 
  ByRef lpExitTime as FileTime, _ 
  ByRef lpKernelTime as FileTime, _ 
  ByRef lpUserTime as FileTime _ 
) as Boolean

Private Declare PtrSafe Function GetExitCodeProcess Lib "Kernel32.dll" ( _ 
  ByVal hProcess as LongPtr, _ 
  ByRef lpExitCode as Long _ 
) as Boolean

Private Declare PtrSafe Function GetPriorityClass Lib "Kernel32.dll" ( _ 
  ByVal hProcess as LongPtr _ 
) as Long


Const MAX_PATH = 260
Const TH32CS_SNAPPROCESS = &H2&

Public Enum EProcessPriority
  ABOVE_NORMAL_PRIORITY_CLASS = &H00008000
  BELOW_NORMAL_PRIORITY_CLASS = &H00004000
  HIGH_PRIORITY_CLASS         = &H00000080
  IDLE_PRIORITY_CLASS         = &H00000040
  NORMAL_PRIORITY_CLASS       = &H00000020
  REALTIME_PRIORITY_CLASS     = &H00000100
End Enum

Public Enum EProcessAccess
  PROCESS_CREATE_PROCESS = &H0080
  PROCESS_CREATE_THREAD = &H0002
  PROCESS_DUP_HANDLE = &H0040
  PROCESS_QUERY_INFORMATION = &H0400
  PROCESS_QUERY_LIMITED_INFORMATION = &H1000
  PROCESS_SET_INFORMATION = &H0200
  PROCESS_SET_QUOTA = &H0100
  PROCESS_SUSPEND_RESUME = &H0800
  PROCESS_TERMINATE = &H0001
  PROCESS_VM_OPERATION = &H0008
  PROCESS_VM_READ = &H0010
  PROCESS_VM_WRITE = &H0020
  SYNCHRONIZE = &H00100000
End Enum

Private Type FileTime
  dwLowDateTime as Long
  dwHighDateTime as long
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

Private pProcessId as long
Private pModuleID as Long
Private pName as string
Private pPath as string

Private pProcessHandle as LongPtr

'Launches a process and creates a stdProcess object for it
'@constructor
'@param {ByVal String} - The command to create the process
'@param {ByVal vbAppWinStyle} - The style to show the window
'@returns {stdProcess} - The launched process
Public Function Create(ByVal sCmd as string, Optional ByVal winStyle As VbAppWinStyle = VbAppWinStyle.vbHide) as stdProcess
    set Create = new stdProcess
    Call Create.protInitFromProcessId(Shell(sCmd,winStyle))
End Function

'Creates a process from a given process id
'@constructor
'@param {ByVal Long} - the process id
'@returns {stdProcess} the queried process
Public Function CreateFromProcessId(ByVal pID as long) as stdProcess
    set CreateFromProcessId = new stdProcess
    Call CreateFromProcessId.protInitFromProcessId(pID)
End Function

'Obtains a the first process which matches the query given
'@constructor
'@param {stdICallable} - The query to search processes for
'@returns {stdProcess} the queried process
Public Function CreateFromQuery(ByVal query as stdICallable) as stdProcess
    Dim pEntry as PROCESSENTRY32: pEntry.dwSize = LenB(pEntry)
    Dim snapshot As LongPtr: snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, ByVal 0&)
    Dim continueSearch as boolean: continueSearch = Process32First(snapshot, pEntry)
    While continueSearch
        Dim oDict as object: set oDict = getQueryDictFromEntry(pEntry)
        if pEnquery.run(oDict) then
            With oDict
                Set CreateFromQuery = New stdProcess
                Call CreateFromQuery.protInitFromProcessId(.item("id"), .item("name"), .item("path"), .item("moduleID"))
                Call CloseHandle(snapshot)
                Exit Function
            End With
        else
            continueSearch = Process32Next(snapshot,pEntry)
        end if
    wend
    Call CloseHandle(snapshot)
End Function

'Obtains a collection of processes which match the query given
'@constructor
'@param {stdICallable} - The query to search processes for
'@returns {Collection<stdProcess>}
Public Function CreateManyFromQuery(ByVal query as stdICallable) as Collection
    Dim pEntry as PROCESSENTRY32: pEntry.dwSize = LenB(pEntry)
    Dim snapshot As LongPtr: snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, ByVal 0&)
    Dim continueSearch as boolean: continueSearch = Process32First(snapshot, pEntry)
    Dim ret as Collection: set ret = new Collection
    While continueSearch
        Dim oDict as object: set oDict = getQueryDictFromEntry(pEntry)
        if pEnquery.run(oDict) then
            With oDict
                Dim proc as stdProcess
                set proc = new stdProcess
                Call proc.protInitFromProcessId(.item("id"), .item("name"), .item("path"), .item("moduleID"))
                ret.add proc
            End With
        end if

        continueSearch = Process32Next(snapshot,pEntry)
    wend
    Call CloseHandle(snapshot)
    set CreateManyFromQuery = ret
End Function

'Obtains a collection of processes which match the query given
'@constructor
'@returns {Collection<stdProcess>} All processes
Public Function CreateAll() As Collection
    Dim pEntry As PROCESSENTRY32: pEntry.dwSize = LenB(pEntry)
    Dim snapshot As LongPtr: snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, ByVal 0&)
    Dim continueSearch As Boolean: continueSearch = Process32First(snapshot, pEntry)
    Dim ret As Collection: Set ret = New Collection
    While continueSearch
        Dim oDict As Object: Set oDict = getQueryDictFromEntry(pEntry)

        With oDict
            Dim proc As stdProcess
            Set proc = New stdProcess
            Call proc.protInitFromProcessId(.Item("id"), .Item("name"), .Item("path"), .Item("moduleID"))
            ret.add proc
        End With

        continueSearch = Process32Next(snapshot, pEntry)

    Wend
    Call CloseHandle(snapshot)
    Set CreateAll = ret
End Function

'================================================================================================
'=      PROTECTED CONSTRUCTORS
'================================================================================================

'Initialises a process object from process id and other data. Optional args are generated if not supplied
'@protected
'@constructor
'@param {ByVal argID} - The process id
'@param {ByVal String = ""} - The process name
'@param {ByVal String = ""} - The process path
'@param {ByVal Long   = 0} - The process module id
'@returns {VOID}
Public Sub protInitFromProcessId(ByVal argID as long, Optional ByVal argName as string = "", Optional ByVal argPath as string = "", Optional ByVal argModuleID as long = 0)
    pProcessId = argID
    pModuleID = argModuleID
    pName = argName
    pPath = argPath
    if argModuleID = 0  then pModuleID = getModuleIDByPid(argID)
    if len(argName) = 0 then pName = getprocessNameById(argID)
    if len(argPath) = 0 then pPath = getProcessImageName(argID)
End Sub

'================================================================================================
'=      PUBLIC PROPERTIES
'================================================================================================

'Obtain the process id
'@returns {Long} - the process id
Public Property Get id() as Long
    id = pProcessId
End Property

'Obtain the process module id
'@returns {Long} - the process module id
Public Property Get moduleID() as long
    moduleID = pModuleID
End Property

'TODO: SessionID
'TODO: ThreadID

'Obtain the process name
'@returns {Long} - the process name
Public Property Get name() as string
    name = pName
End Property

'Obtain the process path
'@returns {Long} - the process path
Public Property Get path() as string
    path = pPath
End Property

'Returns the Winmgmts object representing the process
'@returns {Object<SWbemObjectEx>}
Public Property Get Winmgmt() as Object
    static cache as object
    if cache is nothing then set Cache = GetObject("WINMGMTS:\\.\ROOT\cimv2:Win32_Process.Handle=" & pProcessId)
    set Winmgmt = Cache
End Property

'Attempts to return the command line of the application. Please note [the command line may have changed](https://devblogs.microsoft.com/oldnewthing/20091125-00/?p=15923).
'@returns {String} The string which originally contained the command line arguments passed to this application.
Public Property Get CommandLine() as String
    if not Winmgmt is nothing then CommandLine = Winmgmt.CommandLine
End Property

'Returns whether the process is still running or not
'@returns {Boolean} True if the process is still running, false otherwise
Public Property Get isRunning as Boolean
  isRunning = GetProcessVersion(pProcessId) > 0
End Function

'Returns whether the process is critical or not
'@returns {Boolean} True if the process is critical, false otherwise
Public Property Get isCritical as Boolean
  Call protProcessHandleCreate(PROCESS_QUERY_LIMITED_INFORMATION)
  if not IsProcessCritical(pProcessHandle, isCritical) then
    Err.Raise 1, "stdProcess.isCritical", "Cannot get critical status of process. Error code 0x" & Hex(Err.LastDllError)
  end if
  Call protProcessHandleRelease()
End Function

Public Property Get Priority() as EProcessPriority
  Call protProcessHandleCreate(PROCESS_QUERY_LIMITED_INFORMATION)
  Priority = GetPriorityClass()
  Call protProcessHandleRelease()
End Property


'Get the datetime representing the time the process was started/created
'@returns {String} The datetime representing the time the process was started/created
Public Property Get TimeCreated() as String
  Call protProcessHandleCreate(PROCESS_QUERY_LIMITED_INFORMATION)
  Dim t1 as FileTime, t2 as FileTime, t3 as FileTime, t4 as FileTime
  if GetProcessTimes(pProcessHandle, t1, t2, t3, t4) then
    TimeCreated = t1.dwHighDateTime & t1.dwLowDateTime
  else
    Call Err.Raise(1, "stdProcess.TimeCreated", "Cannot get creation time of process. Error code 0x" & Hex(Err.LastDllError))
  End if
  Call protProcessHandleRelease()
End Property

'Get the datetime representing the time the process was quit
'@returns {String} The datetime representing the time the process was quit
Public Property Get TimeQuit() as String
  Call protProcessHandleCreate(PROCESS_QUERY_LIMITED_INFORMATION)
  Dim t1 as FileTime, t2 as FileTime, t3 as FileTime, t4 as FileTime
  if not isRunning then
    if GetProcessTimes(pProcessHandle, t1, t2, t3, t4) then
      TimeQuit = t2.dwHighDateTime & t2.dwLowDateTime
    else
      Call Err.Raise(1, "stdProcess.TimeQuit", "Cannot get creation time of process. Error code 0x" & Hex(Err.LastDllError))
    End if
  else
    Call Err.Raise(1, "stdProcess.TimeQuit", "Cannot obtain the Exit time for a process which is still running."
  end if
  Call protProcessHandleRelease()
End Property

'Get the amount of time that the process has executed in kernel mode
'@returns {String} The datetime representing the kernel time
Public Property Get TimeKernel as String
  Call protProcessHandleCreate(PROCESS_QUERY_LIMITED_INFORMATION)
  Dim t1 as FileTime, t2 as FileTime, t3 as FileTime, t4 as FileTime
  if GetProcessTimes(pProcessHandle, t1, t2, t3, t4) then
    TimeKernel = t3.dwHighDateTime & t3.dwLowDateTime
  else
    Call Err.Raise(1, "stdProcess.TimeKernel", "Cannot get creation time of process. Error code 0x" & Hex(Err.LastDllError))
  End if
  Call protProcessHandleRelease()
End Property

'Get the amount of time that the process has executed in user mode
'@returns {String} The datetime representing the user time
Public Property Get TimeUser() as String
  Call protProcessHandleCreate(PROCESS_QUERY_LIMITED_INFORMATION)
  Dim t1 as FileTime, t2 as FileTime, t3 as FileTime, t4 as FileTime
  if GetProcessTimes(pProcessHandle, t1, t2, t3, t4) then
    TimeUser = t4.dwHighDateTime & t4.dwLowDateTime
  else
    Call Err.Raise(1, "stdProcess.TimeUser", "Cannot get creation time of process. Error code 0x" & Hex(Err.LastDllError))
  End if
  Call protProcessHandleRelease()
End Property


'@returns {Long} -1 if cannot get ExitCode, otherwise the exit code provided at runtime
Public Property Get ExitCode() as Long
  if not GetExitCodeProcess(pProcessHandle, ExitCode) then
    Err.Raise 1, "stdProcess.ExitCode()", "Cannot get exit code of process. Error code 0x" & Hex(Err.LastDllError)
  end if
End Property

'================================================================================================
'=      PUBLIC METHODS
'================================================================================================

'Wait till the process closes
'@returns {VOID}
Public Sub forceQuit(Optional ByVal ExitCode as long = 0)
  Call protProcessHandleCreate(PROCESS_TERMINATE)
  if pProcessHandle = 0 then Exit Sub
  if not TerminateProcess(pProcessHandle, ExitCode) then
    Err.Raise 1, "stdProcess#ForceQuit()", "Cannot terminate process. Error code 0x" & Hex(Err.LastDllError)
  end if
  Call protProcessHandleRelease()
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
Public Property Get protProcessHandle as LongPtr
  protProcessHandle = pProcessHandle
End Property

'Call to OpenProcess() to save a handle in the class. Typically this is not required, thus this method is listed as protected.
'@protected
'@returns {VOID}
Public Sub protProcessHandleCreate(ByVal access as EProcessAccess)
  if pProcessHandle = 0 then
    pProcessHandle = OpenProcess(access, 0, pProcessId)
  else
    Err.Raise 1, "", "Process already open."
  end if
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
Private Function getQueryDictFromEntry(entry as PROCESSENTRY32) as Object
    Dim ret as object: set ret = CreateObject("Scripting.Dictionary")
    ret("id") = entry.th32ProcessID
    ret("moduleID") = entry.th32ModuleID
    ret("name") = Left(entry.szexeFile, InStr(1, entry.szexeFile, vbNullChar) - 1)
    ret("path") = getProcessImageName(entry.th32ProcessID)
    set getQueryDictFromEntry = ret
End Function

'Obtain the Process name from the ProcessID
'@param {ByVal Long} - The process id
'@returns {Long} - The process name
Private Function getProcessNameById(Byval lProcessID as Long) as String
    Dim entry as PROCESSENTRY32: entry = GetCachedProcessEntry(lProcessID)
    getProcessNameById = Left(entry.szexeFile, InStr(1, entry.szexeFile, vbNullChar) - 1)
End Function

'Obtain the ModuleID from the ProcessID
'@param {ByVal Long} - The process id
'@returns {Long} - The module id
Private Function getModuleIDByPid(Byval lProcessID as Long) as Long
    Dim entry as PROCESSENTRY32: entry = getCachedProcessEntry(lProcessID)
    getModuleIDByPid = entry.th32ModuleID
End Function

'Obtain the PROCESSENTRY32 struct from the ProcessID
'@param {ByVal Long} - The process id
'@returns {PROCESSENTRY32} - The PROCESSENTRY32 representing the process' data
Private Function getCachedProcessEntry(ByVal lProcessID as Long, Optional ByVal Override as Boolean = false) as PROCESSENTRY32
    static pEntry as PROCESSENTRY32
    static stcProcessID as long
    if stcProcessID <> lProcessID or Override then
        stcProcessID = lProcessID
        pEntry.dwSize = LenB(pEntry)

        Dim snapshot As LongPtr: snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, ByVal 0&)
        Dim continueSearch as boolean: continueSearch = Process32First(snapshot, pEntry)
        While continueSearch
            if pEntry.th32ProcessID = lProcessID then
                continueSearch = false
            else
                continueSearch = Process32Next(snapshot,pEntry)
            end if
        wend
        Call CloseHandle(snapshot)
    end if
    GetCachedProcessEntry = pEntry
End Function

'Obtain the full qualified path of a process from it's ProcessID
'@param {ByVal Long} - The process id
'@returns {String} - The fully qualified path to the process
Private Function getProcessImageName(ByVal lProcessID As Long) As String
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
    else
        Debug.print "Error in getProcessImageName - 0x" & Hex(Err.LastDllError)
    End If
    getProcessImageName = sBuf
End Function

