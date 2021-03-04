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
  dwDesiredAccess as Long, _ 
  bInheritHandle as Boolean, _ 
  dwProcessId as Long _ 
) as LongPtr


Const MAX_PATH = 260
Const TH32CS_SNAPPROCESS = &H2&
Const PROCESS_QUERY_INFORMATION = &H400


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
Public Function CreateFromExisting(ByVal query as stdICallable) as stdProcess
    Dim pEntry as PROCESSENTRY32: pEntry.dwSize = LenB(pEntry)
    Dim snapshot As LongPtr: snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, ByVal 0&)
    Dim continueSearch as boolean: continueSearch = Process32First(snapshot, pEntry)
    While continueSearch
        Dim oDict as object: set oDict = getQueryDictFromEntry(pEntry)
        if pEnquery.run(oDict) then
            With oDict
                Set CreateFromExisting = New stdProcess
                Call CreateFromExisting.protInitFromProcessId(.item("id"), .item("name"), .item("path"), .item("moduleID"))
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
Public Function CreateFromAllExisting(ByVal query as stdICallable) as Collection
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
        else
            continueSearch = Process32Next(snapshot,pEntry)
        end if
    wend
    Call CloseHandle(snapshot)
    set CreateFromAllExisting = ret
End Function

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
    set Winmgmt = GetObject("WINMGMTS:\\.\ROOT\cimv2:Win32_Process.Handle=" & pProcessId)
End Property













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
    Dim sBuf As String
    Dim sChar As Long

    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 0, lProcessID)
    If hProcess Then
        sBuf = String$(MAX_PATH, Chr$(0))
        Call QueryFullProcessImageNameA(hProcess, 0, sBuf, MAX_PATH)
        sBuf = Left$(sBuf, sChar)
    End If
    getProcessImageName = sBuf
    Call CloseHandle(hProcess)
End Function


