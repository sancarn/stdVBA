# `stdWindow`

## Spec

### Constructors

#### `Public Function Create(ByVal sCmd As String, Optional ByVal winStyle As VbAppWinStyle = VbAppWinStyle.vbHide) As stdProcess`

Launches a process and creates a stdProcess object for it

```vb
'Open windows explorer
set proc = stdProcess.Create("explorer.exe",vbNormalFocus)

'Open notepad for 5 seconds
With stdProcess.Create("notepad.exe", vbNormalFocus)
  Debug.Print .id
  Debug.Print .TimeCreated
  While Round((CDbl(Now()) - CDbl(.TimeCreated)) * 60 * 60 * 24) - 3600 < 5
    'Debug.Print Round((CDbl(Now()) - CDbl(.TimeCreated)) * 60 * 60 * 24) - 3600
    DoEvents
  Wend
  
  .forceQuit
  .waitTilClose
  Debug.Print .TimeQuit
End With

'Pass command line params into 
set proc = stdProcess.Create("someProcess.exe --param1 --param2=""something""",vbNormalFocus)
```

#### `Public Function CreateFromProcessId(ByVal pID As Long) As stdProcess`

Creates a process from a given process id

```vb

With someWindow
  set proc = stdProcess.CreateFromProcessId(.ProcessId)
End With

```

#### `Public Function CreateFromQuery(ByVal query As stdICallable) As stdProcess`

Obtains a the first process which matches the query given

```vb
stdProcess.CreateFromQuery(stdLambda.Create("$1.name like ""calculator*"""))
```

#### `Public Function CreateManyFromQuery(ByVal query As stdICallable) As Collection<stdProcess>`

Obtains a collection of processes which match the query given

```vb
set col = stdProcess.CreateManyFromQuery(stdLambda.Create("$1.name like ""iexplore*"""))
```

#### `Public Function CreateAll() As Collection<stdProcess>`

Obtains a collection of all processes

```vb
For each proc in stdProcess.CreateAll()
  '...
next
```

### Instance Properties

#### `Public Property Get id() As Long`

Obtains the ProcessID (aka PID), from the process. This identifier can be used for a plethera of things including searching for the window owning the process etc.

```vb
set proc = stdProcess.Create("...")
stdWindow.CreateFromDesktop().FindFirst(stdLambda.Create("$2.ProcessID = $1").bind(proc.id))
```

#### `Public Property Get name() As String`

The name of the executable file for the process.

```vb
Debug.Print stdProcess.Create("notepad.exe").name   'returns "notepad.exe"
```

#### `Public Property Get path() As String`

The path of the executable file for the process.

```vb
Debug.Print stdProcess.Create("notepad.exe").path   'returns "C:\Windows\System32\notepad.exe"
```

#### `Public Property Get Winmgmt() As Object`

Returns the Windows Management object representing the process.

```vb
set oWndMan = proc.Winmgmt
```

#### `Public Property Get CommandLine() As String`

Retrieves `CommandLine` property from `Winmgmts` object.

```vb
Debug.print notepad.CommandLine
```

#### `Public Property Get isRunning() As Boolean`

Returns `true` if the process is running, otherwise `false`.

```vb
While proc.isRunning
  DoEvents
Wend
```

#### `Public Property Get isCritical() As Boolean`

Returns `true` if the process is critical, `false` otherwise.

```vb
if not proc.isCritical then
  'do something
end if
```

#### `Public Property Get Priority() As EProcessPriority`

The priority of the process. This should return a value corresponding to the priority displayed in task manager:

![priority](./assets/stdProcess-priority.png)

```vb
debug.print proc.priority
```

#### `Public Property Get TimeCreated() As Date`

The time at which this process was created.

> Note: Currently this function returns UTC time. Usage of this function is advised against until this bug is fixed.

```vb
'Wait roughly 10 seconds, then close the process
set proc = stdProcess.Create("notepad.exe")
While DateDiff("s", now(), proc.TimeCreated) < 10
  DoEvents
Wend

```

#### `Public Property Get TimeQuit() As Date`

The time at which this process was destroyed.

> Note: Currently this function returns UTC time. Usage of this function is advised against until this bug is fixed.

```vb
'Wait roughly 10 seconds, then close the process
set proc = stdProcess.Create("notepad.exe")
While DateDiff("s", now(), proc.TimeCreated) < 10
  DoEvents
Wend

'Quit and log time
Call proc.forceQuit()
Call proc.waitTilClose()
Debug.Print proc.TimeQuit
```

#### `Public Property Get TimeKernel() As Date`

Get the amount of time that the process has executed in kernel mode

> Note: Currently this function returns UTC time. Usage of this function is advised against until this bug is fixed.

```vb
Debug.print proc.TimeKernel()
```

#### `Public Property Get TimeUser() As Date`

Get the amount of time that the process has executed in user mode

> Note: Currently this function returns UTC time. Usage of this function is advised against until this bug is fixed.

```vb
Debug.print proc.TimeUser()
```

#### `Public Property Get ExitCode() As Long`

Get the exit code of this process.

> Note: An exit code is only ever received if the process has ended. Check `isRunning` prior to calling this function.

```vb
'Either pass the exit code in and check it later:
proc.forceQuit(127)
'...
Call proc.waitTilClose
Debug.Print proc.ExitCode 'returns 127


'Or use it to diagnose an issue with a process:
if not anotherProc.isRunning then Debug.Print "Process quit with code " & anotherProc.exitCode

```

### Instance Methods

#### `Public Sub forceQuit(Optional ByVal ExitCode As Long = 0)`

Force quit an application, optionally passing exit code.

```vb
'Open notepad and close it immediately
set proc = stdProcess.Create("notepad.exe")
Call proc.forceQuit()
```

#### `Public Sub waitTilClose()`

Wait for a process to close. Literally wait for isRunning to return `false`.

```vb
set proc = stdProcess.Create("notepad.exe")
Call proc.forceQuit()
Call proc.waitTilClose()
```

### Protected methods and properties

#### ProcessHandleManagement

`stdProcess` exposes the following 3 subs/properties. These can be used to alter the state of the process using low level functions.

* `Friend Property Get protProcessHandle() As LongPtr`
* `Friend Sub protProcessHandleCreate(ByVal access As EProcessAccess)`
* `Public Sub protProcessHandleRelease()`

An example is used in `forceQuit` below:

```vb
Sub forceQuit(proc as stdProcess)
  Call proc.protProcessHandleCreate(PROCESS_TERMINATE)
  If proc.pProcessHandle = 0 Then Exit Sub
  
  'Note: TerminateProcess can return a weird boolean where `bool` and `Not bool` both result in `True`, which is nonsense...
  'for this reason we explicitely cast to a long here...
  If CLng(TerminateProcess(proc.protProcessHandle, ExitCode)) = 0 Then
    Err.Raise Err.LastDllError, "stdProcess#ForceQuit()", "Cannot terminate process. Error code 0x" & Hex(Err.LastDllError)
  End If
  Call proc.protProcessHandleRelease
End Sub
```

#### `Friend Sub protInitFromProcessId(ByVal argID As Long, Optional ByVal argName As String = "", Optional ByVal argPath As String = "", Optional ByVal argModuleID As Long = 0)`

Initialise a `stdProcess` object from process ID. Can optionally pass other arguments to prevent having to requery data, and improve performance.

> Note: This is an internal method which should only be called if you are creating a factory.

```vb
set o = new stdProcess
Call o.protInitFromProcessId(myID)
```