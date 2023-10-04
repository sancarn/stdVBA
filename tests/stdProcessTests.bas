Attribute VB_Name = "stdProcessTests"
'@lang VBA

Private proc As stdProcess
Private bProcWaited As Boolean
Public Sub testAll()
  Test.Topic "stdProcess"

  if Test.FullTest then
    'Definitions
    Dim proc1 As stdProcess, proc2 As stdProcess

    'Test launching of processes
    Set proc1 = stdProcess.Create("notepad.exe")
    
    'Test basic properties
    Test.Assert "Process ID given", proc1.id > 0
    Test.Assert "Name", proc1.name = "notepad.exe"
    Test.Assert "Path", proc1.path = "C:\Windows\System32\notepad.exe"

    'Test CreateFromProcessId
    Set proc2 = stdProcess.CreateFromProcessId(proc1.id)
    Test.Assert "CreateFromProcessId works", proc1.name = proc2.name And proc1.path = proc2.path And proc1.id = proc2.id

  
    'Test CreateFromQuery - should return 1 process object
    Set proc2 = stdProcess.CreateFromQuery(stdLambda.Create("$1.id = " & proc1.id))
    Test.Assert "CreateFromQuery works", proc1.name = proc2.name And proc1.path = proc2.path And proc1.id = proc2.id

    'Ensure the CreateManyFromQuery returns a collection, and ensure query lambda is running in check (although this is terribly innefficient)
    Dim col As Collection
    Set col = stdProcess.CreateManyFromQuery(stdLambda.Create("$1.id = " & proc1.id))
    Test.Assert "CreateManyFromQuery works 1", col.Count = 1
    Test.Assert "CreateManyFromQuery works 2", proc1.name = col(1).name And proc1.path = col(1).path And proc1.id = col(1).id

    'Create new process for proc2, and attempt to grab all calc.exes
    Set proc2 = stdProcess.Create("notepad.exe")
    Set col = stdProcess.CreateManyFromQuery(stdLambda.Create("$1.name = ""notepad.exe"""))
    Test.Assert "CreateManyFromQuery works - many processes open", col.Count >= 2

    'Ensure CreateAll() returns more than the default number of processes
    Test.Assert "CreateAll() returns more than 10 processes", stdProcess.CreateAll().Count >= 10

    'Attempt to force quit proc2, and retrieve the exit code in the next step
    Call proc2.forceQuit(10)
    Test.Assert "Proc2 forcequit exit code", proc2.ExitCode = 10
  

    'Test to ensure proc1 is running, but proc1 isn't after exit
    Test.Assert "isRunning 1", proc1.isRunning
    Test.Assert "isRunning 2", Not proc2.isRunning

    'Attempt to get wnmgmt of proc1
    Test.Assert "Winmgmt", Not proc1.Winmgmt Is Nothing
    
    'Test crabbing of command line; Calc.exe seems to not consume any of the command line args
    Set proc2 = stdProcess.Create("notepad.exe --testMode")
    Test.Assert "CommandLine", proc2.CommandLine <> ""
    proc2.forceQuit
    
    'Assume all are NORMAL priority on startup from Shell()
    Test.Assert "Priority", proc1.Priority = NORMAL_PRIORITY_CLASS
    Test.Assert "isCritical", proc1.isCritical = False
    
    'These are methods which are still WIP
    Test.Assert "TimeCreated", proc1.TimeCreated < Now()
    Test.Assert "TimeQuit", proc2.TimeQuit < Now()
    
    'Still not sure of the purpose of TimeKernel and time user as they seem to continuously supply 01/01/1601
    'Test.Assert "TimeKernel", proc1.TimeKernel <> ""
    'Test.Assert "TimeUser", proc1.TimeUser <> ""
    
    Set proc = proc1
    bProcWaited = False
    Application.OnTime Now() + TimeSerial(0, 0, 1), "testStdProcessTestsQuit"
    Call proc1.waitTilClose
    Test.Assert "waitTilClose waited", bProcWaited

  end if
End Sub
Sub testStdProcessTestsQuit()
    bProcWaited = True
    proc.forceQuit
End Sub
