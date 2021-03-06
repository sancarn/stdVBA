Attribute VB_Name = "stdProcessTests"

Public Sub testAll()
  Test.Topic "stdProcess"

  'Definitions
  Dim proc1 as stdProcess, proc2 as stdProcess

  'Test launching of processes
  set proc1 = stdProcess.Create("calc.exe")

  'Test basic properties
  Test.Assert "Process ID given", proc1.id > 0
  Test.Assert "Name", proc1.name = "calc.exe"
  Test.Assert "Path", proc1.path = "C:\windows\system32\calc.exe"

  'Test CreateFromProcessId
  set proc2 = stdProcess.CreateFromProcessId(proc1.id)
  Test.Assert "CreateFromProcessId works", proc1.name = proc2.name and proc1.path = proc2.path and proc1.id = proc2.id

  'Test CreateFromQuery - should return 1 process object
  set proc2 = stdProcess.CreateFromQuery(stdLambda.Create("$1.id = " & proc1.id))
  Test.Assert "CreateFromQuery works", proc1.name = proc2.name and proc1.path = proc2.path and proc1.id = proc2.id

  'Ensure the CreateManyFromQuery returns a collection, and ensure query lambda is running in check (although this is terribly innefficient)
  Dim col as collection
  set col = stdProcess.CreateManyFromQuery(stdLambda.Create("$1.id = " & proc1.id))
  Test.Assert "CreateManyFromQuery works 1", col.count = 1
  Test.Assert "CreateManyFromQuery works 2", proc1.name = col(1).name and proc1.path = col(1).path and proc1.id = col(1).id

  'Create new process for proc2, and attempt to grab all calc.exes
  set proc2 = stdProcess.Create("calc.exe")
  set col = stdProcess.CreateManyFromQuery(stdLambda.Create("$1.name = ""calc.exe"""))
  Test.Assert "CreateManyFromQuery works 3 many processes open", col.count >= 2

  'Ensure CreateAll() returns more than the default number of processes
  Test.Assert "CreateAll() returns more than 10 processes", stdProcess.CreateAll().count >= 10

  'Attempt to force quit proc2, and retrieve the exit code in the next step
  Call proc2.ForceQuit(10)
  Test.Assert "Proc2 forcequit exit code", proc2.ExitCode = 10

  'Test to ensure proc1 is running, but proc1 isn't after exit
  Test.Assert "isRunning 1", proc1.isRunning
  Test.Assert "isRunning 2", not proc1.isRunning

  'Attempt to get wnmgmt of proc1
  Test.Assert "Winmgmt", not is nothing proc1.Winmgmt
  
  'Test crabbing of command line; Calc.exe seems to not consume any of the command line args
  set proc2 = stdProcess.Create("calc.exe --testMode")
  Test.Assert "CommandLine", proc2.CommandLine <> ""
  
  'Assume all are NORMAL priority on startup from Shell()
  Test.Assert "Priority", proc1.priority = NORMAL_PRIORITY_CLASS  

  'TODO: isCritical
  'TODO: TimeCreated
  'TODO: TimeQuit
  'TODO: TimeKernel
  'TODO: TimeUser
  'TODO: Test WaitForQuit()
End Sub