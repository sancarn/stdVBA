Attribute VB_Name = "stdProcessTests"

Public Sub testAll()
  Test.Topic "stdProcess"

  Dim proc1 as stdProcess
  set proc1 = stdProcess.Create("calc.exe")
  Test.Assert "Process ID given", proc1.id > 0

  Dim proc2 as stdProcess

  'Test CreateFromProcessId
  set proc2 = stdProcess.CreateFromProcessId(proc1.id)
  Test.Assert "CreateFromProcessId works", proc1.name = proc2.name and proc1.path = proc2.path and proc1.id = proc2.id

  'Test CreateFromProcessId
  set proc2 = stdProcess.CreateFromQuery(stdLambda.Create("$1.id = " & proc1.id))
  Test.Assert "CreateFromQuery works", proc1.name = proc2.name and proc1.path = proc2.path and proc1.id = proc2.id

  Dim col as collection
  set col = stdProcess.CreateManyFromQuery(stdLambda.Create("$1.id = " & proc1.id))
  Test.Assert "CreateManyFromQuery works 1", col.count = 1
  Test.Assert "CreateManyFromQuery works 2", proc1.name = col(1).name and proc1.path = col(1).path and proc1.id = col(1).id

  set proc2 = stdProcess.Create("calc.exe")
  set col = stdProcess.CreateManyFromQuery(stdLambda.Create("$1.name = ""calc.exe"""))
  Test.Assert "CreateManyFromQuery works 3 many processes open", col.count = 2

  Test.Assert "CreateMany() returns more than 10 processes", stdProcess.CreateAll().count >= 10

  Call proc2.ForceQuit(10)
  Test.Assert "Proc2 forcequit exit code", proc2.ExitCode = 10


  'TODO: Name
  'TODO: Path
  'TODO: Winmgmt
  'TODO: CommandLine
  'TODO: isRunning
  'TODO: isCritical
  'TODO: Priority
  'TODO: TimeCreated
  'TODO: TimeQuit
  'TODO: TimeKernel
  'TODO: TimeUser
  'TODO: Test WaitForQuit()
End Sub