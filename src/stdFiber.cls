VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdFiber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@module
'`stdFiber` allows production of partitioned asynchronous procedures which can be run within a fiber runtime.
'`stdFiber.Create()` is used to create our fiber procedures
'`stdFiber.runFibers()` is used to run our fiber procedures within a fiber runtime.
'Fibers can be useful especially when dealing with asynchronous functions
'@example **1. Simple fiber example**:
'```vb
'   Sub main()
'       Dim fibers as new Collection
'       For i = 1 to 100
'           fibers.add createFiber("http://example.com/" & i)
'       next
'       
'       'Run requests concurrently across 10 fibers
'       set results = stdFiber.runFibers(fibers, numFibers:=10)
'   End Sub
'   Function createFiber(ByVal url as string) as stdFiber
'       Dim fib as stdFiber: set fib = stdFiber.Create()
'       set fib.meta("status") = getStatusUpdater()
'       fib.add stdCallback.CreateFromMethods("", "step1_GetRequest").bind(url)
'       fib.add stdCallback.CreateFromMethods("", "step2_AwaitStatus")
'       fib.add stdCallback.CreateFromMethods("", "step3_ReturnData")
'       fib.addErrorHandler stdCallback.CreateFromMethods("", "stepX_OnError")
'       set createFiber = fib
'   End Function
'   Public Function step1_GetRequest(ByVal url as string, ByVal fib as stdFiber) as boolean
'       fib.meta("status").value = "step1"
'       'Store information in bespoke fiber procedure metadata
'       set fib.meta("req") = stdHTTP.Create(url, HTTPAsync)
'       step1_GetRequest = true
'   End Function
'   Public Function step2_AwaitStatus(ByVal fib as stdFiber) as boolean
'       fib.meta("status").value = "step2"
'       step1_GetRequest = fib.meta("req").status <> 0
'   End Function
'   Public Function step3_ReturnData(ByVal fib as stdFiber) as boolean
'       Call fib.resolve(fib.meta("req").ResponseText)
'       fib.meta("status").value = "Complete"
'   End Function
'   Public Sub stepX_OnError(ByVal fib as stdFiber)
'       fib.meta("status").value = fib.ErrorText
'   End Sub

'```
'@example **2. Use of agents**:
'```vb
'   Sub main()
'     Dim fibers as new Collection
'     For i = 1 to 100
'       fibers.add createFiber("http://example.com/" & i)
'     next
'     set results = stdFiber.runFibers(fibers, numFibers:=10)
'   End Sub
'   Function createFiber(ByVal url as string) as stdFiber
'     Dim fib as stdFiber: set fib = stdFiber.Create()
'     fib.add stdCallback.CreateFromPointer(AddressOf step1_CreateAgent, vbBoolean, Array(vbObject))
'     fib.add stdCallback.CreateFromPointer(AddressOf step2_Navigate, vbBoolean, Array(vbObject, vbString)).bind(url)
'     fib.add stdCallback.CreateFromPointer(AddressOf step3_Await, vbBoolean, Array(vbObject))
'     set createFiber = fib
'   End Function
'   Function step1_CreateAgent(ByVal fib as stdFiber) as boolean
'     'An agent is unique to the fiber runtime and can be used across different fiber procedures.
'     set fib.agent("ie") = CreateObject("InternetExplorer.Application")
'     fib.agent("ie").visible = true
'     step1_CreateAgent = true
'   End Function
'   Function step2_Navigate(ByVal url as string, ByVal fib as stdFiber) as Boolean
'     fib.agent("ie").Navigate(url)
'   End Function
'   Function step3_Await(ByVal fib as stdFiber) as Boolean
'     step3_Await = fib.agent("ie").readystate = 4
'   End Function
'   '...
'```


Private Type TFiberQueue
  index as long
  count as long
  queue() as stdFiber
End Type
Private Type TProc
  proc as stdICallable
  name as string
  startTime as Date
End Type

'TODO: Consider sectioning off data between runtime data TFiberRuntime (procStartTime, meta, agent, ...) and TFiberBuildTime (name, procsCount, procs)
Private Type TThis
  name as string
  procStep as Long
  procsCount as long
  resetCount as long
  procs() as TProc
  fiberStartTime as Date
  meta as object
  isFinished as boolean
  result as Variant
  errorText as string
  errorHandler as stdICallable
  agent as object
End Type
Private This as TThis

'Create a fiber
'@constructor
'@param sName - The name of the fiber
'@returns - A fiber procedure - a wrapper around a number of procedures used for concurrent programming.
Public Function Create(ByVal sName as string) as stdFiber
  set Create = new stdFiber
  Call Create.protInit(sName)
End Function

'Create a fiber
'@constructor
'@protected
'@param sName - The name of the fiber
Friend Sub protInit(ByVal sName as string)
  this.name = sName
  this.procStep = 1
  set this.meta = CreateObject("Scripting.Dictionary")
End Sub

'Add a step/sub-process to the fiber
'@param processStep as stdICallable<(stdFiber)=>boolean> - Callback process / step, consumes the fiber parent, returns `true` if caller can proceed to next step. returns `false` otherwise.
'@param stepName - The name of the step. This can be used for resetting and error reporting purposes.
'@returns - A fiber procedure - A reference to the fiber the step has been added to
Public Function add(ByVal processStep as stdICallable, Optional ByVal stepName as string = "") as stdFiber
  this.procsCount = this.procsCount + 1
  Redim Preserve this.procs(1 to this.procsCount)
  With this.procs(this.procsCount)
    set .proc = processStep
    .name = stepName
  end with
  set add = Me
End Function

'Add an error handler to the fiber
'@param handler as stdICallable<(stdFiber)=>void> - Callback to run when an error occurs in the fiber
'@returns - A fiber procedure - A reference to the fiber the step has been added to, to allow chaining
Public Function addErrorHandler(ByVal handler as stdICallable) as stdFiber
  set this.errorHandler = handler
  set addErrorHandler = Me
End Function

'Calls current sub-process in the fiber procedure. Advances the queue if the ran sub-process is finished.
'@param agent as Object<Dictionary> - An object storing information unique to the runtime agent. This agent is bound to the fiber, allowing use of window or accessibility elements initialised on the agent.
'@returns - `True` if this fiber procedure has finished execution (i.e. future stepping not required anymore).
Friend Function protStep(Byval agent as Object) as Boolean
  Dim oMe As Object: Set oMe = Me
  On Error GoTo ErrorHandler
    If this.fiberStartTime = 0 then this.fiberStartTime = now()
    if this.procs(this.procStep).startTime = 0 then this.procs(this.procStep).startTime = now()
    
    set this.agent = agent
    if this.procs(this.procStep).proc.Run(oMe) then 
      this.procStep = this.procStep + 1
      this.isFinished = this.procStep > this.procsCount
    end if
    protStep = this.isFinished
    Exit Function
  On Error GoTo 0
ErrorHandler:
  protStep = True
  this.isFinished = true
  this.result = Null
  this.ErrorText = Err.Description & "(" & Err.Number & ", " & Err.Source & ")"
  Call this.errorHandler.Run(oMe)
End Function

'Usually called by a sub-process, this will mark the fiber procedure as finished, and return the passed result
'@param vResult - The result to return
Public Sub Resolve(ByVal vResult as variant)
  this.isFinished = true
  Call CopyVariant(this.result, vResult)
End Sub

'Reset to a specific step
'@param sName - Step name to jump to (if not specified, will reset to first step).
Public Sub Reset(Optional ByVal sName As string = "")
  this.resetCount = this.resetCount + 1
  if sName <> "" then
    Dim i as long
    For i = 1 to this.procsCount
      if this.procs(i).name = sName then
        Dim j as Long
        For j = i to this.procStep
          this.procs(j).startTime = 0
        next
        This.procStep = i
        exit sub
      end if
    next
  else
    this.procStep = 1
  end if
End Sub

'Obtain the number of times this fiber has reset.
'@returns - The number of times this fiber has been reset.
Public Property Get ResetCount() as Long
  ResetCount = this.resetCount
End Property

'Obtain the result of this fiber
'@returns - The result of this fiber
Public Property Get Result() as Variant
  Call CopyVariant(Result, this.result)
End Property

'Obtain errors if any as string
'@returns - The error text
Public Property Get ErrorText() As string
  ErrorText = this.errorText
End Property

'Obtain the metadata object, usually used by the sub-process
'@returns - The metadata object
Public Property Get Meta() as Object
  set Meta = this.Meta
End Property

'Obtain the current step which the fiber procedure is at.
'@returns - The current step
Public Property Get Step() as Long
  Step = procStep
End Property

'Obtains the duration of the fiber procedure
'@returns - The duration of the fiber procedure in seconds
Public Property Get StepDuration as Long
  StepDuration = DateDiff("s", this.procs(this.procStep).startTime, now())
End Property

'Obtains whether the fiber has finished or not
'@returns - `True` if the fiber has finished, `False` otherwise
Public Property Get isFinished() as Boolean
  isFinished = this.isFinished
End Property

'Obtain the name of the fiber
'@returns - The name of the fiber
Public Property Get Name() as string
  Name = this.Name
End Property

'Obtain the agent information I.E. information unique to the runtime running this fiber
'@returns - The agent information. 
Public Property Get Agent() as Object
  set Agent = this.Agent
End Property

'Run a collection/stdArray/stdEnumerator of fiber procedures, across a number of concurrent running fiber runtimes, returning all results as a collection.
'@param vFibers as Variant<Collection|Array|stdArray|stdEnumerator> - Fiber procedures to execute over
'@param numFibers - the number of fiber runtimes to create (or number of fiber procedures to run concurrently)
'@returns - Collection of results of ran fiber procedures. You can use the variant from param1 if you need to check for errors using ErrorText.
'@TODO: Consider whether return data is correct. If we create fibers from table, how would we refer back to errornous records on the table?
'@TODO: add param AgentInitHandler   for initialising of agent objects, this is called only once per agent.
'@TODO: add param AgentDestroyHandler  for destroying of agent objects / clean up, this is called only once per agent.
Public Function runFibers(ByVal vFibers as variant, Optional byval numFibers as Long = 1) as Collection
  'Intiialise
  Dim fibers as TFiberQueue: fibers = TFiberQueue_Create(vFibers)
  Dim running() as stdFiber: ReDim running(1 to numFibers)
  Dim agents() as Object: Redim agents(1 to numFibers)
  For i = 1 to numFibers
    set agents(i) = CreateObject("Scripting.Dictionary")
    agents(i)("id") = i
  next

  'Await completion of all processes
  Dim isFinished as Boolean
  Do While not isFinished
    isFinished = true
    For i = 1 to numFibers
      if running(i) is nothing then
        if TFiberQueue_HasNext(fibers) then
          set running(i) = TFiberQueue_Get(fibers)
          isFinished = false
        end if
      elseif running(i).protStep(agents(i)) then
        set running(i) = nothing
        isFinished = False
      else
        isFinished = false
      end if

      DoEvents
    next
  Loop

  'Collect results
  Dim results as Collection: set results = new Collection
  For i = 1 to fibers.count
    results.add fibers.queue(i).result
  next

  set runFibers = results
End Function

'***********
'* HELPERS *
'***********

'Copy variant from vSrc to vDest
'@param vDest - Destination
'@param vSrc  - Source
Private Sub CopyVariant(ByRef vDest, ByVal vSrc)
  if isObject(vSrc) then
    set vDest = vSrc
  else
    let vDest = vSrc
  End if
End Sub

'Creates a queue of fiber procedures (TFiberQueue)
'@param vFibers as Variant<Collection|Array|stdArray|stdEnumerator> - Fibers to execute over
'@returns - Queue representing fibers
Private Function TFiberQueue_Create(ByRef vFibers as Variant) as TFiberQueue
  TFiberQueue_Create.index = 0
  Dim i as long
  select case vartype(vFibers)
    case vbObject
      select case typename(vFibers)
        Case "Collection"
          Redim TFiberQueue_Create.queue(1 to vFibers.count)
          Dim oFiber as stdFiber
          For each oFiber in vFibers
            i=i+1
            set TFiberQueue_Create.queue(i) = oFiber
          next
        case "stdArray"
          Dim vQueue() as Variant: vQueue = vFibers.arr
          Redim TFiberQueue_Create.queue(1 to vFibers.Length)
          For i = 1 to vFibers.length
            set TFiberQueue_Create.queue(i) = vQueue(i)
          next
        case "stdEnumerator"
          Dim oQueue() as Object: oQueue = vFibers.asArray(vbObject)
          Redim TFiberQueue_Create.queue(1 to vFibers.Length)
          For i = 1 to vFibers.length
            set TFiberQueue_Create.queue(i) = oQueue(i)
          next
      end select
    case Is > vbArray
      Dim iLB as Long: iLB = lbound(vFibers)
      Dim iLen as Long: iLen = ubound(vFibers) - iLB + 1
      Redim TFiberQueue_Create.queue(1 to iLen)
      For i = 0 to iLen-1
        set TFiberQueue_Create.queue(i+1) = vFibers(i+iLB)
      next
  end select
  TFiberQueue_Create.count = ubound(TFiberQueue_Create.queue)
End Function    

'Gets the next element in a queue
'@param queue - The queue to get the item from
'@returns - The resultant fiber
Private Function TFiberQueue_Get(ByRef queue as TFiberQueue) as stdFiber
  queue.index = queue.index+1
  set TFiberQueue_Get = queue.queue(queue.index)
End Function

'Checks if a next element exists in a queue
'@param queue - The queue to check
'@returns - The result
Private Function TFiberQueue_HasNext(ByRef queue as TFiberQueue) as Boolean
  TFiberQueue_HasNext = queue.index < queue.count
End Function