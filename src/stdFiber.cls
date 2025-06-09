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
'Sub main()
'    Dim fibers as new Collection
'    For i = 1 to 100
'        fibers.add createFiber("http://example.com/" & i)
'    next
'    
'    'Run requests concurrently across 10 fibers
'    set results = stdFiber.runFibers(fibers, numFibers:=10)
'End Sub
'Function createFiber(ByVal url as string) as stdFiber
'    Dim fib as stdFiber: set fib = stdFiber.Create()
'    set fib.meta("status") = getStatusUpdater()
'    fib.add stdCallback.CreateFromMethods("", "step1_GetRequest").bind(url)
'    fib.add stdCallback.CreateFromMethods("", "step2_AwaitStatus")
'    fib.add stdCallback.CreateFromMethods("", "step3_ReturnData")
'    fib.addErrorHandler stdCallback.CreateFromMethods("", "stepX_OnError")
'    set createFiber = fib
'End Function
'Public Function step1_GetRequest(ByVal url as string, ByVal fib as stdFiber) as boolean
'    fib.meta("status").value = "step1"
'    'Store information in bespoke fiber procedure metadata
'    set fib.meta("req") = stdHTTP.Create(url, HTTPAsync)
'    step1_GetRequest = true
'End Function
'Public Function step2_AwaitStatus(ByVal fib as stdFiber) as boolean
'    fib.meta("status").value = "step2"
'    step1_GetRequest = fib.meta("req").status <> 0
'End Function
'Public Function step3_ReturnData(ByVal fib as stdFiber) as boolean
'    Call fib.resolve(fib.meta("req").ResponseText)
'    fib.meta("status").value = "Complete"
'End Function
'Public Sub stepX_OnError(ByVal fib as stdFiber)
'    fib.meta("status").value = fib.ErrorText
'End Sub

'```
'@example **2. Use of agents**:
'```vb
'Sub main()
'  Dim fibers as new Collection
'  For i = 1 to 100
'    fibers.add createFiber("http://example.com/" & i)
'  next
'  set results = stdFiber.runFibers(fibers, numFibers:=10)
'End Sub
'Function createFiber(ByVal url as string) as stdFiber
'  Dim fib as stdFiber: set fib = stdFiber.Create()
'  fib.add stdCallback.CreateFromPointer(AddressOf step1_CreateAgent, vbBoolean, Array(vbObject))
'  fib.add stdCallback.CreateFromPointer(AddressOf step2_Navigate, vbBoolean, Array(vbObject, vbString)).bind(url)
'  fib.add stdCallback.CreateFromPointer(AddressOf step3_Await, vbBoolean, Array(vbObject))
'  set createFiber = fib
'End Function
'Function step1_CreateAgent(ByVal fib as stdFiber) as boolean
'  'An agent is unique to the fiber runtime and can be used across different fiber procedures.
'  set fib.agent("ie") = CreateObject("InternetExplorer.Application")
'  fib.agent("ie").visible = true
'  step1_CreateAgent = true
'End Function
'Function step2_Navigate(ByVal url as string, ByVal fib as stdFiber) as Boolean
'  fib.agent("ie").Navigate(url)
'End Function
'Function step3_Await(ByVal fib as stdFiber) as Boolean
'  step3_Await = fib.agent("ie").readystate = 4
'End Function
''...
'```


Private Type TFiberQueue
  index As Long
  count As Long
  queue() As stdFiber
End Type
Private Type TProc
  proc As stdICallable
  name As String
  startTime As Date
End Type

'TODO: Consider sectioning off data between runtime data TFiberRuntime (procStartTime, meta, agent, ...) and TFiberBuildTime (name, procsCount, procs)
Private Type TThis
  name As String
  procStep As Long
  procsCount As Long
  resetCount As Long
  procs() As TProc
  fiberStartTime As Date
  meta As Object
  isFinished As Boolean
  result As Variant
  errorText As String
  errorHandler As stdICallable
  stepChangeHandler As stdICallable
  agent As Object
  endTime As Date
End Type
Private This As TThis

'Create a fiber
'@constructor
'@param sName - The name of the fiber
'@returns - A fiber procedure - a wrapper around a number of procedures used for concurrent programming.
Public Function Create(ByVal sName As String) As stdFiber
  Set Create = New stdFiber
  Call Create.protInit(sName)
End Function

'Create a fiber
'@constructor
'@protected
'@param sName - The name of the fiber
Friend Sub protInit(ByVal sName As String)
  This.name = sName
  This.procStep = 1
  Set This.meta = CreateObject("Scripting.Dictionary")
End Sub

'Add a step/sub-process to the fiber
'@param processStep as stdICallable<(stdFiber)=>boolean> - Callback process / step, consumes the fiber parent, returns `true` if caller can proceed to next step. returns `false` otherwise.
'@param stepName - The name of the step. This can be used for resetting and error reporting purposes.
'@returns - A fiber procedure - A reference to the fiber the step has been added to
Public Function add(ByVal processStep As stdICallable, Optional ByVal stepName As String = "") As stdFiber
  This.procsCount = This.procsCount + 1
  ReDim Preserve This.procs(1 To This.procsCount)
  With This.procs(This.procsCount)
    Set .proc = processStep
    .name = stepName
  End With
  Set add = Me
End Function

'Add an error handler to the fiber
'@param handler as stdICallable<(stdFiber)=>void> - Callback to run when an error occurs in the fiber
'@returns - A fiber procedure - A reference to the fiber the step has been added to, to allow chaining
Public Function addErrorHandler(ByVal handler As stdICallable) As stdFiber
  Set This.errorHandler = handler
  Set addErrorHandler = Me
End Function

'Add a handler which is triggered at every step change
'@param handler as stdICallable<(stdFiber)=>void> - Callback to run when an error occurs in the fiber
'@returns - A fiber procedure - A reference to the fiber the step has been added to, to allow chaining
Public Function addStepChangeHandler(ByVal handler As stdICallable) As stdFiber
  Set This.stepChangeHandler = handler
  Set addStepChangeHandler = Me
End Function

'Calls current sub-process in the fiber procedure. Advances the queue if the ran sub-process is finished.
'@param agent as Object<Dictionary> - An object storing information unique to the runtime agent. This agent is bound to the fiber, allowing use of window or accessibility elements initialised on the agent.
'@returns - `True` if this fiber procedure has finished execution (i.e. future stepping not required anymore).
Friend Function protStep(ByVal agent As Object) As Boolean
  Dim oMe As Object: Set oMe = Me
  On Error GoTo ErrorHandler
    If This.fiberStartTime = 0 Then This.fiberStartTime = Now()
    If This.procs(This.procStep).startTime = 0 Then This.procs(This.procStep).startTime = Now()
    
    Set This.agent = agent
    If This.procs(This.procStep).proc.Run(oMe) Then
      This.procStep = This.procStep + 1
      This.isFinished = This.procStep > This.procsCount
      If This.isFinished Then
        This.endTime = Now()
      Else
        If Not This.stepChangeHandler Is Nothing Then Call This.stepChangeHandler.Run(oMe)
      End If
    End If
    protStep = This.isFinished
    Exit Function
  On Error GoTo 0
ErrorHandler:
  protStep = True
  This.isFinished = True
  This.result = Null
  This.endTime = Now()
  This.errorText = Err.Description & "(" & Err.Number & ", " & Err.Source & ")"
  Call This.errorHandler.Run(oMe)
End Function

'Usually called by a sub-process, this will mark the fiber procedure as finished, and return the passed result
'@param vResult - The result to return
Public Sub Resolve(ByVal vResult As Variant)
  This.isFinished = True
  This.endTime = Now()
  Call CopyVariant(This.result, vResult)
End Sub

'Reset to a specific step
'@param sName - Step name to jump to (if not specified, will reset to first step).
Public Sub Reset(Optional ByVal sName As String = "")
  This.resetCount = This.resetCount + 1
  If sName <> "" Then
    Dim i As Long
    For i = 1 To This.procsCount
      If This.procs(i).name = sName Then
        Dim j As Long
        For j = i To This.procStep
          This.procs(j).startTime = 0
        Next
        This.procStep = i
        Exit Sub
      End If
    Next
  Else
    This.procStep = 1
  End If
End Sub

'Obtain the number of times this fiber has reset.
'@returns - The number of times this fiber has been reset.
Public Property Get ResetCount() As Long
  ResetCount = This.resetCount
End Property

'Obtain the result of this fiber
'@returns - The result of this fiber
Public Property Get Result() As Variant
  Call CopyVariant(Result, This.result)
End Property

'Obtain errors if any as string
'@returns - The error text
Public Property Get ErrorText() As String
  ErrorText = This.errorText
End Property

'Obtain the metadata object, usually used by the sub-process
'@returns - The metadata object
Public Property Get Meta() As Object
  Set Meta = This.meta
End Property

'Obtain the current step which the fiber procedure is at.
'@returns - The current step
Public Property Get Step() As Long
  Step = This.procStep
End Property

'Obtains the name of the current step in the fiber
'@returns - The name of the fiber step
Public Property Get StepName() As String
  StepName = This.procs(This.procStep).name
End Property

'Obtains the duration of the fiber procedure
'@returns - The duration of the fiber procedure in seconds
Public Property Get StepDuration() As Long
  StepDuration = DateDiff("s", This.procs(This.procStep).startTime, Now())
End Property

'Obtains the duration of the fibers entire runtime
'@returns - The duration of the fibers entire runtime in seconds
Public Property Get TotalDuration() As Long
  TotalDuration = DateDiff("s", This.procs(This.procsCount).startTime, This.endTime)
End Property

'Obtains whether the fiber has finished or not
'@returns - `True` if the fiber has finished, `False` otherwise
Public Property Get isFinished() As Boolean
  isFinished = This.isFinished
End Property

'Obtain the name of the fiber
'@returns - The name of the fiber
Public Property Get Name() As String
  Name = This.name
End Property

'Obtain the agent information I.E. information unique to the runtime running this fiber
'@returns - The agent information. 
Public Property Get Agent() As Object
  Set Agent = This.agent
End Property

'Run a collection/stdArray/stdEnumerator of fiber procedures, across a number of concurrent running fiber runtimes, returning all results as a collection.
'@static
'@param vFibers as Variant<Collection|Array|stdArray|stdEnumerator> - Fiber procedures to execute over
'@param numAgents - the number of fiber runtimes to create (or number of fiber procedures to run concurrently)
'@param agentInitCB as stdICallable<(agent as Object)=>void> - Called when an agent has been initialised. This is called only once per agent object. The agent environment is received as the parameter.
'@param agentDestroyCB as stdICallable<(agent as Object)=>void> - Called when an agent is about to be destroyed. This is called only once per agent object. The agent environment is received as the parameter. Typically used to cleanup.
'@param runtimeUpdateCB as stdICallable<(iFinishedCount as long, iFiberCount as long)=>void> - A callback to call every second to update on progress.
'@returns - Collection of results of ran fiber procedures. You can use the variant from param1 if you need to check for errors using ErrorText.
'@TODO: Consider whether return data is correct. If we create fibers from table, how would we refer back to errornous records on the table?
Public Function runFibers(ByVal vFibers As Variant, Optional ByVal numAgents As Long = 1, Optional ByVal agentInitCB As stdICallable = Nothing, Optional ByVal agentDestroyCB As stdICallable = Nothing, Optional ByVal runtimeUpdateCB As stdICallable = Nothing) As Collection
  'Intiialise
  Dim fibers As TFiberQueue: fibers = TFiberQueue_Create(vFibers)
  Dim running() As stdFiber: ReDim running(1 To numAgents)
  Dim agents() As Object: ReDim agents(1 To numAgents)
  Dim i as long
  For i = 1 To numAgents
    Set agents(i) = CreateObject("Scripting.Dictionary")
    agents(i)("id") = i
    If Not agentInitCB Is Nothing Then Call agentInitCB.Run(agents(i))
  Next

  'Await completion of all processes
  Dim iRuntimeClock As Long: iRuntimeClock = 0
  Dim isFinished As Boolean
  Do While Not isFinished
    Dim iAgentsRunning As Long: iAgentsRunning = 0
    For i = 1 To numAgents
      'Step fiber if it exists, if it finishes remove it to signify this agent is ready for a new fiber
      If Not running(i) Is Nothing Then
        If running(i).protStep(agents(i)) Then
          Set running(i) = Nothing
        End If
      End If

      'Get next fiber if nothing is running in this agent
      If running(i) Is Nothing Then
        If TFiberQueue_HasNext(fibers) Then
          Set running(i) = TFiberQueue_Get(fibers)
          iAgentsRunning = iAgentsRunning + 1
        End If
      Else
        iAgentsRunning = iAgentsRunning + 1
      End If

      DoEvents
    Next

    'Process is finished only when no agents are running
    isFinished = iAgentsRunning = 0

    'Update on progress every 1k cycles
    iRuntimeClock = iRuntimeClock + 1
    If iRuntimeClock Mod 1000 = 0 Then
      iRuntimeClock = 0
      If Not runtimeUpdateCB Is Nothing Then
        Dim iFiberCount As Long: iFiberCount = fibers.count
        Dim iFinishedCount As Long: iFinishedCount = fibers.index - iAgentsRunning
        Call runtimeUpdateCB.Run(iFinishedCount, iFiberCount)
      End If
    End If
  Loop

  'Cleanup agents
  For i = 1 To numAgents
    If Not agentDestroyCB Is Nothing Then Call agentDestroyCB.Run(agents(i))
    Set agents(i) = Nothing 'objects will still be present on the individual fibers until end of this scope
  Next

  'Collect results
  Dim results As Collection: Set results = New Collection
  For i = 1 To fibers.count
    results.Add fibers.queue(i).Result
  Next

  Set runFibers = results
End Function

'Cancel the running fiber
'@param errorText - If an error is required, set this parameter to whatever the text should be.
Public Sub Cancel(Optional ByVal errorText As String = "")
  This.errorText = errorText
  This.isFinished = True
  This.endTime = Now()
  If Not This.errorHandler Is Nothing And errorText <> "" Then
    Dim oMe As Object: Set oMe = Me
    Call This.errorHandler.Run(oMe)
  End If
End Sub

'***********
'* HELPERS *
'***********

'Copy variant from vSrc to vDest
'@param vDest - Destination
'@param vSrc  - Source
Private Sub CopyVariant(ByRef vDest, ByVal vSrc)
  If IsObject(vSrc) Then
    Set vDest = vSrc
  Else
    Let vDest = vSrc
  End If
End Sub

'Creates a queue of fiber procedures (TFiberQueue)
'@param vFibers as Variant<Collection|Array|stdArray|stdEnumerator> - Fibers to execute over
'@returns - Queue representing fibers
Private Function TFiberQueue_Create(ByRef vFibers As Variant) As TFiberQueue
  TFiberQueue_Create.index = 0
  Dim i As Long
  Select Case VarType(vFibers)
    Case vbObject
      Select Case TypeName(vFibers)
        Case "Collection"
          ReDim TFiberQueue_Create.queue(1 To vFibers.count)
          Dim oFiber As stdFiber
          For Each oFiber In vFibers
            i = i + 1
            Set TFiberQueue_Create.queue(i) = oFiber
          Next
        Case "stdArray"
          Dim vQueue() As Variant: vQueue = vFibers.arr
          ReDim TFiberQueue_Create.queue(1 To vFibers.Length)
          For i = 1 To vFibers.length
            Set TFiberQueue_Create.queue(i) = vQueue(i)
          Next
        Case "stdEnumerator"
          Dim oQueue() As Object: oQueue = vFibers.asArray(vbObject)
          ReDim TFiberQueue_Create.queue(1 To vFibers.Length)
          For i = 1 To vFibers.length
            Set TFiberQueue_Create.queue(i) = oQueue(i)
          Next
      End Select
    Case Is > vbArray
      Dim iLB As Long: iLB = LBound(vFibers)
      Dim iLen As Long: iLen = UBound(vFibers) - iLB + 1
      ReDim TFiberQueue_Create.queue(1 To iLen)
      For i = 0 To iLen - 1
        Set TFiberQueue_Create.queue(i + 1) = vFibers(i + iLB)
      Next
  End Select
  TFiberQueue_Create.count = UBound(TFiberQueue_Create.queue)
End Function

'Gets the next element in a queue
'@param queue - The queue to get the item from
'@returns - The resultant fiber
Private Function TFiberQueue_Get(ByRef queue As TFiberQueue) As stdFiber
  queue.index = queue.index + 1
  Set TFiberQueue_Get = queue.queue(queue.index)
End Function

'Checks if a next element exists in a queue
'@param queue - The queue to check
'@returns - The result
Private Function TFiberQueue_HasNext(ByRef queue As TFiberQueue) As Boolean
  TFiberQueue_HasNext = queue.index < queue.count
End Function