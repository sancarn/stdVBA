VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdSentry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@description
'In numerous cases using excel we like to set standard options to optimise performance:
'```vb
'Dim bEnableEvents as Boolean: bEnableEvents = Application.EnableEvents
'Dim bScreenUpdating as Boolean: bScreenUpdating = Application.ScreenUpdating
'Application.EnableEvents = false
'Application.ScreenUpdating = false
'
'  ... our code ...
'
'Application.EnableEvents = bEnableEvents
'Application.ScreenUpdating = bScreenUpdating
'```
'Similarly with std libraries, we have to add and pop from the stack:
'```vb
'stdError.addStack "myMethod"
'  ... our code ...
'stdError.popStack
'```
'stdPerformance today uses the sentry pattern:
'```vb
'With stdPerformance.CreateMeasure("Some measurement")
'  ... our code ...
'End With
'```
'Although the two patterns have the same size (3 lines), the sentry pattern is both more beautiful and more robust.
'This class is an attempt to make the sentry pattern easier to implement in VBA.
'```vb
'Dim ensureScreenUpdating as stdSentry
'set ensureScreenUpdating = stdSentry.Create( _ 
'  stdLambda.CreateMultiLine( _ 
'    Array( _ 
'      "let temp = Application.ScreenUpdating", _
'      "let Application.ScreenUpdating = $1", _
'      "Array(temp)", _
'    ) _ 
'  ), _
'  stdLambda.Create("let Application.ScreenUpdating = $1") _ 
')
'With ensureScreenUpdating(False)
'  '... our code ...
'End With
'```
'@example A standard sentry for optimising performance of range operations:
'```vb
'With stdSentry.CreateOptimiser(ScreenUpdating:=False, Calculation:=xlCalculationManual)
'  '... heavy range based operations ...
'End With
'```
'@example A sentry for error stack management:
'```vb
'Sub MyMethod()
'  With stdSentry.CreateErrorStack("MyMethod")
'    '... our code ...
'  End With
'End Sub
'```
'@example A sentry for AutomationSecurity settings:
'```vb
'With stdSentry.CreateFromAutomationSecurity(msoAutomationSecurityForceDisable)
'  Call ThisWorkbook.SaveAs("C:\temp\test.xlsx")
'End With
'```
'@example A sentry for setting/resetting an object's properties: 
'```vb
'Debug.print Application.DisplayAlerts '=> (Whatever before)
'With stdSentry.CreateFromObjectProperty(Application, "DisplayAlerts", False)
'  Debug.Print Application.DisplayAlerts '=> False
'End With
'Debug.Print Application.DisplayAlerts '=> (Whatever before)
'```



Implements stdICallable
Private Type TFactory
  OnInit as stdICallable
  OnDestroy as stdICallable
  passVars as boolean
End Type
Private Type TInstance
  onDestroy as stdICallable
  DestroyArgs as Variant
  passVars as boolean
End Type
Private Type TThis
  Factory as TFactory
  Instance as TInstance
End Type
Private This as TThis

'Create a sentry factory
'@constructor
'@param OnInit - Function to run when the sentry is created.
'@param OnDestroy - Function to run when the sentry is destroyed.
'@param passVars - If true, the result of OnInit is passed to OnDestroy. If false, no arguments are passed to OnDestroy.
'@returns - The sentry factory
'@remark - This creates a sentry factory, which can be used to create sentries. The factory is not a sentry itself, and will have to be called in order to create a sentry.
'@example ```vb
'Dim factory as stdSentry: set factory = stdSentry.Create( _
'  stdCallback.CreateFromModuleMethod("ModuleName", "OnInit"), _
'  stdCallback.CreateFromModuleMethod("ModuleName", "OnDestroy") _
')
'With factory()
'  'whatever code you want here
'End With
'```
Public Function Create(ByVal OnInit as stdICallable, ByVal OnDestroy as stdICallable, Optional ByVal passVars as boolean = true) as stdSentry
  set Create = new stdSentry
  Call Create.protInitFactory(OnInit, OnDestroy, passVars)
End Function

'Create a sentry which acts as an "optimiser" for code. Note, setting options doesn't always lead to performance improvements. See [vba-articles performance thread](https://sancarn.github.io/vba-articles/performance-tips.html) for details.
'@constructor
'@param EnableEvents   - If not Empty, sets Application.EnableEvents to this value. If Empty, Application.EnableEvents remains unchanged.
'@param ScreenUpdating - If not Empty, sets Application.ScreenUpdating to this value. If Empty, Application.ScreenUpdating remains unchanged.
'@param Calculation    - If not -1, sets Application.Calculation to this value. If -1, Application.Calculation remains unchanged.
'@returns - A sentry which will restore the original settings when destroyed.
'@requires stdLambda
'@example ```vb
'With stdSentry.CreateOptimiser(ScreenUpdating:=False, Calculation:=xlCalculationManual)
'  '... heavy range based operations ...
'End With
'```
Public Function CreateOptimiser(Optional ByVal EnableEvents as vbTriState = vbUseDefault, Optional ByVal ScreenUpdating as vbTriState = vbUseDefault, Optional ByVal Calculation as Long = -1) as stdSentry
  if not isObject(stdLambda) then Err.Raise 5, "stdSentry.CreateFromObjectMethod", "stdLambda is required for this function"
  
  'Validate inputs
  if not (Calculation = -1 or Calculation = 2 or Calculation = -4135  or Calculation = -4105) then Err.Raise 5, "stdSentry.CreateOptimiser", "Calculation must be either xlCalculationAutomatic, xlCalculationSemiautomatic, xlCalculationManual, or -1"
  
  'Convert Calculation to a variant
  Dim vCalculation as Variant: vCalculation = IIf(Calculation = -1, Empty, Calculation)

  'Create the OnInit and OnDestroy functions
  '@remark vbUseDefault == -2
  Dim onInit as stdICallable: Set onInit = stdLambda.CreateMultiLine(Array( _
    "let t1 = -2", _ 
    "let t2 = -2", _ 
    "let t3 = -2", _
    "If $1 <> -2 Then let t1 = Application.EnableEvents   : let Application.EnableEvents = $1    end", _
    "If $2 <> -2 Then let t2 = Application.ScreenUpdating : let Application.ScreenUpdating = $2  end", _
    "If $3 <> -2 Then let t3 = Application.Calculation    : let Application.Calculation = $3     end", _
    "Array(t1,t2,t3)" _
  ))
  Dim onDestroy as stdICallable: Set onDestroy = stdLambda.CreateMultiLine(Array( _
    "If $1 <> -2 Then let Application.EnableEvents = $1 end", _
    "If $2 <> -2 Then let Application.ScreenUpdating = $2 end", _
    "If $3 <> -2 Then let Application.Calculation = $3 end" _
  ))
  
  'Create the sentry
  set CreateOptimiser = Create(onInit, onDestroy).Run(EnableEvents, ScreenUpdating, vCalculation)
End Function

'Create a sentry from a property of an object
'@constructor
'@param obj - Object to bind the property to
'@param sPropertyName - Name of the property to bind
'@param value - Value to set the property to
'@returns - The sentry
'@requires stdLambda
'@example ```vb
'With stdSentry.CreateFromObjectProperty(Application, "DisplayAlerts", false)
'  Call ThisWorkbook.SaveAs("C:\temp\test.xlsx")
'End With
'```
Public Function CreateFromObjectProperty(ByVal obj as Object, ByVal sPropertyName as String, ByVal value as Variant) as stdSentry
  if not isObject(stdLambda) then Err.Raise 5, "stdSentry.CreateFromObjectMethod", "stdLambda is required for this function"
  Dim onInit as stdICallable: Set onInit = stdLambda.CreateMultiLine(Array( _
    "let t1 = $1." & sPropertyName, _ 
    "let $1." & sPropertyName & " = $2", _
    "Array(t1)" _
  )).bind(obj)
  Dim onDestroy as stdICallable: Set onDestroy = stdLambda.CreateMultiLine(Array( _
    "let $1." & sPropertyName & " = $2" _
  )).bind(obj)

  set CreateFromObjectProperty = Create(OnInit, OnDestroy).Run(value)
End Function

'Create a sentry from a method of an object
'@constructor
'@param obj - Object to bind the method to
'@param sOnCreateName - Name of the method to call when the sentry is created
'@param sOnDestroyName - Name of the method to call when the sentry is destroyed
'@returns - The sentry
'@requires stdCallback
'@example ```vb
'With stdSentry.CreateFromObjectMethod(stdError, "AddStack", Array("MethodName"), "PopStack", Array())
'  '... our code ...
'End With
'```
Public Function CreateFromObjectMethod(ByVal obj as Object, ByVal OnInitName as String, ByVal initParams as Variant, ByVal OnDestroyName as String, ByVal destroyParams as variant) as stdSentry
  if not isObject(stdCallback) then Err.Raise 5, "stdSentry.CreateFromObjectMethod", "stdCallback is required for this function"
  Dim onInit as stdICallable: set onInit = stdCallback.CreateFromObjectMethod(obj, OnInitName).BindEx(initParams)
  Dim onDestroy as stdICallable: set onDestroy = stdCallback.CreateFromObjectMethod(obj, OnDestroyName).BindEx(destroyParams)
  set CreateFromObjectMethod = Create(onInit, onDestroy, false).Run()
End Function

'Create a sentry for error stack management
'@constructor
'@param name - Name of the stack to add
'@returns - The sentry
'@requires stdError, stdCallback
'@example ```vb
'Public Sub MyMethod()
'  With stdSentry.CreateErrorStack("MyClass.MyMethod")
'    '... our code ...
'  End With
'End Sub
'```
Public Function CreateErrorStack(ByVal name as String) as stdSentry
  if not isObject(stdError) then Err.Raise 5, "stdSentry.CreateFromObjectMethod", "stdError is required for this function"
  set CreateErrorStack = CreateFromObjectMethod(stdError, "AddStack", Array(name), "PopStack", Array()).Run()
End Function

'Create a sentry from an AutomationSecurity setting
'@constructor
'@param sec - AutomationSecurity setting desired
'@returns - The sentry
'@requires stdLambda
'@example ```vb
'With stdSentry.CreateFromAutomationSecurity(msoAutomationSecurityForceDisable)
'  Call ThisWorkbook.SaveAs("C:\temp\test.xlsx")
'End With
'```
Public Function CreateFromAutomationSecurity(ByVal sec as MsoAutomationSecurity) as stdSentry
  set CreateFromAutomationSecurity = CreateFromObjectProperty(Application, "AutomationSecurity", sec)
End Function

'Initialise the sentry factory
'@protected
'@param OnInit - Function to run when the sentry is created. Should return arguments to pass to OnDestroy. If OnInit returns Empty, no arguments are passed to OnDestroy.
'@param OnDestroy - Function to run when the sentry is destroyed. Should take the arguments returned by OnInit.
Public Sub protInitFactory(ByVal OnInit as stdICallable, ByVal OnDestroy as stdICallable, ByVal passVars as boolean)
  With This.Factory
    Set .OnInit = OnInit
    Set .OnDestroy = OnDestroy
    .passVars = passVars
  End With
End Sub

'Initialise the sentry instance
'@protected
'@param onDestroy - Function to run when the sentry is destroyed.
'@param args - Arguments to pass to the OnDestroy function
'@param passVars - If true, `args` will be passed to OnDestroy. If false, no arguments are passed to OnDestroy.
Public Sub protInitInstance(ByVal onDestroy as stdICallable, ByVal args as Variant, ByVal passVars as boolean)
  With This.Instance
    Set .onDestroy = onDestroy
    .DestroyArgs = args
    .passVars = passVars
  End With
End Sub

'Run the sentry
'@param args - Arguments to pass to the OnInit function
'@returns - The sentry
Public Function Run(ParamArray args() as Variant) as stdSentry
Attribute Run.VB_UserMemId = 0
  Dim copy: copy = args
  set Run = RunEx(copy)
End Function

'Run the sentry
'@param args - Arguments to pass to the OnInit function
'@returns - The sentry
Public Function RunEx(ByVal args as Variant) as stdSentry
  'Run the OnInit function and prepare destroy arguments
  Dim destroyArgs: destroyArgs = This.Factory.OnInit.RunEx(args)
  if IsEmpty(destroyArgs) then destroyArgs = Array()

  'Return the sentry
  set RunEx = new stdSentry
  Call RunEx.protInitInstance(This.Factory.OnDestroy, destroyArgs, This.Factory.passVars)
End Function

'Bind arguments to the sentry factory
'@param args - Arguments to bind to the sentry
'@returns - The sentry
Public Function Bind(ParamArray args() as Variant) as stdSentry
  Dim copy: copy = args
  set Bind = BindEx(copy)
End Function

'Bind arguments to the sentry factory
'@param args as Variant<Array<Variant>> - Arguments to bind to the sentry
'@returns - The sentry
Public Function BindEx(ByVal args as Variant) as stdSentry
  set BindEx = new stdSentry
  'TODO: 
End Function

'Making late-bound calls to `stdSentry` members. Each object which implements `stdICallable`
'will support a different set of latebound calls.
'@protected
'@param sMessage - Message to send. Standard messages include "obj" returning the object, "className" returning the class name. Other messages are implementation specific.
'@param success  - Whether the call was successful
'@param params   - Any variant, typically parameters as an array. Passed along with the message.
'@returns - Any return value.
Public Function SendMessage(ByVal sMessage as string, ByRef success as boolean, ByVal params as variant) as Variant
  select case sMessage
    case "obj":
      success = true
      set SendMessage = Me
    case "className":
      success = true
      SendMessage = "stdSentry"
    case "OnInit":
      success = true
      set SendMessage = This.Factory.OnInit
    case "OnDestroy":
      success = true
      set SendMessage = This.Factory.OnDestroy
    case else
      success = false
  end select
End Function


Private Sub Class_Terminate()
  With This.Instance
    if not .OnDestroy is nothing then
      if this.Instance.passVars then
        Call .OnDestroy.RunEx(.DestroyArgs)
      else
        Call .OnDestroy.Run
      end if
    End if
  End With
End Sub

'Copies a variant
'@param vOut - Variant to copy to
'@param vIn - Variant to copy from
Private Sub CopyVariant(ByRef vOut, ByVal vIn)
  If isObject(vIn) Then
    Set vOut = vIn
  Else
    vOut = vIn
  End If
End Sub

'Implement stdICallable
Private Function stdICallable_Bind(ParamArray params() As Variant) As stdICallable
  Dim v: v = params
  Call CopyVariant(stdICallable_Bind, BindEx(v))
End Function
Private Function stdICallable_Run(ParamArray params() As Variant) As Variant
  Dim v: v = params
  Call CopyVariant(stdICallable_Run, RunEx(v))
End Function
Private Function stdICallable_RunEx(ByVal params As Variant) As Variant
  Call CopyVariant(stdICallable_RunEx, RunEx(params))
End Function
Private Function stdICallable_SendMessage(ByVal sMessage As String, success As Boolean, ByVal params As Variant) As Variant
  Call CopyVariant(stdICallable_SendMessage, SendMessage(sMessage, success, params))
End Function

