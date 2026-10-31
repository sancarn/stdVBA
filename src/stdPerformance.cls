VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdPerformance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@module
'@description A class which implements a performance testing and optimisation library.
'Use this class to measure performance and optimise your code.
'@example Simple example of measuring performance:
'```vb
'Const C_MAX as Long = 1000
'With stdPerformance.CreateMeasure("Some test expression", C_MAX)
'  Dim v
'  For i = 1 to C_MAX
'    v = expressionToTest()
'  next
'End With
'```
'@example Simple example of optimising performance:
'```vb
'With stdPerformance.CreateOptimiser()
'  'some heavy code here
'End With
'```

'Declares for performance counters
#If Mac Then
   #If MAC_OFFICE_VERSION >= 16 Then
      Private Type TMachTimebase
        Numer As Long
        Denom As Long
      End Type
      Private Declare PtrSafe Function mach_absolute_time Lib "/usr/lib/libSystem.dylib" () As LongLong
      Private Declare PtrSafe Function mach_timebase_info Lib "/usr/lib/libSystem.dylib" (ByRef info As TMachTimebase) As Long
   #End If
#Else ' Win32 or Win64
   #If VBA7 Then
      Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
      Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef counter As Currency) As Long
      Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef frequency As Currency) As Long
   #Else
      Private Declare Function GetTickCount Lib "kernel32" () As Long
      Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef counter As Currency) As Long
      Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef frequency As Currency) As Long
   #End If
#End If

'Enum for sentry object type
Public Enum EPerfObjType
  iMeasure = 1
  iOptimiser = 2
End Enum


'For Performance Optimisation sentries
Private Type TOptimiser
  App As Object
  HasScreenUpdating As Boolean
  HasEnableEvents As Boolean
  HasCalculation As Boolean
  AccessEchoScope As Long
  EnableEvents As Boolean
  ScreenUpdating As Boolean
  Calculation As Long
End Type

Private Type TAccessEchoScope
  Active As Boolean
  EchoOn As Boolean
End Type

'For Performance Measurement
Private Type TMeasure
  StartTime As Double
  HighResolution As Boolean
  Counter As Currency
  Frequency As Currency
#If Mac And MAC_OFFICE_VERSION >= 16 Then
  MachStart As LongLong
  MillisecondsPerTick As Double
#End If
  KeyName As String
  Divider As Double
End Type

'For Measurement storage
Private Type FakeDictItem
  key As String
  val As Variant
End Type

Private Type TSingleton
  FakeDict() As FakeDictItem   'was Private FakeDict() as FakeDictItem
  AccessEchoScopes() As TAccessEchoScope
  AccessEchoBaseline As Boolean
  AccessEchoCount As Long
End Type

Private Type TThis
  Singleton As TSingleton

  ObjectType As EPerfObjType
  Measure As TMeasure
  Optimiser As TOptimiser
End Type
Private This As TThis




'Create
'@constructor
'@param objType - Type of performance object to create. iMeasure - used for measuring performance, iOptimiser - used for optimising performance 
'@param params as Variant<Array<Variant>> - Additional params supplied as array.
'@returns stdPerformance<EPerfObjType> - Object termination has special behaviour. See Measure and Optimise methods for further details.
'@remark 2023-11-23 This function is not intended to be called directly and thus has been made private. Use the Measure or Optimise methods instead. If absolutely required use `protInit`.
Private Function Create(ByVal objType As EPerfObjType, ByVal params As Variant) As stdPerformance
  Set Create = New stdPerformance
  Call Create.protInit(objType, params)
End Function

'Initialises the class
'@protected
'@param objType - Type of performance object to initialise.
'
'* iMeasure - used for measuring performance
'* iOptimiser - used for optimising performance 
'@param params as Variant<Array<Variant>> - Additional params supplied as array.
Public Sub protInit(ByVal objType As EPerfObjType, ByVal params As Variant)
  This.ObjectType = objType
  Select Case objType
    Case iMeasure
      This.Measure.KeyName = params(0)
      This.Measure.Divider = params(1)
      Call startMeasure
    Case iOptimiser
      Dim oApp As Object: Set oApp = Application
      Set This.Optimiser.App = oApp
      'TODO: Extend scoped optimiser support beyond Excel, Word, and Access.
      Select Case oApp.Name
        Case "Microsoft Access"
          Dim echoOnExit As Boolean: echoOnExit = True
          If UBound(params) >= 3 Then echoOnExit = params(3)
          This.Optimiser.AccessEchoScope = stdPerformance.protRegisterAccessEcho(params(0), echoOnExit)
          Call oApp.Echo(params(0))
        Case "Microsoft Word"
          This.Optimiser.ScreenUpdating = oApp.ScreenUpdating
          This.Optimiser.HasScreenUpdating = True
          oApp.ScreenUpdating = params(0)
        Case "Microsoft Excel"
          This.Optimiser.ScreenUpdating = oApp.ScreenUpdating
          This.Optimiser.HasScreenUpdating = True
          This.Optimiser.EnableEvents = oApp.EnableEvents
          This.Optimiser.HasEnableEvents = True
          This.Optimiser.Calculation = oApp.Calculation
          This.Optimiser.HasCalculation = True
          oApp.ScreenUpdating = params(0)
          oApp.EnableEvents = params(1)
          oApp.Calculation = params(2)
        Case Else
          Err.Raise 5, "stdPerformance.CreateOptimiser", "Unsupported application: " & oApp.Name
      End Select
  End Select
End Sub

'Create a measure object
'@constructor
'@param sProc   - Name of method or block to measure
'@param nCount - Operation count for printed per-operation time; stored measurements remain total milliseconds. Default = 1.
'@returns stdPerformance<iMeasure> - Object which upon termination, adds measurement of block to global cache
'@example ```vb
'With stdPerformance.CreateMeasure("Some test expression")
'  Dim v
'  For i = 1 to 1000
'    v = expressionToTest()
'  next
'End With
'```
'@example
'```vb
'Const C_MAX as Long = 1000
'With stdPerformance.CreateMeasure("Some test expression", C_MAX)
'  Dim v
'  For i = 1 to C_MAX
'    v = expressionToTest()
'  next
'End With
'```
'@remark Uses Windows performance counters or Mach absolute time on Office 16+ for Mac; older Mac hosts use Timer. Results include scope-release overhead.
'@remark Mach time excludes sleep. API failure uses the fallback clock captured at entry; fallback intervals must be shorter than 49.7 days on Windows or one day on Mac.
Public Function CreateMeasure(ByVal sProc As String, Optional ByVal nCount As Double = 1) As stdPerformance
  Set CreateMeasure = Create(iMeasure, Array(sProc, nCount))
End Function

'Optimise
'@constructor
'@param ScreenUpdating  - ScreenUpdating set value
'@param EnableEvents    - EnableEvents set value
'@param Calculation     - Calculation set value
'@returns stdPerformance<iOptimiser> - Object termination has special behaviour. See Measure and Optimise methods for further details.
'@remark Calculation is defined as long instead of xlCalculation so the function continues to work without compile error in Word, Powerpoint etc.
'@example
'```vb
'With stdPerformance.CreateOptimiser()
'  'some heavy code here
'End With
'```
'@param AccessEchoOnExit - Access has no Echo getter. State to restore after the outermost Access scope; defaults to True.
'@remark In Access, ScreenUpdating controls Echo. EnableEvents and Calculation apply only to Excel.
'@remark Nested Access scopes retain the latest active scope, even when objects are released out of order.
Public Function CreateOptimiser(Optional ByVal ScreenUpdating As Boolean = False, Optional ByVal EnableEvents As Boolean = False, Optional ByVal Calculation As Long = -4135, Optional ByVal AccessEchoOnExit As Boolean = True) As stdPerformance
  Set CreateOptimiser = Create(iOptimiser, Array(ScreenUpdating, EnableEvents, Calculation, AccessEchoOnExit))
End Function

'Register an Access repainting scope. Called on the singleton only.
'@protected
'@param echoOn - Repainting state requested by this scope.
'@param echoOnExit - State to restore once the outermost scope ends.
'@returns - Identifier used to unregister the scope.
Public Function protRegisterAccessEcho(ByVal echoOn As Boolean, ByVal echoOnExit As Boolean) As Long
  If This.Singleton.AccessEchoCount = 0 Then
    ReDim This.Singleton.AccessEchoScopes(0 To 0)
    This.Singleton.AccessEchoBaseline = echoOnExit
  End If
  Dim index As Long: index = UBound(This.Singleton.AccessEchoScopes) + 1
  ReDim Preserve This.Singleton.AccessEchoScopes(0 To index)
  This.Singleton.AccessEchoScopes(index).Active = True
  This.Singleton.AccessEchoScopes(index).EchoOn = echoOn
  This.Singleton.AccessEchoCount = This.Singleton.AccessEchoCount + 1
  protRegisterAccessEcho = index
End Function

'Remove a scope and return the repainting state of the newest remaining scope.
'@protected
'@param index - Identifier returned by protRegisterAccessEcho.
'@returns - Repainting state to apply after removing this scope.
Public Function protUnregisterAccessEcho(ByVal index As Long) As Boolean
  If This.Singleton.AccessEchoScopes(index).Active Then
    This.Singleton.AccessEchoScopes(index).Active = False
    This.Singleton.AccessEchoCount = This.Singleton.AccessEchoCount - 1
  End If
  protUnregisterAccessEcho = This.Singleton.AccessEchoBaseline
  Dim i As Long
  For i = UBound(This.Singleton.AccessEchoScopes) To 1 Step -1
    If This.Singleton.AccessEchoScopes(i).Active Then
      protUnregisterAccessEcho = This.Singleton.AccessEchoScopes(i).EchoOn
      Exit Function
    End If
  Next
End Function


'Obtain a measurement
'@static
'@param sKey - Name of measurement to get
'@returns - Mean total block milliseconds across recorded runs, retaining fractional milliseconds.
'@example Get a measurement
'```vb
'Const C_MAX as Long = 1000
'With stdPerformance.CreateMeasure("Some test expression", C_MAX)
'  Dim v
'  For i = 1 to C_MAX
'    v = expressionToTest()
'  next
'End With
'
'Debug.Print "Average time for Some test expression: " & stdPerformance.Measurement("Some test expression") & " ms"
'```
Public Property Get Measurement(ByVal sKey As String) As Double
  If Me Is stdPerformance Then
    Dim v: v = getDictItem(sKey)
    If TypeName(v) = "Variant()" Then
        Measurement = getDictItem(sKey)(0)
    Else
        Measurement = Empty
    End If
  End If
End Property

'AddMeasurement
'@protected
'@param sKey   - Name of measurement to add to global cache
'@param time   - time to add to global cache
'@param nCount - number of operations (divisor)
'@remark If a time is added that was previously also added then the average of the times is calculated.
Public Sub AddMeasurement(ByVal sKey As String, ByVal time As Double, ByVal nCount As Double)
  If Me Is stdPerformance Then
    Debug.Print sKey & ": " & time & " ms" & IIf(nCount > 1, " (" & (1000 * time / nCount) & Chr(181) & "s per operation)", "")
    Dim ind As Long: ind = getDictIndex(sKey)
    If ind = -1 Then
      Call setDictItem(sKey, Array(time, 1))
    Else
      Dim vItem: vItem = getDictItem(sKey)
      Dim average As Double: average = vItem(0)
      Dim count As Long: count = vItem(1) + 1
      average = average + (time - average) / count
      Call setDictItem(sKey, Array(average, count))
    End If
  End If
End Sub

'MeasuresClear - Clears all procedures/blocks and times that have been measured
'@static
Public Sub MeasuresClear()
  ReDim This.Singleton.FakeDict(0 To 0)
End Sub

'MeasuresKeys
'@returns Array<String> - Array containing the procedures or blocks that have been measured.
Public Property Get MeasuresKeys() As String()
  If Me Is stdPerformance Then
    If UBound(This.Singleton.FakeDict) = 0 Then
      MeasuresKeys = Split("")
    Else
      'Define return array
      Dim sOut() As String
      ReDim Preserve sOut(0 To UBound(This.Singleton.FakeDict) - 1)

      'Fill keys array 
      Dim i As Long
      For i = 0 To UBound(This.Singleton.FakeDict) - 1
        sOut(i) = This.Singleton.FakeDict(i).key
      Next

      'return data
      MeasuresKeys = sOut
    End If
  End If
End Property

'Capture a fallback baseline as well, so API failure never mixes clock origins.
Private Sub startMeasure()
  This.Measure.StartTime = fallbackMilliseconds()
  On Error GoTo Unavailable
#If Mac Then
  #If MAC_OFFICE_VERSION >= 16 Then
    Static info As TMachTimebase
    If info.Denom = 0 Then
      If mach_timebase_info(info) <> 0 Then Exit Sub
    End If
    If info.Denom = 0 Then Exit Sub
    This.Measure.MillisecondsPerTick = CDbl(info.Numer) / CDbl(info.Denom) / 1000000#
    This.Measure.MachStart = mach_absolute_time()
    This.Measure.HighResolution = True
  #End If
#Else
  Static frequency As Currency
  If frequency = 0 Then
    If QueryPerformanceFrequency(frequency) = 0 Then Exit Sub
  End If
  If frequency <= 0 Then Exit Sub
  This.Measure.Frequency = frequency
  This.Measure.HighResolution = QueryPerformanceCounter(This.Measure.Counter) <> 0
#End If
Unavailable:
End Sub

'Subtract raw counter values before conversion to retain precision.
Private Function elapsedMilliseconds() As Double
  On Error GoTo Fallback
  If This.Measure.HighResolution Then
#If Mac Then
  #If MAC_OFFICE_VERSION >= 16 Then
    elapsedMilliseconds = CDbl(mach_absolute_time() - This.Measure.MachStart) * This.Measure.MillisecondsPerTick
    Exit Function
  #End If
#Else
    Dim counter As Currency
    If QueryPerformanceCounter(counter) <> 0 Then
      elapsedMilliseconds = CDbl(counter - This.Measure.Counter) / This.Measure.Frequency * 1000#
      Exit Function
    End If
#End If
  End If
Fallback:
  elapsedMilliseconds = fallbackMilliseconds() - This.Measure.StartTime
  If elapsedMilliseconds < 0 Then
#If Mac Then
    elapsedMilliseconds = elapsedMilliseconds + 86400000#
#Else
    elapsedMilliseconds = elapsedMilliseconds + 4294967296#
#End If
  End If
End Function

'Unsigned Windows ticks, or host-independent Mac seconds since midnight.
Private Function fallbackMilliseconds() As Double
#If Mac Then
  fallbackMilliseconds = CDbl(Timer) * 1000#
#Else
  fallbackMilliseconds = CDbl(GetTickCount())
  If fallbackMilliseconds < 0 Then fallbackMilliseconds = fallbackMilliseconds + 4294967296#
#End If
End Function

'Used by static class only
Private Sub Class_Initialize()
  If Me Is stdPerformance Then
    ReDim This.Singleton.FakeDict(0 To 0)
  End If
End Sub

'Used by instance objects only
'@destructor
Private Sub Class_Terminate()
  If Not Me Is stdPerformance Then
    Select Case This.ObjectType
      Case iMeasure
        Call stdPerformance.AddMeasurement(This.Measure.KeyName, elapsedMilliseconds(), This.Measure.Divider)
      Case iOptimiser
        'Attempt every captured restoration, including after partially failed initialisation.
        On Error Resume Next
        With This.Optimiser
          If .AccessEchoScope <> 0 Then Call .App.Echo(stdPerformance.protUnregisterAccessEcho(.AccessEchoScope))
          If .HasCalculation Then .App.Calculation = .Calculation
          If .HasEnableEvents Then .App.EnableEvents = .EnableEvents
          If .HasScreenUpdating Then .App.ScreenUpdating = .ScreenUpdating
        End With
    End Select
  End If
End Sub




'FakeDict Helpers
'==========================================================================================================================================
'NOTE: These functions are completely unoptimised and are largely in use for the purpose of making this class multi-platform friendly.
'These will be unlikely to be optimised given that this is largely a debugging library.

'Returns the index where a particular key is stored
'@param key - Key to find in dictionary
'@returns   - Index of key in dictionary
Private Function getDictIndex(ByVal key As String) As Long
  On Error GoTo ErrorOccurred
    Dim i As Long
    For i = 0 To UBound(This.Singleton.FakeDict)
      If This.Singleton.FakeDict(i).key = key Then
        getDictIndex = i
        Exit Function
      End If
    Next
  On Error GoTo 0
ErrorOccurred:
  getDictIndex = -1
End Function

'Set an item within a dictionary to a particular value
'@param key - Key to find in dictionary
'@param v   - Value to set dictionary too
'@param ind - Index of item. If not given `getDictIndex()` is used
Private Sub setDictItem(ByVal key As String, ByVal v As Variant, Optional ByVal ind As Long = -1)
  'get index of item in fake dict
  If ind = -1 Then ind = getDictIndex(key)
  
  'If item not exist, add it
  If ind = -1 Then
    ind = getUB(This.Singleton.FakeDict)
    This.Singleton.FakeDict(ind).key = key
    ReDim Preserve This.Singleton.FakeDict(0 To ind + 1)
  End If

  'Assign value to index
  If IsObject(v) Then
    Set This.Singleton.FakeDict(ind).val = v
  Else
    This.Singleton.FakeDict(ind).val = v
  End If
End Sub

'gets the upper bound of an array, if the array is uninitialised return -1
'@param items as Array<FakeDictItem> - Array of dict items.
'@returns - Upper bound of array OR -1 if not initialised
Private Function getUB(ByRef items() As FakeDictItem) As Long
    On Error GoTo ErrorOccurred
        getUB = UBound(items)
        Exit Function
ErrorOccurred:
    getUB = -1
End Function

'Returns the item paired with some key
'@param key - Key to find in dictionary
'@returns - Item stored at key
Private Function getDictItem(ByVal key As String) As Variant
  Dim ind As Long: ind = getDictIndex(key)
  If ind <> -1 Then
    If IsObject(This.Singleton.FakeDict(ind).val) Then
      Set getDictItem = This.Singleton.FakeDict(ind).val
    Else
      getDictItem = This.Singleton.FakeDict(ind).val
    End If
  Else
    getDictItem = Empty
  End If
End Function
