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
'With stdPerformance.Measure("Some test expression", C_MAX)
'  Dim v
'  For i = 1 to C_MAX
'    v = expressionToTest()
'  next
'End With
'```
'@example Simple example of optimising performance:
'```vb
'With stdPerformance.Optimise
'  'some heavy code here
'End With
'```



'Spec:
'This class has been designed to meet your performance testing and optimisation needs. stdPerformance uses the `Sentry Object` design pattern
'which allows for cleaner more maintainable code.
'Functions implemented on the class
'CONSTRUCTORS
'    [X] Create   - With Cache
'    [X] init       #PROTECTED
'    [X] Measure  - Create a performance measuring Sentry Object
'    [X] Optimise - Create a object which toggles runtime options for optimisation. Currently sets: `ScreenUpdating`, `EnableEvents` and `XLCalculation`.
'                   This is intended to be application agnostic.
'
'STATIC PROPERTIES
'    [X] Get MeasureKeys - Get an array of procs/blocks which have been measured
'    [x] Get Measurement(sProcOrBlock) - Get the average time it took to execute a block.
'    [ ] Get MeasuresStr
'    [ ] Get MeasuresHtml
'
'STATIC methods
'    [x] MeasuresClear() - Clear the performance stack.
'
'OUT-OF-SCOPE
'    * Anything performance related which is specific, should realistically be honed to a specific class for that thing.
'
'EXAMPLES
'# 1 - Usage of Optimser
'
'   'Disable numerous options for performance
'   Debug.Print Application.ScreenUpdating, Application.EnableEvents, Application.Calculation         'false,false,-4105 
'   With stdPerformance.Optimiser()
'     Debug.Print Application.ScreenUpdating, Application.EnableEvents, Application.Calculation         'false,false,-4135
'   End With
'   Debug.Print Application.ScreenUpdating, Application.EnableEvents, Application.Calculation         'false,false,-4105 
'   
'   'Disable everything BUT Calculation
'   Debug.Print Application.ScreenUpdating, Application.EnableEvents, Application.Calculation         'false,false,-4105 
'   With stdPerformance.Optimiser(Calculation:=xlCalculation.xlCalculationAutomatic)
'     Debug.Print Application.ScreenUpdating, Application.EnableEvents, Application.Calculation         'false,false,-4105
'   End With
'   Debug.Print Application.ScreenUpdating, Application.EnableEvents, Application.Calculation         'false,false,-4105 
'
'# 2 - measuring performance:
'
'   With stdPerformance.measure("#1 Select then set")
'     For i = 1 to C_MAX
'       cells(1,1).select
'       selection.value = "hello"
'     Next
'   End With
'   
'   With stdPerformance.measure("#2 Set directly")
'     For i = 1 to C_MAX
'       cells(1,1).value = "hello"
'     next
'   End With
'
'Declares for performance counters
#If Mac Then
   #If MAC_OFFICE_VERSION >= 15 Then
      Private Declare Function GetTickCount Lib "/Applications/Microsoft Excel.app/Contents/Frameworks/MicrosoftOffice.framework/MicrosoftOffice" () As Long
   #Else
      Private Declare Function GetTickCount Lib "Applications:Microsoft Office 2011:Office:MicrosoftOffice.framework:MicrosoftOffice" () As Long
   #End If
#Else ' Win32 or Win64
   #If VBA7 Then
      Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
   #Else
      Private Declare Function GetTickCount Lib "kernel32" () As Long
   #End If
#End If

'Enum for sentry object type
Public Enum EPerfObjType
  iMeasure = 1
  iOptimiser = 2
End Enum


'For Performance Optimisation sentries
Private Type TOptimiser
  EnableEvents As Boolean
  ScreenUpdating As Boolean
  Calculation As Long
End Type

'For Performance Measurement
Private Type TMeasure
  StartTime As Long
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
      This.Measure.StartTime = GetTickCount()
    Case iOptimiser
      'TODO: Need to make this more application agnostic
      Dim oApp As Object: Set oApp = Application
      
      'Store vals
      This.Optimiser.ScreenUpdating = oApp.ScreenUpdating
      This.Optimiser.EnableEvents = oApp.EnableEvents
      
      'Set vals
      oApp.ScreenUpdating = params(0)
      oApp.EnableEvents = params(1)
      
      'Different options for different applications
      Select Case oApp.Name
        Case "Microsoft Excel"
          This.Optimiser.Calculation = oApp.Calculation
          oApp.Calculation = params(2)
      End Select
  End Select
End Sub

'Create a measure object
'@constructor
'@param sProc   - Name of method or block to measure
'@param nCount  - Number of times the block will be ran (used to calculate average time), default = 1.
'@returns stdPerformance<iMeasure> - Object which upon termination, adds measurement of block to global cache
'@example ```vb
'With stdPerformance.Measure("Some test expression")
'  Dim v
'  For i = 1 to 1000
'    v = expressionToTest()
'  next
'End With
'```
'@example
'```vb
'Const C_MAX as Long = 1000
'With stdPerformance.Measure("Some test expression", C_MAX)
'  Dim v
'  For i = 1 to C_MAX
'    v = expressionToTest()
'  next
'End With
'```
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
'With stdPerformance.Optimise
'  'some heavy code here
'End With
'```
Public Function CreateOptimiser(Optional ByVal ScreenUpdating As Boolean = False, Optional ByVal EnableEvents As Boolean = False, Optional ByVal Calculation As Long = -4135) As stdPerformance
  Set CreateOptimiser = Create(iOptimiser, Array(ScreenUpdating, EnableEvents, Calculation))
End Function


'Obtain a measurement
'@param sKey - Name of measurement to get
'@returns - Average measurement time
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
      Dim average As Long: average = vItem(0)
      Dim count As Long: count = vItem(1) + 1
      average = average + (time - average) / count
      Call setDictItem(sKey, Array(average, count))
    End If
  End If
End Sub

'MeasuresClear - Clears all procedures/blocks and times that have been measured
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

'Used by static class only
'@constructor
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
        Dim pEndTime As Long: pEndTime = GetTickCount()
        Call stdPerformance.AddMeasurement(This.Measure.KeyName, Abs(pEndTime - This.Measure.StartTime), This.Measure.Divider)
      Case iOptimiser
        'TODO: Make this more application agnostic
        
        Dim oApp As Object: Set oApp = Application
        'Store vals
        oApp.ScreenUpdating = This.Optimiser.ScreenUpdating
        oApp.EnableEvents = This.Optimiser.EnableEvents
        
        'Different options for different applications
        Select Case oApp.Name
          Case "Microsoft Excel"
            oApp.Calculation = This.Optimiser.Calculation
        End Select
    End Select
  End If
End Sub




'FakeDict Helpers
'==========================================================================================================================================
'NOTE: These functions are completely unoptimised and are largely in use for the purpose of making this class multi-platform friendly.
'These will be unlikely to be optimised given that this is largely a debugging library.

'getDictIndex
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

'setDictItem
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

'getUB
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

'getDictIndex
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
