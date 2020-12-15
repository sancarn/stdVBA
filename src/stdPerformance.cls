VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdPerformance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


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


Public Enum EPerfObjType
  iMeasure=1
  iOptimiser=2
End Enum

Private pObjType as EPerfObjType


'iOptimiser Fields...
Private pEnableEvents as boolean
Private pScreenUpdating as boolean


'iMeasure definitions
Private pStartTime as long
Private pKeyName as string 


'Public storage
Private Type FakeDictItem
  key as string
  val as variant
End Type
Private FakeDict() as FakeDictItem

'Create
'@constructor
'@param {EPerfObjType} - Type of performance object to create. iMeasure - used for measuring performance, iOptimiser - used for optimising performance 
'@param {Variant()} - Additional params supplied as array.
'@returns {stdPerformance<EPerfObjType>} - Object termination has special behaviour. See Measure and Optimise methods for further details.
Public Function Create(ByVal objType as EPerfObjType, ByVal params as Variant) as stdPerformance
  Set Create = new stdPerformance
  Call Create.init(objType, params)
End Function

'
Public Sub Init(ByVal objType as EPerfObjType, ByVal params as variant)
  pObjType = objType
  select case objType
    case iMeasure
      pKeyName = params(0)
      pStartTime = GetTickCount()
    case iOptimiser
      'Store vals
      pScreenUpdating = Application.ScreenUpdating
      pEnableEvents = Application.EnableEvents
      
      'Set vals
      Application.ScreenUpdating = params(0)
      Application.EnableEvents = params(1)
  end select
End Sub

'Measure
'@constructor
'@param {String} - Name of method or block to measure
'@returns {stdPerformance<iMeasure>} - Object which upon termination, adds measurement of block to global cache
'@usage
'  ```vb
'  With stdPerformance.Measure("Hello world")
'    For i = 1 to 1000
'      Debug.print "Hello world"
'    next
'  End With
'  ```
Public Function Measure(ByVal sProc as string) as stdPerformance
  set Measure = stdPerformance.Create(iMeasure, Array(sProc))
End Function

'Optimise
'@constructor
'@param {Boolean} - ScreenUpdating set value
'@param {Boolean} - EnableEvents set value
'@returns {stdPerformance<iOptimiser>} - Object termination has special behaviour. See Measure and Optimise methods for further details.
'@usage
'  ```vb
'  With stdPerformance.Optimise
'    'some heavy code here
'  End With
'  ```
Public Function Optimise(Optional ByVal ScreenUpdating as boolean = false, Optional ByVal EnableEvents as boolean = false) as stdPerformance
  set Optimise = stdPerformance.Create(iOptimiser, Array(ScreenUpdating,EnableEvents))
End Function


'GetMeasurement
'@param {String} - Name of measurement to get
'@returns {Double} - Average measurement time
Public Function GetMeasurement(ByVal sKey As String) As Double
  If Me Is stdPerformance Then
    Dim v: v = getDictItem(sKey)
    If TypeName(v) = "Variant()" Then
        GetMeasurement = getDictItem(sKey)(0)
    Else
        GetMeasurement = Empty
    End If
  End If
End Function

'AddMeasurement
'If a time is added that was previously also added then the average of the times is calculated.
'@param {String} - Name of measurement to add to global cache
'@param {Long} - time to add to global cache
Public Sub AddMeasurement(ByVal sKey as string, ByVal time as long)
  if Me is stdPerformance then
    Debug.Print sKey & ": " & time & " ms"
    Dim ind as long: ind = getDictIndex(sKey)
    if ind = -1 then
      Call setDictItem(sKey, Array(time,1))
    else
      Dim vItem: vItem = getDictItem(sKey)
      Dim average as long: average = vItem(0)
      Dim count as long: count = vItem(1) + 1
      average = average + (time - average)/count
      Call setDictItem(sKey, Array(average,count))
    end if
  end if
End Sub

Public Sub MeasurementsClear()
  ReDim FakeDict(0 to 0)
End Sub

Public Property Get MeasurementKeys() as string()
  if Me is stdPerformance then
    if ubound(FakeDict) = 0 then
      MeasurementKeys = Split("")
    else
      'Define return array
      Dim sOut() as string
      Redim Preserve sOut(0 to ubound(FakeDict)-1)

      'Fill keys array 
      Dim i as long
      For i = 0 to ubound(FakeDict)-1
        sOut(i) = FakeDict(i).key
      next

      'return data
      MeasurementKeys = sOut
    end if
  end if 
End Property


Private Sub Class_Initialize()
  if me is stdPerformance then
    Redim FakeDict(0 to 0)
  end if
End Sub

'@destructor
Private Sub Class_Terminate()
  select case pObjType
    case iMeasure
      Dim pEndTime as long: pEndTime = GetTickCount()
      Call stdPerformance.AddMeasurement(pKeyName, Abs(pEndTime - pStartTime))
    case iOptimiser
      'Store vals
      Application.ScreenUpdating = pScreenUpdating
      Application.EnableEvents = pEnableEvents
  end select
End Sub




'FakeDict Helpers
'==========================================================================================================================================
'NOTE: These functions are completely unoptimised and are largely in use for the purpose of making this class multi-platform friendly.
'These will be unlikely to be optimised given that this is largely a debugging library.

'getDictIndex
'Returns the index where a particular key is stored
'@param {string} - Key to find in dictionary
'@returns {long} = Index of key in dictionary
Private Function getDictIndex(ByVal key as string) as Long
  On Error GoTo ErrorOccurred
    Dim i as long
    For i = 0 to ubound(FakeDict)
      if FakeDict(i).key = key then
        getDictIndex = i
        Exit Function
      end if
    next
  On Error Goto 0
ErrorOccurred:
  getDictIndex = -1
End Function

'setDictItem
'Set an item within a dictionary to a particular value
'@param {string} - Key to find in dictionary
'@param {variant} - Value to set dictionary too
'@param {optional long} - Index of item. If not given getDictIndex() is used
Private Sub setDictItem(ByVal key as string, ByVal v as variant, Optional ByVal ind as long = -1)
  'get index of item in fake dict
  if ind = -1 then ind = getDictIndex(key)
  
  'If item not exist, add it
  if ind = -1 then
    ind = getUB(FakeDict)
    FakeDict(ind).key = key
    Redim Preserve FakeDict(0 to ind+1)
  end if

  'Assign value to index
  if isObject(v) then
    set FakeDict(ind).val = v
  else
    FakeDict(ind).val = v
  end if
End Sub

'getUB
'gets the upper bound of an array, if the array is uninitialised return -1
'@param {ByRef FakeDictItem()} Array of dict items
'@returns {Long} - Upper bound of array OR -1 if not initialised
Private Function getUB(ByRef items() As FakeDictItem) As Long
    On Error GoTo ErrorOccurred
        getUB = UBound(items)
        Exit Function
ErrorOccurred:
    getUB = -1
End Function

'getDictIndex
'Returns the item paired with some key
'@param {string} - Key to find in dictionary
'@returns {variant} = Item stored at key
Private Function getDictItem(ByVal key as string) as variant
  Dim ind as Long: ind = getDictIndex(key)
  if ind <> -1 then 
    if isObject(FakeDict(ind).val) then
      set getDictItem = FakeDict(ind).val
    else
      getDictItem = FakeDict(ind).val
    end if
  else
    getDictItem = Empty
  end if
End Function
