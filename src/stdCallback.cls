VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Implements stdICallable

'@module
'@description A class which implements a callback object. Use this class to create a 
'callback object which can be called at a later stage. These callbacks are used extensively
'across the stdVBA library.
'
'Within a class we can create a simple callback object like so:
'```vb
'Dim cb as stdCallback: set cb = stdCallback.CreateFromObjectMethod(Me, "myFunction")
'```
'With an example function declaration as follows:
'```vb
'Public Function myFunction(ByVal x as long, ByVal y as long, ByVal z as long) as long
'  myFunction = x + y + z
'End Function
'```
'
'> Note it's important to note that the function must be public.
'
'The callback can then be called as follows:
'```vb
'Debug.Print cb(1,2,3)   '=> 6
'```
'Or with an array of parameters:
'```vb
'Debug.Print cb.Run(Array(1,2,3))   '=> 6
'```
'You can also bind arguments to the callback. This is useful when you want to prepare a callback with some
'arguments before final execution. For example:
'```vb
'Dim cb as stdCallback: set cb = stdCallback.CreateFromObjectMethod(Me, "myFunction").Bind(1,2,3)
''...later...
'Debug.Print cb() '=> 6
'```
'@example ```vb
''A simple class based callback:
'Public Function myFunction(ByVal x as long, ByVal y as long, ByVal z as long) as long
'  myFunction = x + y + z
'End Function
'Public Sub mySub()
'  Dim cb as stdCallback: set cb = stdCallback.CreateFromObjectMethod(Me, "myFunction")
'  Debug.Print cb(1,2,3)   '=> 6
'End Sub
'```
'@example ```vb
''A simple module based callback. The module in this example is named "MyModule":
'Public Function myFunction(ByVal x as long, ByVal y as long, ByVal z as long) as long
'  myFunction = x + y + z
'End Function
'Public Sub mySub()
'  Dim cb as stdCallback: set cb = stdCallback.CreateFromModule("MyModule", "myFunction")
'  Debug.Print cb(1,2,3)   '=> 6
'End Sub
'```
'@example ```vb
''Create a callback from a module in another workbook:
'Public Sub mySub()
'  Dim cb as stdCallback: set cb = stdCallback.CreateFromWorkbookModule( _
'    "C:\Temp\workbooks\MyWorkbook.xlsm", "MyModule", "myFunction")
'  Debug.Print cb(1,2,3)   '=> 6
'End Sub
'```
'@example ```vb
''Create a callback from a pointer to a function in a public module.
''Note: this is the only instance where the function can be private!
'Private Function myFunction(ByVal x as long, ByVal y as long, ByVal z as long) as long
'  myFunction = x + y + z
'End Function
'Public Sub mySub()
'  Dim cb as stdCallback: set cb = stdCallback.CreateFromPointer(AddressOf myFunction, vbLong, Array(vbLong, vbLong, vbLong))
'  Debug.Print cb(1,2,3)   '=> 6
'End Sub
'```


'TODO:
'* Full implementation of Pointer
'* CallCallback module function calling - is there a better way than application.run?





'FYI
'Const DISPID_UNKNOWN     = -1
'Const DISPID_VALUE       = 0
'Const DISPID_PROPERTYPUT = -3
'Const DISPID_NEWENUM     = -4
'Const DISPID_EVALUATE    = -5
'Const DISPID_CONSTRUCTOR = -6
'Const DISPID_DESTRUCTOR  = -7
'Const DISPID_COLLECT     = -8

'Direct call convention of VBA.CallByName
#If Mac Then
#Else
  #If VBA7 Then
    'VBE7 is interchangable with msvbvm60.dll    however VBE7.dll appears to always be present where as msvbvm60 is only occasionally present.
    Private Declare PtrSafe Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As LongPtr, ByVal offsetinVft As LongPtr, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByVal paTypes As LongPtr, ByVal paValues As LongPtr, ByRef retVAR As Variant) As Long
    Private Declare PtrSafe Function rtcCallByName Lib "VBE7.dll" (ByRef vRet As Variant, ByVal cObj As Object, ByVal sMethod As LongPtr, ByVal eCallType As VbCallType, ByRef pArgs() As Variant, ByVal lcid As Long) As Long
  #Else
    Private Enum LongPtr
      [_]
    End Enum
    Private Declare Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByVal paTypes As LongPtr, ByVal paValues As LongPtr, ByRef retVAR as variant) As Long
    Private Declare Function rtcCallByName Lib "msvbvm60" (ByRef vRet As Variant, ByVal cObj As Object, ByVal sMethod As LongPtr, ByVal eCallType As VbCallType, ByRef pArgs() As Variant, ByVal lcid As Long) As Long
  #End If
#End If

#If Mac Then
  Const vbLongLong As Long = 8
#End If


#If Win64 Then
  Private Const vbLongPtr = vbLongLong
#Else
  Private Const vbLongPtr = vbLong
#End If

Private Enum ParentType
  Module = 1
  object = 2
  Formula = 3
  Bound = 4
  Pointer = 5
End Enum

Private Type TCBByPointer
  iMethodAddress As LongPtr
  bPredictTypes As Boolean
  iReturnType As VbVarType
  iParamTypes() As VbVarType
End Type


Private Type CallbackStruct
  iType As ParentType
  sModuleName As String
  oObject As Object
  sMethodName As String
  iCallType As VbCallType
  sFormula As String
  vBoundArgs As Variant
  CBByPointer As TCBByPointer
End Type
Private Type TThis
  Initialised As Boolean
  Callback As CallbackStruct
End Type
Private This As TThis



'Event raised before the callback is run
'@param callback - The callback object that is about to be run
'@param args     - The arguments that will be passed to the callback. This can be modified by listeners.
Public Event BeforeRun(ByRef callback As stdCallback, ByRef args As Variant)

'Event raised after the callback is run
'@param callback - The callback object that has just been run
'@param result  - The result of the callback. This can be modified by listeners.
Public Event AfterRun(ByRef callback As stdCallback, ByRef result As Variant)



'Create a callback object from module type string, translate to enum and initialise callback
'@constructor
'@param mType                            - The type of module to create the callback from. Can be either "Module" or "Object"
'@param Parent as Variant<String|Object> - The parent object to create the callback from. If mType is "Module" then this should be a string, otherwise it should be an object.
'@param name                             - The name of the method to create the callback from.
'@param cType                            - The call type to use when calling the method. Defaults to VbCallType.VbMethod
'@returns                                - A 1st class callback function object
Public Function Create(mType As String, Parent As Variant, name As String, Optional cType As Integer = VbCallType.VbMethod) As stdCallback
Attribute Create.VB_Description = "Creates an instance of this class."
  Dim iType As ParentType
  If mType = "Module" Then
    iType = ParentType.Module
  ElseIf mType = "Object" Then
    iType = ParentType.object
  Else
    CriticalRaise "Incorrect module type supplied"
  End If
  
  Set Create = New stdCallback
  Call Create.protInit(iType, Parent, name, cType)
End Function

'Create callback from Module function directly
'@constructor
'@param sParentName - The name of the module to create the callback from.
'@param sMethodName - The name of the method to create the callback from.
'@returns           - A 1st class callback function object
Public Function CreateFromModule(ByVal sParentName As String, ByVal sMethodName As String) As stdCallback
Attribute CreateFromModule.VB_Description = "Create callback from module method."
  Set CreateFromModule = New stdCallback
  Call CreateFromModule.protInit(ParentType.Module, sParentName, sMethodName)
End Function

'Create callback from Workbook module function directly
'@constructor
'@param sWorkbookPath - The path of the workbook to create the callback from.
'@param sModuleName   - The name of the module to create the callback from.
'@param sMethodName   - The name of the method to create the callback from.
'@returns             - A 1st class callback function object
Public Function CreateFromWorkbookModule(ByVal sWorkbookPath As String, ByVal sModuleName As String, ByVal sMethodName As String) As stdCallback
Attribute CreateFromWorkbookModule.VB_Description = "Create callback from extenrnal workbook module method."
  Set CreateFromWorkbookModule = New stdCallback
  Call CreateFromWorkbookModule.protInit(ParentType.Module, "'" & sWorkbookPath & "'!" & sModuleName, sMethodName)
End Function

'Create callback from object and method name
'@constructor
'@param object      - The object to create the callback from.
'@param sMethodName - The name of the method to create the callback from.
'@returns           - A 1st class callback function object
Public Function CreateFromObjectMethod(ByRef object As Object, ByVal sMethodName As String) As stdCallback
Attribute CreateFromObjectMethod.VB_Description = "Create callback from object method."
  Set CreateFromObjectMethod = New stdCallback
  Call CreateFromObjectMethod.protInit(ParentType.object, object, sMethodName, VbCallType.vbMethod)
End Function

'Create callback from object and property name
'@constructor
'@param object      - The object to create the callback from.
'@param sPropName   - The name of the property to create the callback from.
'@param cType       - The call type to use when calling the method. Defaults to VbCallType.VbGet
'@returns           - A 1st class callback function object
Public Function CreateFromObjectProperty(ByRef object As Object, ByVal sPropName As String, ByVal cType As VbCallType) As stdCallback
Attribute CreateFromObjectProperty.VB_Description = "Create callback from object property."
  Set CreateFromObjectProperty = New stdCallback
  Call CreateFromObjectProperty.protInit(ParentType.object, object, sPropName, cType)
End Function

'Create callback from function pointer
'@constructor
'@param MethodPointer                     - The address of the function to create the callback from.
'@param iRetType                          - The return type of the function to create the callback from.
'@param vParamTypes as Array<VbVarType>   - The parameter types of the function to create the callback from.
'@returns                                 - A 1st class callback function object
#If VBA7 Then
  Public Function CreateFromPointer(ByVal MethodPointer As LongPtr, Optional ByVal iRetType As VbVarType = 0, Optional vParamTypes = Null) As stdCallback
#Else
  Public Function CreateFromPointer(ByVal MethodPointer As Long, Optional ByVal iRetType As VbVarType = 0, Optional vParamTypes = Null) As stdCallback
#End If
  Set CreateFromPointer = New stdCallback
  Call CreateFromPointer.protInit(ParentType.Pointer, MethodPointer, iRetType, vParamTypes)
End Function

'Initialise the callback object
'@protected
'@param iParentType - The type of parent to create the callback from. Can be either ParentType.Module or ParentType.object
'@param params      - The parameters to initialise the callback with. This should not be called directly unless building a ctor routine e.g. Create().
Friend Sub protInit(iParentType As Long, ParamArray params() As Variant)
Attribute protInit.VB_Description = "Initialises this object. This should not be called directly unless building a ctor routine e.g. Create()."
  If Not This.Initialised Then
    Select Case iParentType
      Case ParentType.Module
        This.Callback.iType = iParentType
        This.Callback.sModuleName = params(0)
        This.Callback.sMethodName = params(1)
      Case ParentType.object
        This.Callback.iType = iParentType
        Set This.Callback.oObject = params(0)
        This.Callback.sMethodName = params(1)
        This.Callback.iCallType = params(2)
      Case ParentType.Bound
        This.Callback.iType = iParentType
        Set This.Callback.oObject = params(0)
        If IsArray(params(1)) Then
          This.Callback.vBoundArgs = params(1)
        Else
          This.Callback.vBoundArgs = Array()
        End If
      Case ParentType.Pointer
        This.Callback.iType = iParentType
        With This.Callback.CBByPointer
          .iMethodAddress = params(0)
          .iReturnType = params(1)
          .bPredictTypes = IsNull(params(2))
          If Not .bPredictTypes Then
            ReDim .iParamTypes(0 To UBound(params(2)))
            Dim i As Long
            For i = 0 To UBound(params(2))
              .iParamTypes(i) = params(2)(i)
            Next
          End If
        End With
      Case Else
        CriticalRaise "Invalid module type submitted"
    End Select
  Else
    CriticalRaise "Class is already initialised"
  End If
  This.Initialised = True
End Sub

'Run the function
'@param params - The parameters to run the function with
'@returns      - The result of the function
Private Function stdICallable_Run(ParamArray params() As Variant) As Variant
  Call CopyVariant(stdICallable_Run, RunEx(params))
End Function

'Run the function with a array of parameters
'@param params as Variant<Array<Variant>> - The parameters to run the function with
'@returns                                 - The result of the function
Private Function stdICallable_RunEx(ByVal params As Variant) As Variant
  Call CopyVariant(stdICallable_RunEx, RunEx(params))
End Function

'Bind a parameter to the function
'@param params - The parameters to bind to the function
'@returns      - A new callback with the parameters bound
Private Function stdICallable_Bind(ParamArray params() As Variant) As stdICallable
  Set stdICallable_Bind = BindEx(params)
End Function

'Bind a parameter to the function. Bound parameters are bound to the left of the passed parameters.
'@param params as Array<Variant> - The parameters to bind to the function
'@returns                        - A new callback with the parameters bound
Public Function Bind(ParamArray params() As Variant) As stdCallback
  Set Bind = BindEx(params)
End Function

'Bind a parameter to the function. Bound parameters are bound to the left of the passed parameters.
'@param params as Variant<Array<Variant>> - The parameters to bind to the function
'@returns                        - A new callback with the parameters bound
Public Function BindEx(ByVal params As Variant) As stdCallback
  Set BindEx = New stdCallback
  Dim callable As stdICallable: Set callable = Me
  Call BindEx.protInit(ParentType.Bound, callable, params)
End Function

'Get information about the callback
'@param sMessage as "obj"|"className" - The message to send to the callback. "obj" returns the callback object, "className" returns the name of the class.
'@param success as Boolean            - Whether the message was successfully sent
'@param params as Variant             - The parameters to send to the callback
'@returns                             - The result of the message
Private Function stdICallable_SendMessage(ByVal sMessage As String, ByRef success As Boolean, ByVal params As Variant) As Variant
    Select Case sMessage
        Case "obj"
            Set stdICallable_SendMessage = Me
            success = True
        Case "className"
            stdICallable_SendMessage = "stdCallback"
            success = True
        Case Else
            success = False
    End Select
End Function

'Run the function
'@param params - The parameters to run the function with
'@returns      - The result of the function
Public Function Run(ParamArray params() As Variant)
Attribute Run.VB_Description = "Runs the specified callback."
Attribute Run.VB_UserMemId = 0
  Call CopyVariant(Run, RunEx(params))
End Function

'Run the function with a array of parameters
'@param vArr as Variant<Array<Variant>> - The parameters to run the function with
'@returns                                 - The result of the function
Public Function RunEx(ByVal vArr As Variant) As Variant
  If This.Initialised And This.Callback.iType Then
    If This.Callback.iType = ParentType.Bound Then
      vArr = ConcatArrays(This.Callback.vBoundArgs, vArr)
    End If

    'Raise event for listeners
    RaiseEvent BeforeRun(Me, vArr)
    
    'Select procedure based on parent type, as each requires different call conventions
    Dim vRet As Variant
    Select Case This.Callback.iType
      Case ParentType.Module
        'TODO: Use fancy accessor mechanism https://codereview.stackexchange.com/questions/274532/low-level-vba-hacking-making-private-functions-public 
        'Get currentLength
        Dim currentLength As Integer
        currentLength = UBound(vArr) - LBound(vArr) + 1
        
        'Preserve array but make it 29 args long
        ReDim Preserve vArr(0 To 29)
        
        'Get missing value
        static vMissing As Variant: if IsEmpty(vMissing) then vMissing = GetMissing()

        'Loop over args and bind missing to args not present in initial array
        Dim i As Integer
        For i = 0 To 29
          If i > currentLength - 1 Then
            vArr(i) = vMissing
          End If
        Next
        
        'Copy results of Application.Run into vRet
        CopyVariant vRet, Application.Run( _
            This.Callback.sModuleName & "." & This.Callback.sMethodName, _
            vArr(0), vArr(1), _
            vArr(2), vArr(3), _
            vArr(4), vArr(5), _
            vArr(6), vArr(7), _
            vArr(8), vArr(9), _
            vArr(10), vArr(11), _
            vArr(12), vArr(13), _
            vArr(14), vArr(15), _
            vArr(16), vArr(17), _
            vArr(18), vArr(19), _
            vArr(20), vArr(21), _
            vArr(22), vArr(23), _
            vArr(24), vArr(25), _
            vArr(26), vArr(27), _
            vArr(28), vArr(29))
      Case ParentType.object
        'Call function directly
        'Use rtcCallByName to avoid type errors from argument array
        'Src: http://www.vbforums.com/showthread.php?866039-RESOLVED-Store-CallByName-Arguments-list
        
        If IsArray(vArr) Then
          Dim vArgs() As Variant
          vArgs = vArr
            
          'Call rtcCallByName
          Dim hResult As Long
          On Error GoTo ErrorInRTCCallByName
          #If Mac Then
            Call macCallByName(vRet, This.Callback.oObject, This.Callback.sMethodName, This.Callback.iCallType, vArgs)
          #Else
            hResult = rtcCallByName(vRet, This.Callback.oObject, StrPtr(This.Callback.sMethodName), This.Callback.iCallType, vArgs, &H409)
          #End If
          On Error GoTo stdErrorWrapper_ErrorOccurred
        Else
          CriticalRaise "Error in rtcCallByName. Arguments supplied to RunEx needs to be an array."
        End If
      Case ParentType.Bound
        Dim callable As stdICallable
        Set callable = This.Callback.oObject
        Call CopyVariant(vRet, callable.RunEx(vArr))
      Case ParentType.Pointer
        'Create array of types and pointers to vars, for function call
        Dim vPtr() As LongPtr
        Dim vType() As Integer
        
        'Create array from params
        Dim vParam() As Variant: vParam = vArr
        Dim iParamCount As Long: iParamCount = UBound(vParam) - LBound(vParam) + 1
        
        'Initialise vType based on whether it should be predicted or not
        Dim lIdx As Long
        With This.Callback.CBByPointer
          If .bPredictTypes Then
            If iParamCount = 0 Then
              ReDim vType(0 To 0)
            Else
              ReDim vType(0 To iParamCount - 1)
              For lIdx = 0 To iParamCount - 1
                vType(lIdx) = VarType(vParam(lIdx))
              Next
            End If
          Else
            If (UBound(.iParamTypes) - LBound(.iParamTypes) + 1) <> iParamCount Then
              Err.Raise 1, "", "Contradicting number of variables passed into callback compared to initialisation"
            End If
            
            ReDim vType(0 To iParamCount - 1)
            For lIdx = 0 To iParamCount - 1
              vType(lIdx) = .iParamTypes(lIdx)
            Next
          End If
                
          'Populate array of pointers to params, and array of types for DispCallFunc
          If iParamCount = 0 Then
            ReDim vPtr(0 To 0)
          Else
            ReDim vPtr(0 To iParamCount - 1)
            For lIdx = 0 To iParamCount - 1
              vPtr(lIdx) = VarPtr(vParam(lIdx))
            Next
          End If
          
          'Call function
          Const CC_STDCALL = 4
          Dim hResult2 As Long
          #If Mac Then
            Call Err.Raise(1, "stdCallback", "Calling functions by pointer is not supported on Mac")
          #Else
            hResult2 = DispCallFunc(0, .iMethodAddress, CC_STDCALL, .iReturnType, iParamCount, VarPtr(vType(0)), VarPtr(vPtr(0)), vRet)
          #End If
        End With
        
        If hResult2 < 0 Then
            Err.Raise hResult2, "DispCallFunc", "Error when calling function pointer?"
        End If

      Case Else
        CriticalRaise "Cannot call callback. Invalid parent type assigned (" & This.Callback.iType & ")."
      End Select

      'Allow events to modify result prior to submitting to return
      RaiseEvent AfterRun(Me, vRet)
      
      'Return result
      CopyVariant RunEx, vRet
  Else
    CriticalRaise "No callback specified. Please initialise with a Create function."
  End If
  Exit Function
ErrorInRTCCallByName:
  CriticalRaise Err.Description & " (" & Err.Number & ")"
End Function

#If Mac Then
  'Call a method on an object using the mac version of CallByName
  '@param o - The object to call the method on
  '@param p - The name of the method to call
  '@param v - The call type to use
  Private Sub macCallByName(ByRef vRet As Variant, ByVal o As Object, ByVal p As String, ByVal v As VbCallType, ByRef args() As Variant)
      Select Case UBound(args) - LBound(args) + 1
        Case 0: Call CopyVariant(vRet, CallByName(o, p, v))
        Case 1: Call CopyVariant(vRet, CallByName(o, p, v, args(0)))
        Case 2: Call CopyVariant(vRet, CallByName(o, p, v, args(0), args(1)))
        Case 3: Call CopyVariant(vRet, CallByName(o, p, v, args(0), args(1), args(2)))
        Case 4: Call CopyVariant(vRet, CallByName(o, p, v, args(0), args(1), args(2), args(3)))
        Case 5: Call CopyVariant(vRet, CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4)))
        Case 6: Call CopyVariant(vRet, CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5)))
        Case 7: Call CopyVariant(vRet, CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6)))
        Case 8: Call CopyVariant(vRet, CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7)))
        Case 9: Call CopyVariant(vRet, CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8)))
        Case 10: Call CopyVariant(vRet, CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9)))
        Case 11: Call CopyVariant(vRet, CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10)))
        Case 12: Call CopyVariant(vRet, CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11)))
        Case 13: Call CopyVariant(vRet, CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12)))
        Case 14: Call CopyVariant(vRet, CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13)))
        Case 15: Call CopyVariant(vRet, CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14)))
        Case 16: Call CopyVariant(vRet, CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15)))
        Case 17: Call CopyVariant(vRet, CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16)))
        Case 18: Call CopyVariant(vRet, CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17)))
        Case 19: Call CopyVariant(vRet, CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18)))
        Case 20: Call CopyVariant(vRet, CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19)))
        Case 21: Call CopyVariant(vRet, CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20)))
        Case 22: Call CopyVariant(vRet, CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21)))
        Case 23: Call CopyVariant(vRet, CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22)))
        Case 24: Call CopyVariant(vRet, CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23)))
        Case 25: Call CopyVariant(vRet, CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24)))
        Case 26: Call CopyVariant(vRet, CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25)))
        Case 27: Call CopyVariant(vRet, CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25), args(26)))
        Case 28: Call CopyVariant(vRet, CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25), args(26), args(27)))
        Case 29: Call CopyVariant(vRet, CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25), args(26), args(27), args(28)))
        Case Else: Call CopyVariant(vRet, CallByName(o, p, v, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25), args(26), args(27), args(28), args(29)))
      End Select
  End Sub
#End If

Private Function Serialize(data As Variant) As String
  Select Case TypeName(data)
    Case "Integer", "Double", "Float", "Date"
      Serialize = CDbl(data)
    Case "String"
      Serialize = """" & data & """"
    Case Else
      If VarType(data) = vbObject Then
        'allows for stuff like --> CreateEvaluator("@(@(@($1,""Workbooks""),""Item"",1),""Name"")")(Application)
        'Assuming @() calls
        Serialize = ObjPtr(data)
      Else
        Serialize = data
      End If
  End Select
End Function

Private Sub CriticalRaise(ByVal sMessage As String)
  'If stdError exists
    If VarType(stdError) Then
      Call stdError.Raise(sMessage)
    Else
      Call Err.Raise(1, "stdCallback", sMessage)
    End If
    End
End Sub

Private Function GetMissing(Optional arg As Variant) As Variant
  GetMissing = arg
End Function


'Copies one variant to a destination
'@param dest  - Destination to copy variant to
'@param value - Source to copy variant from.
'@perf This appears to be a faster variant of "oleaut32.dll\VariantCopy" + it's multi-platform
Private Sub CopyVariant(ByRef dest As Variant, ByVal value As Variant)
  If IsObject(value) Then
    Set dest = value
  Else
    dest = value
  End If
End Sub


'Used by Bind() for binding arguments ontop of BoundArgs and binding bound args to passed arguments
'@param Arr1 as Variant() - The 1st array which will
'@param Arr2 as Variant() - The 2nd array which will be concatenated after the 1st
'@returns                 - The concatenated array
'@complexity O(1)
Private Function ConcatArrays(ByVal Arr1 As Variant, ByVal Arr2 As Variant) As Variant
    Dim ub1 As Long: ub1 = UBound(Arr1)
    Dim lb1 As Long: lb1 = LBound(Arr1)
    Dim ub2 As Long: ub2 = UBound(Arr2)
    Dim lb2 As Long: lb2 = LBound(Arr2)
    Dim iub As Long: iub = ub1 + ub2 - lb2 + 1
    
    If iub > -1 Then
        Dim v() As Variant
        ReDim v(lb1 To iub)
        
        
        Dim i As Long
        For i = LBound(v) To UBound(v)
            If i <= ub1 Then
                Call CopyVariant(v(i), Arr1(i))
            Else
                Call CopyVariant(v(i), Arr2(i - ub1 - 1 + lb2))
            End If
        Next
        ConcatArrays = v
    Else
        ConcatArrays = Array()
    End If
End Function
