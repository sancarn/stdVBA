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

'TODO:
'* Full implementation of Pointer
'* Full implementation of Lambda (Will likely require stdVBAExpression library)
'* StdRun
'* CallCallback module function calling - is there a better way than application.run?

'OBJECT TESTING:
'   Sub main()
'     Dim cb As stdCallback
'     Set cb = stdCallback.Create("Object", Me, "f")
'
'     Dim x As Variant
'     Set x = cb()
'     Debug.Print x.Count
'
'     '---------------------------------------
'
'     Dim cbv As stdCallback
'     Set cbv = stdCallback.Create("Object", Me, "g")
'
'     Dim v As Variant
'     v = cbv()
'     Debug.Print v
'   End Sub
'
'   Function f() As Variant
'     Set f = New Collection
'   End Function
'   Function g() As Variant
'     g = 101
'   End Function
'
'MODULE TESTING:
'   Sub main()
'     Dim cb As stdCallback
'     Set cb = stdCallback.Create("Module", "MyModuleName", "f")
'
'     Dim x As Variant
'     Set x = cb()
'     Debug.Print x.Count
'
'     '---------------------------------------
'
'     Dim cbv As stdCallback
'     Set cbv = stdCallback.Create("Module", "MyModuleName", "g")
'
'     Dim v As Variant
'     v = cbv()
'     Debug.Print v
'   End Sub
'
'   Function f() As Variant
'     Set f = New Collection
'   End Function
'   Function g() As Variant
'     g = 101
'   End Function



'FYI
'Const DISPID_UNKNOWN     = -1
'Const DISPID_VALUE       = 0
'Const DISPID_PROPERTYPUT = -3
'Const DISPID_NEWENUM     = -4
'Const DISPID_EVALUATE    = -5
'Const DISPID_CONSTRUCTOR = -6
'Const DISPID_DESTRUCTOR  = -7
'Const DISPID_COLLECT     = -8

'TODO:
'* Implementation of Pointer
'* CreateLambda()
'* stdCallback.CreateFromAddress(AddressOf Main.Podrick)

'TODO: Lambda Expressions:
'******************
'Callback.[(a,b,c) => a + b * c].Call(1,2,3) ' ==> 7
'Callback.[obj => obj.id].Call(row)  ' ==> row.id


'Direct call convention of VBA.CallByName
#If VBA7 Then
  'VBE7 is interchangable with msvbvm60.dll    however VBE7.dll appears to always be present where as msvbvm60 is only occasionally present.
  Private Declare PtrSafe Function rtcCallByName Lib "VBE7.dll" (ByRef vRet As Variant, ByVal cObj As Object, ByVal sMethod As LongPtr, ByVal eCallType As VbCallType, ByRef pArgs() As Variant, ByVal lcid As Long) As Long
#Else
  Private Declare Function rtcCallByName Lib "msvbvm60" (ByRef vRet As Variant, ByVal cObj As Object, ByVal sMethod As LongPtr, ByVal eCallType As VbCallType, ByRef pArgs() As Variant, ByVal lcid As Long) As Long
#End If

Dim pInitialised As Boolean
' Dim pCallback As Collection
' Dim pCompiledLambda As Collection
' Dim pLambdaData As Collection

Private Enum ParentType
  Module = 1
  object = 2
  Formula = 3
  Bound = 4
End Enum
Private Type CallbackStruct
  iType As ParentType
  sModuleName As String
  oObject As Object
  sMethodName As String
  iCallType As VbCallType
  sFormula As String
  vBoundArgs As Variant
End Type
Private pCallback As CallbackStruct






Public Event BeforeRun(ByRef callback As stdCallback, ByRef args As Variant)
Public Event AfterRun(ByRef callback As stdCallback, ByRef sResult As Variant)



'Create a callback object from module type string, translate to enum and initialise callback
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
  Call Create.Init(iType, Parent, name, cType)
End Function

'Create callback from Module function directly
Public Function CreateFromModule(ByVal sParentName As String, ByVal sMethodName As String) As stdCallback
Attribute CreateFromModule.VB_Description = "Create callback from module method."
  Set CreateFromModule = New stdCallback
  Call CreateFromModule.Init(ParentType.Module, sParentName, sMethodName)
End Function

'Create callback from object and method name
Public Function CreateFromObjectMethod(ByRef object As Object, ByVal sMethodName As String) As stdCallback
Attribute CreateFromObjectMethod.VB_Description = "Create callback from object method."
  Set CreateFromObjectMethod = New stdCallback
  Call CreateFromObjectMethod.Init(ParentType.object, object, sMethodName, VbCallType.VbMethod)
End Function

'Create callback from object and property name
Public Function CreateFromObjectProperty(ByRef object As Object, ByVal sPropName As String, ByVal cType As VbCallType) As stdCallback
Attribute CreateFromObjectProperty.VB_Description = "Create callback from object property."
  Set CreateFromObjectProperty = New stdCallback
  Call CreateFromObjectProperty.Init(ParentType.object, object, sPropName, cType)
End Function

'CreateEvaluator from
Public Function CreateEvaluator(ByVal sFormula As String) As stdICallable
Attribute CreateEvaluator.VB_Description = "Create callback from Excel formula. This method is VBA Only"
  If IsObject(stdLambda) Then
    Set CreateEvaluator = stdLambda.Create(sFormula)
  Else
    CriticalRaise "Cannot find stdLambda library, required for Evaluator creation"
  End If
End Function

Friend Sub Init(iParentType As Long, ParamArray params() As Variant) 'Parent As Variant, Name As String
Attribute Init.VB_Description = "Initialises this object. This should not be called directly unless building a ctor routine e.g. Create()."
  If Not pInitialised Then
    Select Case iParentType
      Case ParentType.Module
        pCallback.iType = iParentType
        pCallback.sModuleName = params(0)
        pCallback.sMethodName = params(1)
      Case ParentType.object
        pCallback.iType = iParentType
        Set pCallback.oObject = params(0)
        pCallback.sMethodName = params(1)
        pCallback.iCallType = params(2)
      Case ParentType.Bound
        pCallback.iType = iParentType
        Set pCallback.oObject = params(0)
        If isArray(params(1)) Then
          pCallback.vBoundArgs = params(1)
        Else
          pCallback.vBoundArgs = Array()
        End If
      Case Else
        CriticalRaise "Invalid module type submitted"
    End Select
  Else
    CriticalRaise "Class is already initialised"
  End If
  vBoundArgs = Array()
  pInitialised = True
End Sub


Private Function stdICallable_Run(ParamArray params() As Variant) As Variant
  Call CopyVariant(stdICallable_Run, RunEx(params))
End Function
Private Function stdICallable_RunEx(ByVal params As Variant) As Variant
  Call CopyVariant(stdICallable_RunEx, RunEx(params))
End Function

'Bind a parameter to the function
Private Function stdICallable_Bind(ParamArray params() As Variant) As stdICallable
  Set stdICallable_Bind = BindEx(params)
End Function
Public Function Bind(ParamArray params() As Variant) As stdCallback
  Set Bind = BindEx(params)
End Function
Public Function BindEx(ByVal params As Variant) As stdCallback
  Set BindEx = New stdCallback
  Dim callable As stdICallable: Set callable = Me
  Call BindEx.Init(ParentType.Bound, callable, params)
End Function



Public Function Run(ParamArray params() As Variant)
Attribute Run.VB_Description = "Runs the specified callback."
Attribute Run.VB_UserMemId = 0
  Call CopyVariant(Run, RunEx(params))
End Function

'Default Property:
Public Function RunEx(ByVal vArr As Variant) As Variant
  If pInitialised And pCallback.iType Then
    If pCallback.iType = ParentType.Bound Then
      vArr = ConcatArrays(pCallback.vBoundArgs, vArr)
    End If

    'Raise event for listeners
    RaiseEvent BeforeRun(Me, vArr)
    
    'Select procedure based on parent type, as each requires different call conventions
    Dim vRet As Variant
    Select Case pCallback.iType
      Case ParentType.Module
        'TODO: Investigate whether we can call these directly via GetModuleHandleA(), GetProcAddress() and CallWindowProcW()
        'Get currentLength
        Dim currentLength As Integer
        currentLength = UBound(vArr) - LBound(vArr) + 1
        
        'Preserve array but make it 29 args long
        ReDim Preserve vArr(0 To 29)
        
        'Loop over args and bind missing to args not present in initial array
        Dim i As Integer
        For i = 0 To 29
          If i > currentLength - 1 Then
            vArr(i) = GetMissing
          End If
        Next
        
        'Copy results of Application.Run into vRet
        CopyVariant vRet, Application.Run( _
            pCallback.sModuleName & "." & pCallback.sMethodName, _
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
        
        If isArray(vArr) Then
          Dim vArgs() As Variant
          vArgs = vArr
            
          'Call rtcCallByName
          Dim hr As Long
          hr = rtcCallByName(vRet, pCallback.oObject, StrPtr(pCallback.sMethodName), pCallback.iCallType, vArgs, &H409)
            
          'If error then raise
          If hr < 0 Then
            CriticalRaise "Error in rtcCallByName. Error number: " & hr
          End If
        Else
          CriticalRaise "Error in rtcCallByName. Arguments supplied to RunEx needs to be an array."
        End If
      Case ParentType.Bound
        Dim callable As stdICallable
        Set callable = pCallback.oObject
        Call CopyVariant(vRet, callable.RunEx(vArr))
      Case Else
        CriticalRaise "Cannot call callback. Invalid parent type assigned (" & iParentType & ")."
      End Select

      'Allow events to modify result prior to submitting to return
      RaiseEvent AfterRun(Me, vRet)
      
      'Return result
      CopyVariant RunEx, vRet
  Else
    CriticalRaise "No callback specified. Please initialise with a Create function."
  End If
End Function





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
      Call stdError.Raise("Can only create evaluators in Excel as they rely on Application.Evaluate()")
    Else
      Call Err.Raise(1, "stdCallback", sMessage)
    End If
    End
End Sub

Private Function GetMissing(Optional arg As Variant) As Variant
  GetMissing = arg
End Function


'Copies one variant to a destination
'@param {ByRef Variant} dest Destination to copy variant to
'@param {Variant} value Source to copy variant from.
'@perf This appears to be a faster variant of "oleaut32.dll\VariantCopy" + it's multi-platform
Private Sub CopyVariant(ByRef dest As Variant, ByVal value As Variant)
  If IsObject(value) Then
    Set dest = value
  Else
    dest = value
  End If
End Sub


'Used by Bind() for binding arguments ontop of BoundArgs and binding bound args to passed arguments
'@param {Variant()} The 1st array which will
'@param {Variant()} The 2nd array which will be concatenated after the 1st
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
