VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'TODO:
'* Full implementation of Pointer


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

#If VBA6 Then
  Private Declare PtrSafe Function rtcCallByName Lib "msvbvm60" (ByRef vRet As Variant, ByVal cObj As Object, ByVal sMethod As LongPtr, ByVal eCallType As VbCallType, ByRef pArgs() As Variant, ByVal lcid As Long) As Long
  Private Declare PtrSafe Function GetMem4 Lib "msvbvm60.dll" (ByVal Addr As Long, ByRef RetVal As Any) As Long
 
#Else
  Private Declare Function rtcCallByName Lib "msvbvm60" (ByRef vRet As Variant, ByVal cObj As Object, ByVal sMethod As LongPtr, ByVal eCallType As VbCallType, ByRef pArgs() As Variant, ByVal lcid As Long) As Long
  Private Declare Function GetMem4 Lib "msvbvm60.dll" (ByVal Addr As Long, ByRef RetVal As Any) As Long
#End If

Dim pInitialised As Boolean
Dim pCallback As Collection
Dim pCompiledLambda As Collection
Dim pLambdaData As Collection

'Token for Lambda
Private Type TToken
  pType As String
  pValue As String
End Type


Public Property Get Callback() As Collection
  Set Callback = pCallback
End Property

Public Property Set Callback(val As Collection)
Attribute Callback.VB_Description = "Defines the specified callback."
  If pInitialised Then
    Err.Raise 0, "stdCallback::ctor", "Cannot initialise initialised Callback"
  Else
    Set pCallback = val
  End If
End Property

'Default Property:
Public Property Get Run(ParamArray args() As Variant) As Variant
Attribute Run.VB_Description = "Runs the specified callback."
Attribute Run.VB_UserMemId = 0
  If pInitialised And Not pCallback Is Nothing Then
    Dim vvArgs() As Variant
    vvArgs = args
    
    Dim vRet As Variant
    vRet = CallCallback(vvArgs)
    If VarType(vRet) = vbObject Then
      Set Run = vRet
    Else
      Run = vRet
    End If
  Else
    Err.Raise 0, "stdCallback::Run", "No callback specified. Please initialise with Create()."
  End If
End Property

Private Function StdRun(ParamArray args() As Variant) As Variant
Attribute StdRun.VB_Description = "Run method for Win32APIs."
  If pInitialised And Not pCallback Is Nothing Then
    Dim vRet As Variant
    vRet = CallCallback(args)
    If VarType(vRet) = "Object" Then
      Set StdRun = vRet
    Else
      StdRun = vRet
    End If
  Else
    Err.Raise 0, "stdCallback::Run", "No callback specified. Please initialise with Create()."
  End If
End Function

'0 = StdCall
'1 = Get call
Property Get Pointer(Optional cType As Integer = 0) As LongPtr
Attribute Pointer.VB_Description = "Gets the pointer to the Run function property of the current Callback for usage in Win32 API calls."
  Dim MePtr As LongPtr
  'MePtr = STD.Automation.COM.Deref(Me)
  If cType = 0 Then
    'Navigate to VB_UserMemId 99
  Else
    'Navigate to VB_UserMemId 0
  End If
End Property

Friend Sub init()
Attribute init.VB_Description = "Initialises this object. This should not be called directly unless building a ctor routine e.g. Create()."
  pInitialised = True
End Sub

'Create a callback object:
Public Function Create(mType As String, Parent As Variant, Name As String, Optional cType As Integer = VbCallType.VbMethod) As stdCallback
Attribute Create.VB_Description = "Creates an instance of this class."
  Set Create = New stdCallback
  Set Create.Callback = New Collection
  Create.Callback.add mType, "Type"
  Create.Callback.add Parent, "Parent"
  Create.Callback.add Name, "Name"
  Create.Callback.add cType, "cType"
  Create.init
End Function

'Examples:
'******************
'Callback.[(a,b,c) => a + b * c].Call(1,2,3) ' ==> 7
'Callback.[obj => obj.id].Call(row)  ' ==> row.id
'...
Public Function Lambda(ByVal Lambda As String)
Attribute Lambda.VB_UserMemId = -5
  Set Lambda = New stdCallback
  Lambda.Callback = New Collection
  Lambda.Callback.add "Lambda", "Type"
  Dim sLambda() As String
  sLambda = Split(Lambda, "=>", 1)
  Lambda.Callback.add sLambda(0), "args"
  Lambda.Callback.add sLambda(1), "lambda"
End Function

'CALLBACK STRUCTURE:
'Type   = Module|Object
'Parent = ModuleName or ParentObject
'Name   = Name of method to call
'cType  = Call type (for Objects)
Private Function CallCallback(args() As Variant) As Variant
  Dim vRet As Variant
  If Callback("Type") = "Module" Then
    'TODO: Investigate whether we can call these directly via GetModuleHandleA(), GetProcAddress() and CallWindowProcW()
    'Get currentLength
    Dim currentLength As Integer
    currentLength = UBound(args) - LBound(args) + 1
    
    'Preserve array but make it 29 args long
    ReDim Preserve args(0 To 29)
    
    'Loop over args and bind missing to args not present in initial array
    Dim i As Integer
    For i = 0 To 29
      If i > currentLength - 1 Then
        args(i) = GetMissing
      End If
    Next
    
    'Call callback on module function
    'Dim vp As LongPtr
    'vp = VarPtr(Application.Run(Callback("Parent") & "." & Callback("Name"), args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25), args(26), args(27), args(28), args(29)))
    'Call GetMem4(vp, vRet)
    
    CallCallback = Application.Run( _
      Callback("Parent") & "." & Callback("Name"), _
      args(0), args(1), _
      args(2), args(3), _
      args(4), args(5), _
      args(6), args(7), _
      args(8), args(9), _
      args(10), args(11), _
      args(12), args(13), _
      args(14), args(15), _
      args(16), args(17), _
      args(18), args(19), _
      args(20), args(21), _
      args(22), args(23), _
      args(24), args(25), _
      args(26), args(27), _
      args(28), args(29))
  ElseIf Callback("Type") = "Object" Then
    'Call function directly
    'Use rtcCallByName to avoid type errors from argument array
    'Src: http://www.vbforums.com/showthread.php?866039-RESOLVED-Store-CallByName-Arguments-list
    
    Dim vArr() As Variant
    vArr = args
    
    Dim hr As Long
    hr = rtcCallByName(CallCallback, Callback("Parent"), StrPtr(Callback("Name")), Callback("cType"), vArr, &H409)
    
    'If error then raise
    If hr < 0 Then
      Err.Raise hr, "stdCallback::CallCallback"
    End If
  Else
    Err.Raise 0, "STD.Types.Callback::Call", "Cannot call callback. Invalid type """ & Callback("Type") & """assigned."
  End If
End Function

Private Function GetMissing(Optional arg As Variant) As Variant
  GetMissing = arg
End Function






Private Function TokeniseLambda(ByRef Tokens() As TToken)
  Dim Lambda() As Byte
  Lambda = Callback("lambda")
  
  ReDim Tokens(0 To UBound(Lambda) - LBound(Lambda))
  
  Dim TokenHash(255) As Boolean, i As Integer
  For i = 32 To 47: TokenHash(i) = True: Next
  For i = 58 To 64: TokenHash(i) = True: Next
  For i = 91 To 96: TokenHash(i) = True: Next
  For i = 123 To 126: TokenHash(i) = True: Next
  
  Dim s As String, sToken As String, ch As Integer
  Dim token As TToken
  For i = 0 To UBound(Lambda) Step 2
    ch = Lambda(i + 1)
    If ch = 10 Then ch = 58 'Assume vbLf = ":"
    'Do not pass any disallow all control characters
    'TODO: if ch < 32 then Continue For
    If TokenHash(ch) Then
      'Previous token
      token.pType = "Literal"
      token.pValue = s
      TokeniseLambda.add token
      
      'Current Token
      token.pType = "Symbol"
      token.pValue = Asc(ch)
      TokeniseLambda.add token
      
      'Next token
      s = ""
    Else
      'Concatenate string literal
      s = s & Asc(ch)
    End If
  Next
  
  'Add final token
  If s <> "" Then
    token.pType = "Literal"
    token.pValue = s
    TokeniseLambda.add token
  End If
End Function
