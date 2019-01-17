VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "STD_Types_Callback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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





Dim pInitialised As Boolean
Dim pCallback As Collection
Dim pCompiledLambda as Collection
Dim pLambdaData as Collection

'Token for Lambda
Private Type TToken
  pType as String
  pValue as String
End Type


Property Set Callback(val As Collection)
Attribute Callback.VB_Description = "Defines the specified callback."
  If pInitialised Then
    Err.Raise 0, "STD_Types_Callback::ctor", "Cannot initialise initialised Callback"
  Else
    pCallback = val
  End If
End Property

'Default Property:
Property Get Run(ParamArray args() As Variant) As Variant
Attribute Run.VB_Description = "Runs the specified callback."
Attribute Run.VB_UserMemId = 0
  If pInitialised And Not pCallback Is Nothing Then
    Dim vRet As Variant
    vRet = CallCallback(args)
    If VarType(vRet) = "Object" Then
      Set Run = vRet
    Else
      Run = vRet
    End If
  Else
    Err.Raise 0, "STD_Types_Callback::Run", "No callback specified. Please initialise with Create()."
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
    Err.Raise 0, "STD_Types_Callback::Run", "No callback specified. Please initialise with Create()."
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

Friend Sub Init()
Attribute Init.VB_Description = "Initialises this object. This should not be called directly unless building a ctor routine e.g. Create()."
  pInitialised = True
End Sub

'Create a callback object:
Public Function Create(mType As String, Parent As Variant, Name As String, Optional cType As Integer = VbCallType.VbMethod) As STD_Types_Callback
Attribute Create.VB_Description = "Creates an instance of this class."
  Set Create = New STD_Types_Callback
  Create.Callback = New Collection
  Create.Callback.Add mType, "Type"
  Create.Callback.Add Parent, "Parent"
  Create.Callback.Add Name, "Name"
  Create.Callback.Add cType, "cType"
  Create.Init
End Function

'Examples:
'******************
'Callback.[(a,b,c) => a + b * c].Call(1,2,3) ' ==> 7
'Callback.[obj => obj.id].Call(row)  ' ==> row.id
'...
Public Function Lambda(ByVal lambda as string)
Attribute Lambda.VB_UserMemId = -5
  set Lambda = New STD_Types_Callback
  Lambda.Callback = new Collection
  Lambda.Callback.Add "Lambda", "Type"
  Dim sLambda() as String
  sLambda = Split(lambda,"=>",1)
  Lambda.Callback.Add sLambda(0), "args"
  Lambda.Callback.Add sLambda(1), "lambda"
End Function

'CALLBACK STRUCTURE:
'Type   = Module|Object
'Parent = ModuleName or ParentObject
'Name   = Name of method to call
'cType  = Call type (for Objects)
Private Function CallCallback(args() As Variant) As Variant
  Dim vRet As Variant
  If Callback("Type") = "Module" Then
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
    CallCallback = Application.Run( _
      Callback("Parent") & "." * Callback("Name"), _
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
    CallCallback = VBA.CallByName(Callback("Parent"), Callback("Name"), Callback("cType"), args)
  Else
    Err.Raise 0, "STD.Types.Callback::Call", "Cannot call callback. Invalid type """ & Callback("Type") & """assigned."
  End If
End Function

Private Function GetMissing(Optional arg As Variant) As Variant
  GetMissing = arg
End Function






Private Function TokeniseLambda(ByRef Tokens() as TToken)
  Dim Lambda() as Byte
  Lambda = Callback("lambda")
  
  Redim Tokens(0 to UBound(Lambda) - LBound(Lambda))
  
  Dim TokenHash(255) as Boolean, i as integer
  For i = 32 to 47: TokenHash(i)=True: Next
  For i = 58 to 64: TokenHash(i)=True: Next
  For i = 91 to 96: TokenHash(i)=True: Next
  For i = 123 to 126: TokenHash(i)=True: Next
  
  Dim s as string, sToken as string, ch as integer
  Dim token as TToken
  For i = 0 to UBound(Lambda) step 2
    ch = Lambda(i+1)
    if ch=10 then ch=58 'Assume vbLf = ":"
    'Do not pass any disallow all control characters
    if ch < 32 then Continue For
    if TokenHash(ch)
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
    else
      'Concatenate string literal
      s = s & Asc(ch)
    end if
  Next
  
  'Add final token
  if s <> "" then
    token.pType = "Literal"
    token.pValue = s
    TokeniseLambda.add token
  end if
End Function
