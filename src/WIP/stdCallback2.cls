VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdCallback2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Design:
' EByApplicationRun
'  On Windows:
'    At initialise:
'      QueryInterface(Application, IDispatch)           [By DispCallFunc]
'      dispid = IDispatch::GetIDsOfNames("Run")         [By DispCallFunc]
'    At runtime:
'      args = args.unshift(MacroName)
'      IDispatch::Invoke(This.ApplicationRun.dispid)    [By DispCallFunc]
'  On Mac:
'    Application.Run macro, ...                         [see stdCallback1]
' EByObjectMember
'   On Windows:
'     At initialise:
'       QueryInterface(parent, IDispatch)               [By DispCallFunc]
'       dispid = IDispatch::GetIDsOfNames(runName)      [By DispCallFunc]
'     At runtime:
'       IDispatch::Invoke(This.dispid)                  [By DispCallFunc]
'   On Mac
'     macCallByName(obj, funcName, callerType, args)    [see stdLambda]
' EByPointer
'  On Windows:
'    At runtime:
'      DispCallFunc(ptr, 0, ...)
'  On Mac:
'    At initialise:
'      Throw "Not implemented"


Implements stdICallable

Private Const CC_STDCALL    As Long = 4

#If VBA7 Then
  'VBE7 is interchangable with msvbvm60.dll    however VBE7.dll appears to always be present where as msvbvm60 is only occasionally present.
  Private Declare PtrSafe Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As LongPtr, ByVal offsetinVft As LongPtr, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByVal paTypes As LongPtr, ByVal paValues As LongPtr, ByRef retVAR As Variant) As Long
#Else
  Private Enum LongPtr
    [_]
  End Enum
  Private Declare Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As LongPtr, ByVal offsetinVft As LongPtr, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByVal paTypes As LongPtr, ByVal paValues As LongPtr, ByRef retVAR as variant) As Long
#End If

Public Enum ECallbackType
  EByApplicationRun=1
  EByObjectMember=2
  EByPointer=3
End Enum
Private Enum ECallbackType2
  EByClone=4
End Enum

'[IDispatch](https://en.wikipedia.org/wiki/IDispatch)  extends IUnknown
'0      HRESULT  QueryInterface ([in] REFIID riid, [out] void **ppvObject)
'1      ULONG    AddRef ()
'2      ULONG    Release ()
'3      HRESULT  GetTypeInfoCount(unsigned int * pctinfo)
'4      HRESULT  GetTypeInfo(unsigned int iTInfo, LCID lcid, ITypeInfo ** ppTInfo)
'5      HRESULT  GetIDsOfNames(REFIID riid, OLECHAR ** rgszNames, unsigned int cNames, LCID lcid, DISPID * rgDispId)
'6      HRESULT  Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS * pDispParams, VARIANT * pVarResult, EXCEPINFO * pExcepInfo, unsigned int * puArgErr)
Public Enum EIDispatch
  QueryInterface
  AddRef
  Release
  GetTypeInfoCount
  GetTypeInfo
  GetIDsOfNames
  Invoke
End Enum



'BUILDTIME - At point of `Create` execution
'RUNTIME   - At point of `Run` execution

Private Type dispCallFuncInfo
  Instance as LongPtr      'BUILDTIME | Either ptr to function, or objptr to IDispatch
  VTableOffset as LongPtr  'BUILDTIME | Either 0 or IDispatch::Invoke offset
  CallConv as Long         'BUILDTIME | stdcall
  ReturnType as VbVarType  'BUILDTIME | Either returnType or vbVariant
  aTypes() as Long         'BUILDTIME | Pointer to array of parameter types
  aValues() as Variant     'RUNTIME   | Pointer to array of values as Variant
End Type

Private Type dispInfo
  obj as Object            'BUILDTIME | Object which pointer points to
  id as long               'BUILDTIME | DISPID
  invokeFlag as vbCallType 'BUILDTIME | Whether it's a method or property etc.
  args as LongPtr          'RUNTIME   | pointer to Variant() array of passed in args
  argsCount as Long        'RUNTIME   | count of args passed to function and present in args
End Type

Private Type TThis
  ctype as ECallbackType
  decl as dispCallFuncInfo
  disp as dispInfo
End Type
Private This as TThis


Public Function CreateFromRunCommand(ByVal sRunCommand as string) as stdCallback2
  set CreateFromRunCommand = new stdCallback2
  Call CreateFromRunCommand.protInit(EByApplicationRun, sRunCommand)
End Function

Public Function CreateFromObjectMember(Byval o as object, ByVal sMemberName as string, ByVal cType as vbCallType) as stdCallback2
  set CreateFromRunCommand = new stdCallback2
  Call CreateFromRunCommand.protInit(EByObjectMember, o, sMemberName, cType)
End Function

Public Function CreateFromObjectOffset(ByVal iObjPtr as LongPtr, ByVal iVTableOffset as LongPtr, ByVal iRetType as vbVarType, ByRef vParamTypes() as vbVarType, Optional ByVal callConvention as Long = CC_STDCALL) as stdCallback2
  set CreateFromObjectOffset = new stdCallback2
  Call CreateFromObjectOffset.protInit(EByPointer, iObjPtr, iVTableOffset, iRetType, vParamTypes, callConvention)
End Function

Public Function Clone() as stdCallback2
  Dim oClone as Collection: set oClone = new Collection
  Call oClone.add(This.ctype, "This.ctype")
  Call oClone.add(This.decl.Instance, "This.decl.Instance")
  Call oClone.add(This.decl.VTableOffset, "This.decl.VTableOffset")
  Call oClone.add(This.decl.CallConv, "This.decl.CallConv")
  Call oClone.add(This.decl.ReturnType, "This.decl.ReturnType")
  Call oClone.add(This.decl.aTypes, "This.decl.aTypes")
  Call oClone.add(This.decl.aValues, "This.decl.aValues")
  Call oClone.add(This.disp.args, "This.disp.args")
  Call oClone.add(This.disp.argsCount, "This.disp.argsCount")
  Call oClone.add(this.disp.id, "this.disp.id")
  Call oClone.add(this.disp.obj, "this.disp.obj")

  set Clone = new stdCallback2
  Call Clone.protInit(EByClone, oClone)
End Function

Public Function Bind(ParamArray vParams()) as stdCallback

End Function

'Create callback from a module member.
'@param {string} Name of macro/function to execute
'@param {string=""} Name of module containing macro. Optional. If not provided, any module will be targeted.
'@param {Workbook=Nothing} Workbook of module to execute macro in. Optional. If not provided, active workbook / any loaded addins are targeted.
Public Function CreateFromModuleMember(ByVal sMemberName as string, Optional ByVal sModuleName as string = "", Optional ByVal wb as Workbook = nothing) as stdCallback2
  Dim sPath as string
  if not wb is nothing then
    sPath = wb.fullname
  else
    sPath = ""
  end if

  'Build a run command of the form '{Workbook}'!{Module}.{Macro}
  Dim sRunCommand as string: sRunCommand = iif(sPath<>"","'","") &  sPath & iif(sPath<>"","'!","") & sModuleName & iif(sModuleName <> "", ".", "") & sMemberName
  set CreateFromModuleMember = CreateFromRunCommand(sRunCommand)
End Function

'Create callback from a pointer to an existing function.
'@param {LongPtr}     Pointer to function to execute. Use `AddressOf` to obtain this value.
'@param {vbVarType}   Return type of function.
'@param {vbVarType()} Array of parameter types
'@returns {stdCallback2} The callback itself
'@remark Uses `CreateFromObjectOffset` with `iObjPtr:=0`
Public Function CreateFromPointer(ByVal ptr as LongPtr, ByVal iRetType as vbVarType, ByRef vParamTypes() as vbVarType, Optional ByVal callConvention as Long = CC_STDCALL) as stdCallback2
  set CreateFromPointer = CreateFromObjectOffset(0, ptr, iRetType, vParamTypes, callConvention)
End Function

Private Enum ParamEnumeration
  byObjectMember_object = 0
  byObjectMember_MemberName = 1
  byObjectMember_CallType = 2
  byApplicationRun_Command = 0
  byPointer_Instance = 0
  byPointer_VTableOffset = 1
  byPointer_RetType = 2
  byPointer_ParamTypes = 3
  byPointer_CallConv = 4
End Enum

Friend Sub protInit(callbackType As ECallbackType, ParamArray params() As Variant)
  this.ctype = callbackType
  select case callbackType
    Case EByObjectMember
      'Byval o as object, ByVal sMemberName as string, ByVal cType as vbCallType
      set this.disp.obj      = unkQueryInterface(params(byObjectMember_object), "{00020400-0000-0000-C000-000000000046}")
      this.disp.id           = getDispID(this.disp.obj, params(byObjectMember_MemberName))
      this.disp.invokeFlag   = params(byObjectMember_CallType)
      this.decl.Instance     = objptr(this.disp.obj)
      this.decl.VTableOffset = EIDispatch.Invoke
      this.decl.CallConv     = CC_STDCALL
      this.decl.ReturnType   = vbLong
      'Signiature of IDispatch::Invoke()
      this.decl.aTypes       = getTypeArray(vbLong, vbLong, vbLong, vbLong, vbLongPtr, vbLongPtr, vbLongPtr, vbLongPtr)
    Case EByPointer
      'ByVal iObjPtr as LongPtr, ByVal iVTableOffset as LongPtr, ByVal iRetType as vbVarType, ByRef vParamTypes() as vbVarType
      this.decl.Instance     = params(byPointer_Instance)
      this.decl.VTableOffset = params(byPointer_VTableOffset)
      this.decl.CallConv     = params(byPointer_CallConv)
      this.decl.ReturnType   = params(byPointer_RetType)
      this.decl.ParamTypes   = params(byPointer_ParamTypes)
    Case EByApplicationRun
      Const EApplication_Run = 50
      set this.disp.obj      = unkQueryInterface(Application, "{00020400-0000-0000-C000-000000000046}")
      this.disp.id           = getDispID(this.disp.obj, "Run")
      this.disp.invokeFlag   = vbMethod
      this.decl.Instance     = objptr(this.disp.obj)
      this.decl.VTableOffset = EIDispatch.Invoke
      this.decl.CallConv     = CC_STDCALL
      this.decl.ReturnType   = vbLong
      'Signiature of IDispatch::Invoke()
      this.decl.aTypes       = getTypeArray(vbLong, vbLong, vbLong, vbLong, vbLongPtr, vbLongPtr, vbLongPtr, vbLongPtr)
      this.decl.aValues      = Array(params(byApplicationRun_Command))
    Case EByClone
      With params(0)
        This.ctype             = .item("This.ctype")
        This.decl.Instance     = .item("This.decl.Instance")
        This.decl.VTableOffset = .item("This.decl.VTableOffset")
        This.decl.CallConv     = .item("This.decl.CallConv")
        This.decl.ReturnType   = .item("This.decl.ReturnType")
        This.decl.aTypes       = .item("This.decl.aTypes")
        This.decl.aValues      = .item("This.decl.aValues")
        This.disp.args         = .item("This.disp.args")
        This.disp.argsCount    = .item("This.disp.argsCount")
        this.disp.id           = .item("this.disp.id")
        set this.disp.obj      = .item("this.disp.obj")
      end with
    case EByExisting
  end select
End Sub

Public Property Get CallbackType() as ECallbackType
  CallbackType = This.ctype
End Property






Public Function Run(ParamArray params() As Variant)
Attribute Run.VB_Description = "Runs the specified callback."
Attribute Run.VB_UserMemId = 0
  Call ExecuteCallback(params, Run)
End Function
Public Function RunEx(ByVal vArr As Variant) As Variant
  Call ExecuteCallback(params, RunEx)
End Function

Private Function stdICallable_Run(ParamArray params() As Variant) As Variant
  Call ExecuteCallback(params, stdICallable_Run)
End Function
Private Function stdICallable_RunEx(ByVal params As Variant) As Variant
  Call ExecuteCallback(params, stdICallable_RunEx)
End Function


Private Sub ExecuteCallback(ByRef params as variant, ByRef returnValue as Variant)

End Sub