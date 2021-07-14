VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdCOMIDispatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = true

Private pCOM as stdCOM
Private pRIID as GUID   'rIID must be IID_NULL == GUID.Empty

'https://docs.microsoft.com/en-us/windows/win32/winprog/windows-data-types
'Win32   | VBA Equiv | Size         | Signed?
'--------|-----------|--------------|-----------
' WORD   | Integer   | 2            | No
' DWORD  | Long      | 4            | No
' HFILE  | Long      | 4            | Yes
' UINT   | Long      | 4            | No
' HANDLE | LongPtr   | x32:4/x64:8  |
'



'These DISPIDs should be Invoked using IDispatch#Invoke() e.g.
'myCom.QueryKnownInterface(IDispatch).Invoke(DISPID_NEWENUM, ...)
'https://docs.microsoft.com/en-us/previous-versions/windows/desktop/automat/dispid-constants
Public Enum stdCOMDispIDs
  stdCOMDispIDValue = 0
  stdCOMDispIDUnknown = -1
  stdCOMDispIDPropertyPut = -3
  stdCOMDispIDNewEnum = -4
  stdCOMDispIDEvaluate = -5
  stdCOMDispIDConstructor = -6
  stdCOMDispIDDestructor = -7
  stdCOMDispIDCollect = -8
End Enum

'@documentation https://docs.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-dispparams
'Contains the arguments passed to a method or property.
'Example:
' Function DoSomething(a as long, optional b as string, optional c as object)
'  DoSomething(1,"2",myObj) --> {rgVArg:[1,"2",myObj],rgDispIDNamedArgs:[], cArgs:3, cNamedArgs:0}
'  DoSomething(a:=1,"a") --> Invalid syntax, named arguments must always follow positional arguments
'  DoSomething(1, b:="a") --> Valid syntax: {rgVArg: [1,"a"], gDispIDNamedArgs:[2],cArgs:2, cNamedArgs:1}
'  DoSomething(1, c:=myObj) --> Valid syntax: {rgVArg: [1,myObj], gDispIDNamedArgs:[3], cArgs:2, cNamedArgs:1}
'I.E. cNamedArgs represents arguments to the right that have a name
Private Type DISPPARAMS
  rgVArg as Variant()           'Array of arguments (as variant)
  rgDispIDNamedArgs as Long()   'DispIDs of named args. Named arguments must always follow positional arguments.
  cArgs as long                 'Count of arguments
  cNamedArgs as long            'Count of named arguments
End Type

'Contains error information
'@documentation https://docs.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-excepinfo
Private Type EXCEPINFO
  wCode as Integer
  wReserved as Integer
  bstrSource as string
  bstrDescription as string
  bstrHelpFile as string
  dwHelpContext as long
  #if VBA7 then
    pvReserved as longptr
    pfnDeferredFillIn as longptr
  #else
    pvReserved as long
    pfnDeferredFillIn as long
  #end if
  statusCode as long               'https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/7d93a429-b652-4c77-9b63-d7a762521e4d
End Type

'GUID struct for QueryInterface
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type


'Create a IDispatch wrapper 
'@constructor
Public Function Create(ByVal com as stdCOM) as stdCOMIDispatch
  set Create = new stdCOMIDispatch
  Call Create.protInit(com)
End Function


'invoke a dispid of a function
'@param {Long}           DispID   - DISPID of function to call
'@param {VbCallType}     callType - Whether this method is a method call, property getter, letter or setter.
'@param {Variant()}      params   - Parameters in
'@param {Variant()=null} nameLocations - locations of named parameters
'@param {Long=en}        lcid - LocaleID. By default `9` == English
'@returns {variant}      Return value of invoked function
'@throws
Public Function dispInvoke(ByVal DispId as long, ByVal callType as VbCallType, ByVal params as variant, optional ByVal nameLocations as variant = null, optional ByVal lcid as long = 9) as Variant
  Dim vResult as variant
  if isNull(nameLocations) then nameLocations = Array()
  
  'Create dispparams structure
  Dim dpParams as DISPPARAMS
  dpParams.rgVArg = params
  dpParams.rgDispIDNamedArgs = nameLocations
  dpParams.cArgs = UBound(params)-lbound(params)+1
  dpParams.cNamedArgs = Ubound(nameLocations)-LBound(nameLocations)+1
  Dim excep as EXCEPINFO 'Create exception to be populated

  Dim result as long
  result = pCOM.CallVT(EIDispatch.Invoke,vbLong, DispID, varptr(pRIID.data1), lcid, callType, VarPtr(dpParams), VarPtr(vResult), VarPtr(excep))

  if result = 0 then
    Call CopyVariant(dispInvoke, vResult)
  else
    Err.Raise excep.wCode, excep.bstrSource, excep.bstrDescription
  end if
End Function

'Obtain dispids of names
'@param {ByVal Variant()} names - Names to find IDs of
'@param {ByVal Long=9}    lcid  - Locale (language) id
'@returns Array of dispids
Public Function dispGetIDsOfNames(ByVal names as Variant, optional byval lcid as long = 9) as long()
  Dim iUB as long: iUB = ubound(names)
  Dim iLB as long: iLB = lbound(names)
  
  Dim sNames() as string: Redim sNames(iLB to iUB)
  Dim ids() as long:      Redim ids(iLB to iUB)

  Dim i as long
  For i = iLB to iUB
    sNames(i) = names(i)
  next
  
  Dim result as long
  result = pCOM.CallVT(EIDispatch.GetIDsOfNames, vbLong, varptr(pRIID.data1), varptr(sNames(iLB)), iUB-iLB+1, lcid, varptr(ids(iLB))

  if result = 0 then
    dispGetIDsOfNames = ids
  else
    Err.Raise 1, "", "..."
  end if
End Function

Public Property Get dispGetIEnumVARIANT() as stdCOM
  On Error GoTo CannotQuery
    Dim disp as stdCOM
    Set disp = Me.QueryKnownInterface(IDispatch)
    
    Dim oIEnumVariant as IUnknown
    set oIEnumVariant = dispInvoke(stdCOMDispIDNewEnum,VbGet | VbMethod)

    Exit Property
CannotQuery:
End Property



Friend Sub protInit(ByVal com as stdCOM)
  set pCOM = com.QueryKnownInterface(IDispatch)
End Sub