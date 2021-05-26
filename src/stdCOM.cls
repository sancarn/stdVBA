VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdCOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
'Dim x as stdCOM: set x = stdCOM.create(obj)
'set int = x.QueryInterface("")
'retVar = int.CallVt(EMyInterface.MyFunction, vbReturnVar, param1, param2, param3, param4, ...)




#If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As LongPtr, ByVal offsetinVft As LongPtr, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As VbVarType, ByRef paValues As LongPtr, ByRef retVAR As Variant) As Long
    Private Declare PtrSafe Sub SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long)
    Private Declare PtrSafe Function IIDFromString Lib "ole32.dll" (ByVal pstCLS As LongPtr, ByRef iid As GUID) As Long
#Else
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Private Declare Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As VbVarType, ByRef paValues As Long, ByRef retVAR As Variant) As Long
    Private Declare Sub SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long)
    Private Declare Function IIDFromString Lib "ole32.dll" (ByVal pstCLS As Long, ByRef iid As guid) As Long
#End If
'
'GUID struct for QueryInterface
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type TTYPEDESC
    #If Win64 Then
        pTypeDesc As LongLong
    #Else
        pTypeDesc As Long
    #End If
    vt            As Integer
End Type

Private Type TPARAMDESC
    #If Win64 Then
        pPARAMDESCEX  As LongLong
    #Else
        pPARAMDESCEX  As Long
    #End If
    wParamFlags       As Integer
End Type


Private Type TELEMDESC
    tdesc  As TTYPEDESC
    pdesc  As TPARAMDESC
End Type

Private Type TYPEATTR
        aGUID As GUID
        LCID As Long
        dwReserved As Long
        memidConstructor As Long
        memidDestructor As Long
        #If Win64 Then
            lpstrSchema As LongLong
        #Else
            lpstrSchema As Long
        #End If
        cbSizeInstance As Integer
        typekind As Long
        cFuncs As Integer
        cVars As Integer
        cImplTypes As Integer
        cbSizeVft As Integer
        cbAlignment As Integer
        wTypeFlags As Integer
        wMajorVerNum As Integer
        wMinorVerNum As Integer
        tdescAlias As Long
        idldescType As Long
End Type


Private Type FUNCDESC
    memid As Long                  'The function member ID (DispId).
    #If Win64 Then
        lprgscode As LongLong         'Pointer to status code
        lprgelemdescParam As LongLong 'Pointer to description of the element.
    #Else
        lprgscode As Long             'Pointer to status code
        lprgelemdescParam As Long     'Pointer to description of the element.
    #End If
    funckind As Long                 'virtual, static, or dispatch-only
    INVOKEKIND As Long               'VbMethod / VbGet / VbSet / VbLet
    CallConv As Long                 'typically will be stdecl
    cParams As Integer               'number of parameters
    cParamsOpt As Integer            'number of optional parameters
    oVft As Integer                  'For FUNC_VIRTUAL, specifies the offset in the VTBL.
    cScodes As Integer               'The number of possible return values.
    elemdescFunc As TELEMDESC        'The function return type
    wFuncFlags As Integer            'The function flags. See FUNCFLAGS.
End Type


'Many interfaces are stored [here](http://cbotton.com/pictures/net/ole2.reg)
'not sure if they are useful.
'Not all of these IIDs will be useful. But hopefully some will be.
Public Enum EKnownInterfaces
  IUnknown                '00000000-0000-0000-C000-000000000046
  IDispatch               '00020400-0000-0000-C000-000000000046
  ITypeInfo               '00020401-0000-0000-C000-000000000046
  ITypeLib                '00020402-0000-0000-C000-000000000046
  ITypeComp               '00020403-0000-0000-C000-000000000046
  IEnumVARIANT            '00020404-0000-0000-C000-000000000046
  ICreateTypeInfo         '00020405-0000-0000-C000-000000000046
  ICreateTypeLib          '00020406-0000-0000-C000-000000000046
  IClassFactory           '00000001-0000-0000-C000-000000000046
  IMalloc                 '00000002-0000-0000-C000-000000000046
  IMarshal                '00000003-0000-0000-C000-000000000046
  IRpcChannel             '00000004-0000-0000-C000-000000000046
  IRpcStub                '00000005-0000-0000-C000-000000000046
  IRpcProxy               '00000007-0000-0000-C000-000000000046
  IPSFactory              '00000009-0000-0000-C000-000000000046
  ILockBytes              '0000000A-0000-0000-C000-000000000046
  IStorage                '0000000B-0000-0000-C000-000000000046
  IStream                 '0000000C-0000-0000-C000-000000000046
  IEnumSTATSTG            '0000000D-0000-0000-C000-000000000046
  IBindCtx                '0000000E-0000-0000-C000-000000000046
  IMoniker                '0000000F-0000-0000-C000-000000000046
  IRunningObjectTable     '00000010-0000-0000-C000-000000000046
  IRootStorage            '00000012-0000-0000-C000-000000000046
  IMessageFilter          '00000016-0000-0000-C000-000000000046
  IStdMarshalInfo         '00000018-0000-0000-C000-000000000046
  IExternalConnection     '00000019-0000-0000-C000-000000000046
  IEnumUnknown            '00000100-0000-0000-C000-000000000046
  IEnumString             '00000101-0000-0000-C000-000000000046
  IEnumMoniker            '00000102-0000-0000-C000-000000000046
  IEnumFORMATETC          '00000103-0000-0000-C000-000000000046
  IEnumOLEVERB            '00000104-0000-0000-C000-000000000046
  IEnumSTATDATA           '00000105-0000-0000-C000-000000000046
  IPersistStream          '00000109-0000-0000-C000-000000000046
  IPersistStorage         '0000010A-0000-0000-C000-000000000046
  IPersistFile            '0000010B-0000-0000-C000-000000000046
  IPersist                '0000010C-0000-0000-C000-000000000046
  IViewObject             '0000010D-0000-0000-C000-000000000046
  IDataObject             '0000010E-0000-0000-C000-000000000046
  IAdviseSink             '0000010F-0000-0000-C000-000000000046
  IDataAdviseHolder       '00000110-0000-0000-C000-000000000046
  IOleAdviseHolder        '00000111-0000-0000-C000-000000000046
  IOleObject              '00000112-0000-0000-C000-000000000046
  IOleInPlaceObject       '00000113-0000-0000-C000-000000000046
  IOleWindow              '00000114-0000-0000-C000-000000000046
  IOleInPlaceUIWindow     '00000115-0000-0000-C000-000000000046
  IOleInPlaceFrame        '00000116-0000-0000-C000-000000000046
  IOleInPlaceActiveObject '00000117-0000-0000-C000-000000000046
  IOleClientSite          '00000118-0000-0000-C000-000000000046
  IOleInPlaceSite         '00000119-0000-0000-C000-000000000046
  IParseDisplayName       '0000011A-0000-0000-C000-000000000046
  IOleContainer           '0000011B-0000-0000-C000-000000000046
  IOleItemContainer       '0000011C-0000-0000-C000-000000000046
  IOleLink                '0000011D-0000-0000-C000-000000000046
  IOleCache               '0000011E-0000-0000-C000-000000000046
  IDropSource             '00000121-0000-0000-C000-000000000046
  IDropTarget             '00000122-0000-0000-C000-000000000046
  IDebugStream            '00000124-0000-0000-C000-000000000046
  IAdviseSink2            '00000125-0000-0000-C000-000000000046
  IRunnableObject         '00000126-0000-0000-C000-000000000046
  IViewObject2            '00000127-0000-0000-C000-000000000046
  IOleCache2              '00000128-0000-0000-C000-000000000046
  IOleCacheControl        '00000129-0000-0000-C000-000000000046
  IInspectable            'af86e2e0-b12d-4c6a-9c5a-d7aa65101e90
End Enum

'Enumerator for InterfaceInfo parts, this can be useful while scraping / reflection
Public Enum EInterfaceInfoIndex
  sName
  iInvokeKind
  iOffset
  iDispID
End Enum

'[IUnknown](https://en.wikipedia.org/wiki/IUnknown)
'0      HRESULT  QueryInterface ([in] REFIID riid, [out] void **ppvObject)
'1      ULONG    AddRef ()
'2      ULONG    Release ()
Public Enum EIUnknown
  QueryInterface
  AddRef
  Release
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

'[ITypeInfo](https://github.com/tpn/winsdk-10/blob/master/Include/10.0.16299.0/um/OAIdl.h#L2683) extends IUnknown
'0      HRESULT  QueryInterface ([in] REFIID riid, [out] void **ppvObject)
'1      ULONG    AddRef ()
'2      ULONG    Release ()
'3      HRESULT  GetTypeAttr([out] TYPEATTR **ppTypeAttr )
'4      HRESULT  GetTypeComp([out] ITypeComp **ppTComp )
'5      HRESULT  GetFuncDesc([in] UINT index, [out] FUNCDESC **ppFuncDesc)
'6      HRESULT  GetVarDesc([in] UINT index, [out] VARDESC **ppVarDesc)
'7      HRESULT  GetNames([in] MEMBERID memid, [out] BSTR *rgBstrNames, [in] UINT cMaxNames, [out] UINT *pcNames)
'8      HRESULT  GetRefTypeOfImplType( [in] UINT index, [out] HREFTYPE *pRefType)
'9      HRESULT  GetImplTypeFlags( [in] UINT index, [out] INT *pImplTypeFlags)
'10     HRESULT  GetIDsOfNames( [in] LPOLESTR *rgszNames, [in] UINT cNames, [out] MEMBERID *pMemId)
'11     HRESULT  Invoke( [in] PVOID pvInstance, [in] MEMBERID memid, [in] WORD wFlags, [out][in] DISPPARAMS *pDispParams, [out] VARIANT *pVarResult, [out] EXCEPINFO *pExcepInfo, [out] UINT *puArgErr)
'12     HRESULT  GetDocumentation( [in] MEMBERID memid, [out] BSTR *pBstrName, [out] BSTR *pBstrDocString, [out] DWORD *pdwHelpContext, [out] BSTR *pBstrHelpFile)
'13     HRESULT  GetDllEntry( [in] MEMBERID memid, [in] INVOKEKIND invKind, [out] BSTR *pBstrDllName, [out] BSTR *pBstrName, [out] WORD *pwOrdinal)
'14     HRESULT  GetRefTypeInfo( [in] HREFTYPE hRefType, [out] ITypeInfo **ppTInfo)
'15     HRESULT  AddressOfMember( [in] MEMBERID memid, [in] INVOKEKIND invKind, [out] PVOID *ppv)
'16     HRESULT  CreateInstance( [in] IUnknown *pUnkOuter, [in] REFIID riid, [out] PVOID *ppvObj)
'17     HRESULT  GetMops( [in] MEMBERID memid, [out] BSTR *pBstrMops)
'18     HRESULT  GetContainingTypeLib( [out] ITypeLib **ppTLib, [out] UINT *pIndex)
'19     void     ReleaseTypeAttr( [in] TYPEATTR *pTypeAttr)
'20     void     ReleaseFuncDesc( [in] FUNCDESC *pFuncDesc)
'21     void     ReleaseVarDesc( [in] VARDESC *pVarDesc)
Public Enum EITypeInfo
  QueryInterface
  AddRef
  Release
  GetTypeAttr
  GetTypeComp
  GetFuncDesc
  GetVarDesc
  GetNames
  GetRefTypeOfImplType
  GetImplTypeFlags
  GetIDsOfNames
  Invoke
  GetDocumentation
  GetDllEntry
  GetRefTypeInfo
  AddressOfMember
  CreateInstance
  GetMops
  GetContainingTypeLib
  ReleaseTypeAttr
  ReleaseFuncDesc
  ReleaseVarDesc
End Enum


Private Const S_OK = 0
Private Const E_NOINTERFACE As Long = &H80004002
Private Const E_POINTER     As Long = &H80004003

#If Win64 Then
  Private Const PTRSIZE = 8
  Private Const vbLongPtr = VbVarType.vbLongLong
#Else
  Private Const PTRSIZE = 4
  Private Const vbLongPtr = VbVarType.vbLong
#End If
Private Const CC_STDCALL As Long = 4

Private pObjPtr As LongPtr
Private oMeta As IUnknown

Public Function Create(ByRef obj As IUnknown) As stdCOM
  Set Create = New stdCOM
  Call Create.protInitFromObject(obj)
End Function

Public Function CreateFromPtr(ByVal ptr As LongPtr)
  Set CreateFromPtr = New stdCOM
  Call CreateFromPtr.protInitFromPtr(ptr)
End Function


Public Function unkQueryInterface(ByVal sIID As String) As stdCOM
  Dim IInterface As IUnknown
  Dim tGUID As GUID
  If IIDFromString(StrPtr(sIID), tGUID) = S_OK Then
    Dim hResult As Long: hResult = CallVT(EIUnknown.QueryInterface, vbLong, VarPtr(tGUID.Data1), VarPtr(IInterface))
    If hResult = S_OK Then
      Set unkQueryInterface = stdCOM.Create(IInterface)
    ElseIf hResult = E_POINTER Then
      Set unkQueryInterface = Nothing
    ElseIf hResult = E_NOINTERFACE Then
      Err.Raise E_NOINTERFACE, "stdCOM", "IUnknown::QueryInterface - Interface not implemented."
    Else
      Err.Raise hResult, "stdCOM", "IUnknown::QueryInterface - Unknown error occurred."
    End If
  Else
    Err.Raise Err.LastDllError, "stdCOM", "IUnknown::QueryInterface - Cannot retrieve IID from String."
  End If
End Function

Public Sub unkAddRef()
  CallVT EIUnknown.AddRef, vbLong
End Sub

Public Sub unkRelease()
  CallVT EIUnknown.Release, vbLong
End Sub

Public Function QueryKnownInterface(ByVal interface As EKnownInterfaces) As stdCOM
  Dim sIID As String: sIID = KnownIID(interface)
  If sIID <> "" Then
    Set QueryKnownInterface = unkQueryInterface(sIID)
  Else
    Err.Raise 1, "stdCOM::QueryKnownInterface", "Interface IID not known."
  End If
End Function

Public Function CallVT(ByVal iOffset As Long, ByVal ReturnType As VbVarType, ParamArray A() As Variant) As Variant
    'Create array from params
    Dim vParam() As Variant: vParam = A
    
    'Create array of pointers to params, and array of types for DispCallFunc
    Dim lIdx As Long, vType(0 To 63) As VbVarType
    #If VBA7 Then
      Dim vPtr(0 To 63) As LongPtr
    #Else
      Dim vPtr(0 To 63) As Long
    #End If
    
    For lIdx = 0 To UBound(vParam)
        vType(lIdx) = VarType(vParam(lIdx))
        vPtr(lIdx) = VarPtr(vParam(lIdx))
    Next

    'Call function
    Dim hResult As Long
    hResult = DispCallFunc(pObjPtr, iOffset * PTRSIZE, CC_STDCALL, ReturnType, lIdx, vType(0), vPtr(0), CallVT)
    If hResult < 0 Then
        Err.Raise hResult, "DispCallFunc"
    End If
End Function


Public Property Get TypeInfo() As stdCOM
  On Error GoTo CannotQuery
    Set TypeInfo = Me.QueryKnownInterface(ITypeInfo)
    Exit Property
CannotQuery:
  On Error GoTo UnknownError
  'Attempt to query from IDispatch
  With Me.QueryKnownInterface(IDispatch)
    Dim pTypeInfo As IUnknown
    Dim hResult As Long: hResult = .CallVT(EIDispatch.GetTypeInfo, vbLongPtr, NULL_PTR, NULL_PTR, VarPtr(pTypeInfo))
    If hResult <> S_OK Then Err.Raise hResult, "InterfaceInfo", "stdCOM::InterfaceInfo"
    Set TypeInfo = stdCOM.Create(pTypeInfo)
  End With
  Exit Property
UnknownError:
  Set TypeInfo = Nothing
End Property



Public Property Get InterfaceInfo() As Collection
  Static pInterfaceInfo As Collection
  If pInterfaceInfo Is Nothing Then
    'Wrap pTypeInfo in stdCOM
    Dim oTypeInfo As stdCOM: Set oTypeInfo = Me.TypeInfo
    
    #If Win64 Then
      Dim lPtr As LongLong
      Const CountPos As Long = 12
    #Else
      Dim lPtr As Long
      Const CountPos As Long = 11
    #End If
    'Get TypeAttribute
    Call oTypeInfo.CallVT(EITypeInfo.GetTypeAttr, vbLong, VarPtr(lPtr))
    
    'Fix bug in VBA/VB6 engine - for some reason the next CopyMemory statement will set the following
    'value to nothing
    Dim vTmp: Set vTmp = oTypeInfo
    
    'Only fill the type size which is required, otherwise we'd use LENB(tTypeAttr)
    Dim tTypeAttr As TYPEATTR
    CopyMemory tTypeAttr, ByVal lPtr, CountPos * PTRSIZE
    
    'Release TypeAttribute
    Call oTypeInfo.CallVT(EITypeInfo.ReleaseTypeAttr, vbLong, lPtr)
    
    Dim oRet As Collection
    Set oRet = New Collection
    
    Dim tFuncDesc As FUNCDESC
    Dim iFuncMemSize As Long: iFuncMemSize = 2 * PTRSIZE + 12: iFuncMemSize = LenB(tFuncDesc)
    
    'Enumerate over all members of typeinfo
    Dim i As Long
    For i = 0 To tTypeAttr.cFuncs - 1
      lPtr = 0
      Call oTypeInfo.CallVT(EITypeInfo.GetFuncDesc, vbLong, i, VarPtr(lPtr))
      If lPtr <> 0 Then
        CopyMemory tFuncDesc, ByVal lPtr, iFuncMemSize
        Call oTypeInfo.CallVT(EITypeInfo.ReleaseFuncDesc, vbLong, lPtr)
        
        Dim sName As String: sName = vbNullString
        Call oTypeInfo.CallVT(EITypeInfo.GetDocumentation, vbLong, tFuncDesc.memid, VarPtr(sName), NULL_PTR, NULL_PTR, NULL_PTR)
        oRet.Add Array(sName, tFuncDesc.INVOKEKIND, i, tFuncDesc.memid)
      End If
      If (i Mod 50) = 0 Then DoEvents
    Next
    Set pInterfaceInfo = oRet
  End If
  
  Set InterfaceInfo = pInterfaceInfo
  Exit Property
UnknownError:
  Set InterfaceInfo = Nothing
End Property

Public Property Get Methods() As Collection
  Static oMethods As Collection
  If oMethods Is Nothing Then
    Set oMethods = New Collection
    Dim info As Variant
    For Each info In InterfaceInfo
      If info(1) = vbMethod Then
        oMethods.Add info(0)
      End If
    Next
  End If
  Set Methods = oMethods
End Property

Public Property Get Properties() As Collection
  Static oProperties As Collection
  If oProperties Is Nothing Then
    Set oProperties = New Collection
    Dim info As Variant
    For Each info In InterfaceInfo
      If info(1) <> vbMethod Then
        On Error Resume Next
          oProperties.Add info(0), info(0)
        On Error GoTo 0
      End If
    Next
  End If

  Set Properties = oProperties
End Property







'*************
'* PROTECTED *
'*************
Friend Sub protInitFromObject(ByVal obj As IUnknown)
  pObjPtr = ObjPtr(obj)
  Set oMeta = obj
End Sub
Friend Sub protInitFromPtr(ByVal ptr As LongPtr)
  pObjPtr = ptr
  Call unkAddRef
End Sub
Private Sub Class_Terminate()
  If oMeta Is Nothing Then Call unkRelease
End Sub

Private Function KnownIID(ByVal iIID As EKnownInterfaces) As String
  Select Case iIID
    Case IUnknown:                KnownIID = "00000000-0000-0000-C000-000000000046"
    Case IDispatch:               KnownIID = "00020400-0000-0000-C000-000000000046"
    Case ITypeInfo:               KnownIID = "00020401-0000-0000-C000-000000000046"
    Case ITypeLib:                KnownIID = "00020402-0000-0000-C000-000000000046"
    Case ITypeComp:               KnownIID = "00020403-0000-0000-C000-000000000046"
    Case IEnumVARIANT:            KnownIID = "00020404-0000-0000-C000-000000000046"
    'Note sure if these are going to be useful,
    'but I found the IIDs and IIDs are generally hard to find
    'therefore, I'm providing them here
    Case ICreateTypeInfo:         KnownIID = "00020405-0000-0000-C000-000000000046"
    Case ICreateTypeLib:          KnownIID = "00020406-0000-0000-C000-000000000046"
    Case IClassFactory:           KnownIID = "00000001-0000-0000-C000-000000000046"
    Case IMalloc:                 KnownIID = "00000002-0000-0000-C000-000000000046"
    Case IMarshal:                KnownIID = "00000003-0000-0000-C000-000000000046"
    Case IRpcChannel:             KnownIID = "00000004-0000-0000-C000-000000000046"
    Case IRpcStub:                KnownIID = "00000005-0000-0000-C000-000000000046"
    Case IRpcProxy:               KnownIID = "00000007-0000-0000-C000-000000000046"
    Case IPSFactory:              KnownIID = "00000009-0000-0000-C000-000000000046"
    Case ILockBytes:              KnownIID = "0000000A-0000-0000-C000-000000000046"
    Case IStorage:                KnownIID = "0000000B-0000-0000-C000-000000000046"
    Case IStream:                 KnownIID = "0000000C-0000-0000-C000-000000000046"
    Case IEnumSTATSTG:            KnownIID = "0000000D-0000-0000-C000-000000000046"
    Case IBindCtx:                KnownIID = "0000000E-0000-0000-C000-000000000046"
    Case IMoniker:                KnownIID = "0000000F-0000-0000-C000-000000000046"
    Case IRunningObjectTable:     KnownIID = "00000010-0000-0000-C000-000000000046"
    Case IRootStorage:            KnownIID = "00000012-0000-0000-C000-000000000046"
    Case IMessageFilter:          KnownIID = "00000016-0000-0000-C000-000000000046"
    Case IStdMarshalInfo:         KnownIID = "00000018-0000-0000-C000-000000000046"
    Case IExternalConnection:     KnownIID = "00000019-0000-0000-C000-000000000046"
    Case IEnumUnknown:            KnownIID = "00000100-0000-0000-C000-000000000046"
    Case IEnumString:             KnownIID = "00000101-0000-0000-C000-000000000046"
    Case IEnumMoniker:            KnownIID = "00000102-0000-0000-C000-000000000046"
    Case IEnumFORMATETC:          KnownIID = "00000103-0000-0000-C000-000000000046"
    Case IEnumOLEVERB:            KnownIID = "00000104-0000-0000-C000-000000000046"
    Case IEnumSTATDATA:           KnownIID = "00000105-0000-0000-C000-000000000046"
    Case IPersistStream:          KnownIID = "00000109-0000-0000-C000-000000000046"
    Case IPersistStorage:         KnownIID = "0000010A-0000-0000-C000-000000000046"
    Case IPersistFile:            KnownIID = "0000010B-0000-0000-C000-000000000046"
    Case IPersist:                KnownIID = "0000010C-0000-0000-C000-000000000046"
    Case IViewObject:             KnownIID = "0000010D-0000-0000-C000-000000000046"
    Case IDataObject:             KnownIID = "0000010E-0000-0000-C000-000000000046"
    Case IAdviseSink:             KnownIID = "0000010F-0000-0000-C000-000000000046"
    Case IDataAdviseHolder:       KnownIID = "00000110-0000-0000-C000-000000000046"
    Case IOleAdviseHolder:        KnownIID = "00000111-0000-0000-C000-000000000046"
    Case IOleObject:              KnownIID = "00000112-0000-0000-C000-000000000046"
    Case IOleInPlaceObject:       KnownIID = "00000113-0000-0000-C000-000000000046"
    Case IOleWindow:              KnownIID = "00000114-0000-0000-C000-000000000046"
    Case IOleInPlaceUIWindow:     KnownIID = "00000115-0000-0000-C000-000000000046"
    Case IOleInPlaceFrame:        KnownIID = "00000116-0000-0000-C000-000000000046"
    Case IOleInPlaceActiveObject: KnownIID = "00000117-0000-0000-C000-000000000046"
    Case IOleClientSite:          KnownIID = "00000118-0000-0000-C000-000000000046"
    Case IOleInPlaceSite:         KnownIID = "00000119-0000-0000-C000-000000000046"
    Case IParseDisplayName:       KnownIID = "0000011A-0000-0000-C000-000000000046"
    Case IOleContainer:           KnownIID = "0000011B-0000-0000-C000-000000000046"
    Case IOleItemContainer:       KnownIID = "0000011C-0000-0000-C000-000000000046"
    Case IOleLink:                KnownIID = "0000011D-0000-0000-C000-000000000046"
    Case IOleCache:               KnownIID = "0000011E-0000-0000-C000-000000000046"
    Case IDropSource:             KnownIID = "00000121-0000-0000-C000-000000000046"
    Case IDropTarget:             KnownIID = "00000122-0000-0000-C000-000000000046"
    Case IDebugStream:            KnownIID = "00000124-0000-0000-C000-000000000046"
    Case IAdviseSink2:            KnownIID = "00000125-0000-0000-C000-000000000046"
    Case IRunnableObject:         KnownIID = "00000126-0000-0000-C000-000000000046"
    Case IViewObject2:            KnownIID = "00000127-0000-0000-C000-000000000046"
    Case IOleCache2:              KnownIID = "00000128-0000-0000-C000-000000000046"
    Case IOleCacheControl:        KnownIID = "00000129-0000-0000-C000-000000000046"
    Case IInspectable:            KnownIID = "af86e2e0-b12d-4c6a-9c5a-d7aa65101e90"
  End Select
  KnownIID = "{" & KnownIID & "}"
End Function







