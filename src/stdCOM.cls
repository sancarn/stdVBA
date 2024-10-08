VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdCOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'Dim x as stdCOM: set x = stdCOM.create(obj)
'set int = x.QueryInterface("")
'retVar = int.CallVt(EMyInterface.MyFunction, vbReturnVar, param1, param2, param3, param4, ...)

'TODO: Analyse and add anything missing from:
' * https://referencesource.microsoft.com/#system.data/fx/src/data/System/Data/OleDb/OLEDB_Util.cs,4254532afac0bd58,references
#If Win64 Then
  Private Const NULL_PTR as LongLong = 0^
#Else
  Private Const NULL_PTR As Long = 0&
#End If

#If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
    Private Declare PtrSafe Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As LongPtr, ByVal offsetinVft As LongPtr, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByVal paTypes As LongPtr, ByVal paValues As LongPtr, ByRef retVAR As Variant) As Long
    Private Declare PtrSafe Sub SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long)
    Private Declare PtrSafe Function IIDFromString Lib "ole32.dll" (ByVal pstCLS As LongPtr, ByRef iid As GUID) As Long
    Private Declare PtrSafe Function CLSIDFromString Lib "ole32.dll" (ByVal pstCLS As LongPtr, ByRef iid As GUID) As Long
    
    'Iterate the ROT
    Private Declare PtrSafe Function GetRunningObjectTable Lib "ole32.dll" (ByVal dwReserved As Long, pROT As LongPtr) As Long
    Private Declare PtrSafe Function CreateBindCtx Lib "ole32.dll" (ByVal dwReserved As Long, pBindCtx As LongPtr) As Long
    Private Declare PtrSafe Function SysReAllocString Lib "oleAut32.dll" (ByVal pBSTR As LongPtr, Optional ByVal pszStrPtr As LongPtr) As Long
    Private Declare PtrSafe Function ProgIDFromCLSID Lib "ole32.dll" (ByRef rclsid As GUID, ByVal ProgID as LongPtr) as Long

    'Register to ROT
    Private Declare PtrSafe Function RegisterActiveObject32 Lib "oleAut32.dll" Alias "RegisterActiveObject" (ByVal pUnk As IUnknown, ByRef rclsid As GUID, ByVal dwFlags As Long, ByRef pdwRegister As Long) As Long
    Private Declare PtrSafe Function RevokeActiveObject32 Lib "oleAut32.dll" Alias "RevokeActiveObject" (ByVal dwRegister As Long, ByVal pvReserved As LongPtr) As Long
    Private Declare PtrSafe Function CoDisconnectObject Lib "ole32.dll" (ByVal pUnk As IUnknown, ByRef pvReserved As Long) As Long
    Private Declare PtrSafe Function CLSIDFromProgID Lib "ole32.dll" (ByVal ProgID As LongPtr, ByRef rclsid As GUID) As Long
    Private Declare PtrSafe Function CoCreateInstance Lib "ole32.dll" (ByRef rclsid As GUID, ByVal pUnkOuter As LongPtr, ByVal dwClsContext As Long, ByRef riid As GUID, ByRef ppv As LongPtr) As Long

    'WinRT
    Private Declare PtrSafe Function WindowsCreateString Lib "Combase.dll" (ByVal sourceString As LongPtr, ByVal Length As Long, ByRef hString As LongPtr) As Long
    Private Declare PtrSafe Function WindowsDeleteString Lib "Combase.dll" (ByVal sourceString As LongPtr) As Long
    Private Declare PtrSafe Function RoGetActivationFactory Lib "Combase.dll" (ByVal activatableClassId As LongPtr, ByRef riid As GUID, ByRef factory As Long) As Long
    Private Declare PtrSafe Function RoActivateInstance Lib "Combase.dll" (ByVal activatableClassId As LongPtr, ByRef instance As LongPtr) As Long
#Else
    Private Enum LongPtr
      [_]
    End Enum
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
    'FIX: The use of `VbVarType` for the type of `paTypes` on 32 bit causes Invalid Callee error. Workaround is to use `Integer` instead as below.
    Private Declare Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As LongPtr, ByVal offsetinVft As LongPtr, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByVal paTypes As LongPtr, ByVal paValues As LongPtr, ByRef retVAR As Variant) As Long
    Private Declare Sub SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long)
    Private Declare Function IIDFromString Lib "ole32.dll" (ByVal pstCLS As LongPtr, ByRef iid As GUID) As Long
    Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal pstCLS As LongPtr, ByRef iid As GUID) As Long
    
    'Iterate the ROT
    Private Declare Function GetRunningObjectTable Lib "ole32.dll" (ByVal dwReserved As Long, pROT As LongPtr) As Long
    Private Declare Function CreateBindCtx Lib "ole32.dll" (ByVal dwReserved As Long, pBindCtx As LongPtr) As Long
    Private Declare Function SysReAllocString Lib "oleAut32.dll" (ByVal pBSTR As LongPtr, Optional ByVal pszStrPtr As LongPtr) As Long
    Private Declare Function ProgIDFromCLSID Lib "ole32.dll" (ByRef rclsid As GUID, ByVal ProgID as LongPtr) as Long

    'Register to ROT
    Private Declare Function RegisterActiveObject32 Lib "oleAut32.dll" Alias "RegisterActiveObject" (ByVal pUnk As IUnknown, ByRef rclsid As GUID, ByVal dwFlags As Long, ByRef pdwRegister As Long) As Long
    Private Declare Function RevokeActiveObject32 Lib "oleAut32.dll" Alias "RevokeActiveObject" (ByVal dwRegister As Long, ByVal pvReserved As Long) As Long
    Private Declare Function CoDisconnectObject Lib "ole32.dll" (ByVal pUnk As IUnknown, ByRef pvReserved As Long) As Long
    Private Declare Function CLSIDFromProgID Lib "ole32.dll" (ByVal ProgID As Long, ByRef rclsid As GUID) As Long
    Private Declare Function CoCreateInstance Lib "ole32.dll" (ByRef rclsid As GUID, ByVal pUnkOuter As LongPtr, ByVal dwClsContext As Long, ByRef riid As GUID, ByRef ppv As LongPtr) As Long

    'WinRT
    Private Declare Function WindowsCreateString Lib "Combase.dll" (ByVal sourceString As LongPtr, ByVal Length As Long, ByRef hString As LongPtr) As Long
    Private Declare Function WindowsDeleteString Lib "Combase.dll" (ByVal sourceString As LongPtr) As Long
    Private Declare Function RoGetActivationFactory Lib "Combase.dll" (ByVal activatableClassId As LongPtr, ByRef riid As GUID, ByRef factory As Long) As Long
    Private Declare Function RoActivateInstance Lib "Combase.dll" (ByVal activatableClassId As LongPtr, ByRef instance As LongPtr) As Long
#End If

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


'[IRunningObjectTable](https://github.com/tpn/winsdk-10/blob/9b69fd26ac0c7d0b83d378dba01080e93349c2ed/Include/10.0.14393.0/um/ObjIdl.Idl#L394-L440)
'@remark also see [IRunningObjectTable](https://github.com/khchen/winim/blob/master/winim/inc/objbase.nim#L598-L605) extends IUnknown
'0      HRESULT  QueryInterface      ([in]  REFIID riid, [out] void **ppvObject)
'1      HRESULT  AddRef              ()
'2      HRESULT  Release             ()
'3      HRESULT  Register            ([in]  DWORD grfFlags, [in] IUnknown *punkObject, [in] IMoniker *pmkObjectName, [out] DWORD *pdwRegister)
'4      HRESULT  Revoke              ([in]  DWORD dwRegister)
'5      HRESULT  IsRunning           ([in]  IMoniker *pmkObjectName)
'6      HRESULT  GetObject_           ([in]  IMoniker *pmkObjectName, [out] IUnknown **ppunkObject)
'7      HRESULT  NoteChangeTime      ([in]  DWORD dwRegister,[in] FILETIME *pfiletime)
'8      HRESULT  GetTimeOfLastChange ([in]  IMoniker *pmkObjectName, [out] FILETIME *pfiletime)
'9      HRESULT  EnumRunning         ([out] IEnumMoniker **ppenumMoniker)
Public Enum EIRunningObjectTable
  QueryInterface
  AddRef
  Release
  Register
  Revoke
  IsRunning
  GetObject_
  NoteChangeTime
  GetTimeOfLastChange
  EnumRunning
End Enum

'[IEnumVARIANT](https://github.com/tpn/winsdk-10/blob/master/Include/10.0.16299.0/um/OAIdl.h#L2398-L2432)
'0      HRESULT  QueryInterface      ([in]  REFIID riid, [out] void **ppvObject)
'1      HRESULT  AddRef              ()
'2      HRESULT  Release             ()
'3      HRESULT  Next                ([in]  ULONG   countToFetch, [out] VARIANT *varArray ,[out] ULONG *pCountFetched)
'4      HRESULT  Skip                ([in]  ULONG   countToSkip)
'5      HRESULT  Reset               ()
'6      HRESULT  Clone               ([out] IEnumVARIANT **ppEnum)
Public Enum EIEnumVARIANT
  QueryInterface
  AddRef
  Release
  Nextt
  Skip
  Reset
  Clone
End Enum

'[IEnumMoniker](https://github.com/tpn/winsdk-10/blob/9b69fd26ac0c7d0b83d378dba01080e93349c2ed/Include/10.0.14393.0/um/ObjIdl.Idl#L325-L352)
'@remark Also see https://fossies.org/dox/wine-6.0.2/interfaceIEnumMoniker.html; RemoteNext is not documented on docs.microsoft.com
'0      HRESULT  QueryInterface      ([in]  REFIID riid, [out] void **ppvObject)
'1      HRESULT  AddRef              ()
'2      HRESULT  Release             ()
'3      HRESULT  Next                ([in]  ULONG   countToFetch, [out] IMoniker *varArray ,[out] ULONG *pCountFetched)
'4      HRESULT  RemoteNext          ([in]  ULONG   countToFetch, [out] IMoniker *varArray ,[out] ULONG *pCountFetched)
'5      HRESULT  Skip                ([in]  ULONG   countToSkip)
'6      HRESULT  Reset               ()
'7      HRESULT  Clone               ([out] IEnumMoniker **ppenum)
Public Enum EIEnumMoniker
  QueryInterface
  AddRef
  Release
  Nextt    'or RemoteNext
  Skip
  Reset
  Clone
End Enum

'[IMoniker](https://github.com/tpn/winsdk-10/blob/9b69fd26ac0c7d0b83d378dba01080e93349c2ed/Include/10.0.14393.0/um/ObjIdl.Idl#L504-L621)
'@devRemark includes [IUnknown](https://en.wikipedia.org/wiki/IUnknown)
'@devRemark includes [IPersist](https://github.com/tpn/winsdk-10/blob/9b69fd26ac0c7d0b83d378dba01080e93349c2ed/Include/10.0.14393.0/um/ObjIdl.Idl#L453-L462)
'@devRemark includes [IPersistStream](https://github.com/tpn/winsdk-10/blob/9b69fd26ac0c7d0b83d378dba01080e93349c2ed/Include/10.0.14393.0/um/ObjIdl.Idl#L470-L496)
'0      HRESULT QueryInterface      ([in]  REFIID riid, [out] void **ppvObject)
'1      HRESULT AddRef              ()
'2      HRESULT Release             ()
'3      HRESULT GetClassID          ([out] CLSID *pClassID)
'4      HRESULT IsDirty             ()
'5      HRESULT Load                ([in] IStream *pStm)
'6      HRESULT Save                ([in] IStream *pStm, [in] BOOL fClearDirty)
'7      HRESULT GetSizeMax          ([out] ULARGE_INTEGER *pcbSize)
'8      HRESULT BindToObject        ([in] IBindCtx *pbc, [in] IMoniker *pmkToLeft, [in] REFIID riidResult, [out] void **ppvResult)       'or RemoteBindToObject
'9      HRESULT BindToStorage       ([in] IBindCtx *pbc, [in] IMoniker *pmkToLeft, [in] REFIID riidResult, [out] void **ppvResult)       'or RemoteBindToStorage
'10     HRESULT Reduce              ([in] IBindCtx *pbc, [in] DWORD dwReduceHowFar,[in, out] IMoniker **ppmkToLeft, [out] IMoniker **ppmkReduced)
'11     HRESULT ComposeWith         ([in] IMoniker *pmkRight, [in] BOOL fOnlyIfNotGeneric, [out] IMoniker **ppmkComposite)
'12     HRESULT Enum                ([in] BOOL fForward, [out] IEnumMoniker **ppenumMoniker)
'13     HRESULT IsEqual             ([in] IMoniker *pmkOtherMoniker)
'14     HRESULT Hash                ([out] DWORD *pdwHash)
'15     HRESULT IsRunning           ([in] IBindCtx *pbc, [in] IMoniker *pmkToLeft, [in] IMoniker *pmkNewlyRunning)
'16     HRESULT GetTimeOfLastChange ([in] IBindCtx *pbc, [in] IMoniker *pmkToLeft, [out] FILETIME *pFileTime)
'17     HRESULT Inverse             ([out] IMoniker **ppmk)
'18     HRESULT CommonPrefixWith    ([in] IMoniker *pmkOther, [out] IMoniker **ppmkPrefix)
'19     HRESULT RelativePathTo      ([in] IMoniker *pmkOther, [out] IMoniker **ppmkRelPath)
'20     HRESULT GetDisplayName      ([in] IBindCtx *pbc, [in] IMoniker *pmkToLeft, [out] LPOLESTR *ppszDisplayName)
'21     HRESULT ParseDisplayName    ([in] IBindCtx *pbc, [in] IMoniker *pmkToLeft, [in] LPOLESTR pszDisplayName, [out] ULONG *pchEaten, [out] IMoniker **ppmkOut)
'22     HRESULT IsSystemMoniker     ([out] DWORD *pdwMksys)
Public Enum EIMoniker
  QueryInterface
  AddRef
  Release
  GetClassID
  IsDirty
  Load
  Save
  GetSizeMax
  BindToObject
  BindToStorage
  reduce
  ComposeWith
  Enumm
  IsEqual
  Hash
  IsRunning
  GetTimeOfLastChange
  Inverse
  CommonPrefixWith
  RelativePathTo
  GetDisplayName
  ParseDisplayName
  IsSystemMoniker
End Enum


Private Const S_OK          As Long = 0
Private Const E_NOINTERFACE As Long = &H80004002
Private Const E_POINTER     As Long = &H80004003
Private Const CC_STDCALL    As Long = 4

#If Win64 Then
  Private Const PTRSIZE = 8
  Private Const vbLongPtr = VbVarType.vbLongLong
#Else
  Private Const PTRSIZE = 4
  Private Const vbLongPtr = VbVarType.vbLong
#End If


'*******************
'* STATE VARIABLES *
'*******************
Private Type TThis
  Meta as IUnknown
  RegisteredOLEInstance as Long
  RegisteredTimerOLEInstance as Long
  #if VBA7 then
    ObjPtr as LongPtr
  #Else
    ObjPtr as Long
  #End If
End Type
Private This as TThis

'**********
'* EVENTS *
'**********

'****************
'* CONSTRUCTORS *
'****************

'Create from an object
'@constructor
'@param obj - Object to create raw COM object from
'@returns - A wrapper around a raw COM object
Public Function Create(ByRef obj As IUnknown) As stdCOM
  Set Create = New stdCOM
  Call Create.protInitFromObject(obj)
End Function

'Create from a pointer
'@constructor
'@param ptr - Object pointer to wrap in a raw COM object
'@returns - A wrapper around a raw COM object
'@remark Specifically this function reuires a `*Interface` pointer. E.G. `*IUnknown`/`*IDispatch`/`*ITypeInfo`. If your DLL function provides a `[out] **Interface` object the
'typical method is to pass in a varptr to the dll function as follows `Call SomeDLLCall(VarPtr(myPointer)): ... later ...: stdCOM.CreateFromPtr(myPointer)`
#if VBA7 then
  Public Function CreateFromPtr(ByVal ptr As LongPtr) As stdCOM
#else
  Public Function CreateFromPtr(ByVal ptr As Long) As stdCOM
#end if 
  Set CreateFromPtr = New stdCOM
  Call CreateFromPtr.protInitFromPtr(ptr)
End Function

'Create `stdCOM` objects from existing objects in the RunningObjectTable (ROT)
'@constructor
'@returns - A collection of `stdCOM` objects
'@remark This function is heavily inspired off code provided by `Jaafar Tribak`
'@devRemark stdCOM objects will auto-release when not needed
Public Function CreateFromActiveObjects() As Collection
  #If VBA7 Then
    Dim pROT As LongPtr, pBindCtx As LongPtr, pEnumMoniker As LongPtr, pMoniker As LongPtr, pStr As LongPtr, pMonikerObject As LongPtr
  #Else
    Dim pROT As Long, pBindCtx As Long, pEnumMoniker As Long, pMoniker As Long, pStr As Long, pMonikerObject As Long
  #End If
  
  Dim ret As Collection: Set ret = New Collection 'Collection to return
  If CLng(GetRunningObjectTable(0, pROT)) <> S_OK Then Exit Function                               '*IRunningObjectTable
  If CLng(CreateBindCtx(0, pBindCtx)) <> S_OK Then Exit Function                                   '*IBindCtx
  Dim CTX As stdCOM: Set CTX = stdCOM.CreateFromPtr(pBindCtx)                                      'stdCOM<IBindCtx>
  Dim ROT As stdCOM: Set ROT = stdCOM.CreateFromPtr(pROT).QueryKnownInterface(IRunningObjectTable) 'stdCOM<IRunningObjectTable>
  Call ROT.CallVT(EIRunningObjectTable.EnumRunning, vbLong, VarPtr(pEnumMoniker))                  '*IEnumMoniker
  Dim e As stdCOM: Set e = stdCOM.CreateFromPtr(pEnumMoniker)                                      'stdCOM<IEnumMoniker>
  'Dim e As stdCOM: Set e = ROT.CallVTW(EIRunningObjectTable.EnumRunning, vbLong, Null)            'stdCOM<IEnumMoniker>
  
  'Enumerate ROT
  Dim nCount As Long: nCount = 1&
  Dim nCountFetched As Long: nCountFetched = 1& 'Note: This has to be set else a crash occurs
  Do While e.CallVT(EIEnumMoniker.Nextt, vbLong, nCount, VarPtr(pMoniker), VarPtr(nCountFetched)) = S_OK     '*IMoniker
    'Note: This check is required, in some instances (still not certain when specifically) S_OK is returned, despite `nCountFetched` holding `0` and `nCount` holding `1`
    if nCountFetched = 0 then Exit Do

    Dim moniker As stdCOM: Set moniker = stdCOM.CreateFromPtr(pMoniker)                                   'stdCOM<IMoniker>

    'Get DisplayName and Object
    Dim sName As String: If moniker.CallVT(EIMoniker.GetDisplayName, vbLong, VarPtr(pBindCtx), VarPtr(pMoniker), VarPtr(pStr)) = S_OK Then _
        sName = GetStrFromPtrW(pStr)
    Call ROT.CallVT(EIRunningObjectTable.GetObject_, vbLong, pMoniker, VarPtr(pMonikerObject))                          '*Interface
    Dim obj As stdCOM: Set obj = stdCOM.CreateFromPtr(pMonikerObject)                                                  'stdCOM<Interface>
    
    'Get ProgID from PathName
    Dim CLSID As GUID: CLSID.Data1 = 0: CLSID.Data2 = 0: CLSID.Data3 = 0: Erase CLSID.Data4
    Dim ProgID As String: ProgID = ""
    if sName like "!{*}" then
      Dim tCLSID as string: tCLSID = mid(sName, 2)
      Call IIDFromString(StrPtr(tCLSID), clsid)
      If ProgIDFromCLSID(clsid, VarPtr(pStr)) = S_OK then
        ProgID = GetStrFromPtrW(pStr)
      end if
    end if
    
    'Wrap return value in dictionary for easy enumeration with stdLambda
    Dim oDict As Object: Set oDict = CreateObject("Scripting.Dictionary")
    oDict("Name") = sName
    oDict("Type") = TypeName(obj.Object)
    oDict("ProgID") = ProgID
    Set oDict("COM") = obj
    ret.Add oDict
  Loop
  
  Set CreateFromActiveObjects = ret
End Function

'Create `stdCOM` object from WinRT Activatable ClassID.
'@constructor
'@param sActivatableClassId - An activatable WinRT class ID.
'@returns - A COM wrapper for the object.
'@example `stdCOM.CreateFromWinRTClassID("Windows.Storage.Pickers.FolderPicker")`
Public Function CreateFromWinRTClassID(ByVal sActivatableClassId as string) as stdCOM
  Dim hString as LongPtr
  if WindowsCreateString(StrPtr(sActivatableClassId), Len(sActivatableClassId), hString) <> S_OK then Err.Raise 1, "stdCOM::CreateFromWinRTClassID", "Cannot create HSTRING"
  if hString = NULL_PTR then Err.Raise 1, "stdCOM::CreateFromWinRTClassID", "Cannot create HSTRING"

  'Obtain stdCOM object
  Dim pIInspectable as LongPtr
  if RoActivateInstance(hString, pIInspectable) <> S_OK then Err.Raise 1, "stdCOM::CreateFromWinRTClassID", "Cannot activate class ID"
  if pIInspectable = 0 then Err.Raise 1, "stdCOM::CreateFromWinRTClassID", "Cannot activate class ID"
  set CreateFromWinRTClassID = CreateFromPtr(pIInspectable)

  'Delete string
  Call WindowsDeleteString(hString)
End Function


'Create `stdCOM` object from a CLSID and IID.
'@constructor
'@param sClassID - A CLSID as a string.
'@param sIID - An IID as a string. Default is `IDispatch`.
'@returns - A COM wrapper for the object.
'@example ```
'  stdCOM.CreateFromClassAndIID("FF48DBA4-60EF-4201-AA87-54103EEF594E","30CBE57D-D9D0-452A-AB13-7AC5AC4825EE") 'IUIAutomation
'```
Public Function CreateFromClassAndIID(ByVal sClassID As String, Optional ByVal sIID As String = "00020400-0000-0000-C000-000000000046") As stdCOM
  Dim clsid As GUID: If IIDFromString(StrPtr(sClassID), clsid) <> S_OK Then Err.Raise 1, "stdCOM::CreateFromClassAndIID", "Cannot retrieve CLSID from String."
  Dim iid As GUID: If IIDFromString(StrPtr(sIID), iid) <> S_OK Then Err.Raise 1, "stdCOM::CreateFromClassAndIID", "Cannot retrieve IID from String."
  Const CLSCTX_INPROC_SERVER = &H1

  Dim pIUnknown As LongPtr
  If CoCreateInstance(clsid, NULL_PTR, CLSCTX_INPROC_SERVER, iid, pIUnknown) <> S_OK Then Err.Raise 1, "stdCOM::CreateFromClassAndIID", "Cannot create object from CLSID and IID."
  
  Set CreateFromClassAndIID = CreateFromPtr(pIUnknown)
End Function


'#################################
'# INSTANCE METHODS / PROPERTIES #
'#################################

'Obtain raw object from stdCOM wrapper
'@returns - Raw object
'@example `stdCOM.Create(Application).Object.Name`
'@example `stdCOM.CreateFromPtr(ObjPtr(Application)).Object.Name`
Public Property Get Object() As IUnknown
  If this.Meta Is Nothing Then
    Set Object = Deref(This.ObjPtr)
  Else
    Set Object = this.Meta
  End If
End Property

'Obtain raw object from stdCOM wrapper
'@returns - `IDispatch` for object
'@example `stdCOM.Create(Application).Object.Name`
'@example `stdCOM.CreateFromPtr(ObjPtr(Application)).Object.Name`
'@throws {...} Unsupported interface IDispatch
Public Property Get AsObject() As Object
  If this.Meta Is Nothing Then
    Set AsObject = Deref(This.ObjPtr)
  Else
    Set AsObject = this.Meta
  End If
End Property

'Obtain pointer as LongPtr (VBA)/Long (VB6)
'@returns - Pointer to object
#If VBA7 Then
  Public Property Get Pointer() As LongPtr
#Else
  Public Property Get Pointer() As Long
#End If
  If this.Meta Is Nothing Then
    Pointer = This.ObjPtr
  Else
    Pointer = ObjPtr(this.Meta)
  End If
End Property

'Call `IUnknown::QueryInterface` on the object
'@param sIID - Interface Identifier as string
'@returns - `stdCOM` object for interface
'@throws stdCOM<IUnknown>#unkQueryInterface - Cannot retrieve IID from String.
'@throws stdCOM<IUnknown>#unkQueryInterface - Interface not implemented.
'@throws stdCOM<IUnknown>#unkQueryInterface - Unknown error occurred.
'@example ```
'  'Get ITypeInfo
'  set typeInfo = com.unkQueryInterface("00020401-0000-0000-C000-000000000046")
'```
Public Function unkQueryInterface(ByVal sIID As String) As stdCOM
  Dim IInterface As IUnknown
  Dim tGUID As GUID
  If IIDFromString(StrPtr(sIID), tGUID) = S_OK Then
    Dim hResult As Long: hResult = CallVT(EIUnknown.QueryInterface, vbLong, VarPtr(tGUID.Data1), VarPtr(IInterface))
    If hResult = S_OK Then
      if IInterface is nothing then 
        set unkQueryInterface = nothing
      else
        Set unkQueryInterface = stdCOM.Create(IInterface)
      end if
    ElseIf hResult = E_POINTER Then
      Set unkQueryInterface = Nothing
    ElseIf hResult = E_NOINTERFACE Then
      Err.Raise E_NOINTERFACE, "stdCOM<IUnknown>#unkQueryInterface", "Interface not implemented."
    Else
      Err.Raise hResult, "stdCOM<IUnknown>#unkQueryInterface", "Unknown error occurred."
    End If
  Else
    Err.Raise Err.LastDllError, "stdCOM<IUnknown>#unkQueryInterface", "Cannot retrieve IID from String."
  End If
End Function

'Call `IUnknown::AddRef()`
Public Sub unkAddRef()
  CallVT EIUnknown.AddRef, vbLong
End Sub

'Call IUnknown::Release()
Public Sub unkRelease()
  CallVT EIUnknown.Release, vbLong
End Sub

'Query the object for a known interface.
'@param interface - Known interface to query for
'@returns - interface queried.
'@throws "Interface IID not known."
Public Function QueryKnownInterface(ByVal interface As EKnownInterfaces) As stdCOM
  Dim sIID As String: sIID = KnownIID(interface)
  If sIID <> "" Then
    Dim oRet as stdCOM: Set oRet = unkQueryInterface(sIID)
    if not oRet is nothing then Call oRet.unkAddRef
    set QueryKnownInterface = oRet
  Else
    Err.Raise 1, "stdCOM::QueryKnownInterface", "Interface IID not known."
  End If
End Function


'Obtain Type Information from object
'@returns stdCOM<ITypeInfo> - TypeInfo structure for this object
'@example `myCom.TypeInfo.CallVT(EITypeInfo.GetTypeAttr, vbLong, VarPtr(lPtr))`
Public Property Get TypeInfo() As stdCOM
  On Error GoTo CannotQuery
    Set TypeInfo = QueryKnownInterface(ITypeInfo)
    If TypeInfo Is Nothing Then GoTo CannotQuery
    Exit Property
CannotQuery:
  On Error GoTo UnknownError
  'Attempt to query from IDispatch
  With QueryKnownInterface(IDispatch)
    Dim pTypeInfo As IUnknown
    Dim hResult As Long: hResult = .CallVT(EIDispatch.GetTypeInfo, vbLongPtr, NULL_PTR, NULL_PTR, VarPtr(pTypeInfo))
    If hResult <> S_OK Then Err.Raise hResult, "InterfaceInfo", "stdCOM::InterfaceInfo"
    if not pTypeInfo is nothing then Set TypeInfo = stdCOM.Create(pTypeInfo)
  End With
  Exit Property
UnknownError:
  Set TypeInfo = Nothing
End Property

'Obtain all properties and methods of an object via it's TypeInfo
'@returns Collection<Array<string, vbCallType, long, long>> - Collection of arrays containing Method/Property name, CallType, Index in type info and DispID
Public Property Get InterfaceInfo() As Collection
  Static pInterfaceInfo As Collection
  If pInterfaceInfo Is Nothing Then
    'Wrap pTypeInfo in stdCOM
    Dim oTypeInfo As stdCOM: Set oTypeInfo = TypeInfo()
    If oTypeInfo Is Nothing Then
        Set InterfaceInfo = New Collection
        Exit Property
    End If
    
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
    CopyMemory tTypeAttr, ByVal lPtr, LenB(tTypeAttr)
    
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

'Obtain a collection of all method names on an object
'@returns Collection<string> - Collection of method names
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

'Obtain a collection of all property names on an object
'@param bIncludeReadWriteType - Include GET/LET/SET/UNK prefixes representing call conv
'@returns Collection<string> - Collection of property names
'@example ```
'  for each vProp in stdCOM.Create(Application).Properties
'    Debug.Print vProp
'  next
'```
Public Property Get Properties(Optional bIncludeReadWriteType As Boolean = False) As Collection
  Static oProperties As Collection
  If oProperties Is Nothing Then
    Set oProperties = New Collection
    Dim info As Variant
    For Each info In InterfaceInfo
      If info(1) <> vbMethod Then
          Dim sReadWritePrefix As String
          If bIncludeReadWriteType Then
            Select Case info(1)
              Case VbGet: sReadWritePrefix = "GET"
              Case VbLet: sReadWritePrefix = "LET"
              Case VbSet: sReadWritePrefix = "SET"
              Case Else: sReadWritePrefix = "UNK"
            End Select
            oProperties.Add sReadWritePrefix & " " & info(0)
          Else
            On Error Resume Next
              oProperties.Add info(0), info(0)
            On Error GoTo 0
          End If
      End If
    Next
  End If

  Set Properties = oProperties
End Property

'Register the object held in stdCOM to the ROT as a CLSID or ProgID
'@param sCLSID - CLSID if required. If this arg isn't provided, a ProgID must be provided.
'@param ProgID - ProgID if required. If this arg isn't provided, a CLSID must be provided.
Public Sub RegisterActiveObject(Optional ByVal sCLSID As String = "", Optional ByVal ProgID As String = "")
    Call RegisterActiveObjectEx(Object, this.RegisteredOLEInstance, sCLSID, ProgID)
End Sub

'Revoke the previously registered object
Public Sub RevokeActiveObject()
    Call RevokeActiveObjectEx(Object, this.RegisteredOLEInstance)
End Sub

'############################
'# RAW METHODS / PROPERTIES #
'############################

'Call a function in the VTable of the object at a specified offset
'@param VTableOffset                         - The offset of the method to call. IUnknown::QueryInterface=0; IUnknown::AddRef=1; IUnknown::Release=2; IDispatch::Invoke=6
'@param ReturnType                           - Return value of function, typically Long (HResult) for COM Object calls
'@param FunctionParameters as Array<Variant> - Arguments to send to function.
'@returns                                    - Return value of function
'@throws {...} "DispCallFunc - Unknown error occurred."
'@example ```
'  Dim pITypeInfo As IUnknown
'  Dim hResult As Long: hResult = com.CallVT(EIUnknown.QueryInterface, vbLong, VarPtr(tGUID.Data1), VarPtr(pITypeInfo))
'  If hResult <> S_OK Then Err.Raise hResult, "InterfaceInfo", "stdCOM::InterfaceInfo"
'```
Public Function CallVT(ByVal VTableOffset As Long, ByVal ReturnType As VbVarType, ParamArray FunctionParameters() As Variant) As Variant
    'Create array from params
    Dim vParam() As Variant: vParam = FunctionParameters
    Dim iParamCount As Long: iParamCount = UBound(vParam) - LBound(vParam) + 1
    
    'Create array of types and pointers to vars, for function call
    #If VBA7 Then
      Dim vPtr() As LongPtr
    #Else
      Dim vPtr() As Long
    #End If
    Dim vType() As Integer

    'Populate array of pointers to params, and array of types for DispCallFunc
    If iParamCount = 0 Then
      ReDim vType(0 To 0)
      ReDim vPtr(0 To 0)
    Else
      ReDim vType(0 To iParamCount - 1)
      ReDim vPtr(0 To iParamCount - 1)
    
      Dim lIdx As Long
      For lIdx = 0 To iParamCount - 1
        vType(lIdx) = varType(vParam(lIdx))
        vPtr(lIdx) = VarPtr(vParam(lIdx))
      Next
    End If
    
    'Call function
    Dim hResult As Long
    hResult = DispCallFunc(This.ObjPtr, VTableOffset * PTRSIZE, CC_STDCALL, ReturnType, iParamCount, VarPtr(vType(0)), VarPtr(vPtr(0)), vv)
    
    If hResult < 0 Then
        Err.Raise hResult, "DispCallFunc", "DispCallFunc - Unknown error occurred."
    End If
End Function

'*************
'* PROTECTED *
'*************
'Initialise from object's IUnknown
'@protected
'@param obj - Object to create raw COM object from
Friend Sub protInitFromObject(ByVal obj As IUnknown)
  This.ObjPtr = ObjPtr(obj)
  Set this.Meta = obj
End Sub

'Initialise from pointer to object
'@protected
'@param ptr - Pointer to object
Friend Sub protInitFromPtr(ByVal ptr As LongPtr)
  This.ObjPtr = ptr
  Call unkAddRef
End Sub

'
'@destructor
Private Sub Class_Terminate()
  'Revoke active objects
  If this.RegisteredOLEInstance > 0 Then Call RevokeActiveObjectEx(Object(), this.RegisteredOLEInstance)
  If this.RegisteredTimerOLEInstance > 0 Then Call RevokeActiveObjectEx(Me, this.RegisteredTimerOLEInstance)
  
  'Call unknown release (prevent memory leaks)
  If this.Meta Is Nothing Then Call unkRelease
End Sub

'Register an object in the ROT
'@param obj         - Object to register to ROT
'@param OLEInstance - ROT Cookie. This variable will be populated with data.
'@param sCLSID      - CLSID if required. If this arg isn't provided, a ProgID must be provided.
'@param ProgID      - ProgID if required. If this arg isn't provided, a CLSID must be provided.
Private Sub RegisterActiveObjectEx(obj As IUnknown, ByRef OLEInstance As Long, Optional ByVal sCLSID As String = "", Optional ByVal ProgID As String = "")
    Const ACTIVEOBJECT_WEAK = 1

    'Get clsid whether it be from CLSID string or ProgID
    Dim clsid As GUID, result As Long
    If ProgID <> "" Then
        result = CLSIDFromProgID(StrPtr(ProgID), clsid)
        If result <> S_OK Then
          Select Case Hex(result)
            Case "800401F3" 'CO_E_CLASSSTRING
              Err.Raise 1, "", "The registered CLSID for the ProgID '" & ProgID & "' is invalid"
            Case "80040151" 'REGDB_E_WRITEREGDB
              Err.Raise 1, "", "An error occurred writing the CLSID to the registry. See Remarks https://docs.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-clsidfromprogid."
          End Select
        End If
    ElseIf sCLSID <> "" Then
        result = CLSIDFromString(StrPtr(sCLSID), clsid)
        If result <> S_OK Then
          Select Case Hex(result)
            Case "800401F3" 'CO_E_CLASSSTRING
              Err.Raise 1, "", "The CLSID was incorrectly formatted."
            Case "80040154" 'REGDB_E_CLASSNOTREG
              Err.Raise 1, "", "The CLSID corresponding to the class string was not found in the registry."
            Case "80040150" 'REGDB_E_READREGDB
              Err.Raise 1, "", "The registry could not be opened for reading."
              
          End Select
        End If
    Else
        Err.Raise 7, "No CLSID or ProgId"
    End If
    
    'Attempt to register object to ROT
    If RegisterActiveObject32(obj, clsid, ACTIVEOBJECT_WEAK, OLEInstance) <> S_OK Then Err.Raise Err.LastDllError, "Object couldn't be registered"
End Sub

'Revoke an object in the ROT
'@param obj         - Object to revoke from ROT
'@param OLEInstance - ROT Cookie. This variable will be populated with data.
Private Sub RevokeActiveObjectEx(obj As IUnknown, ByRef OLEInstance As Long)
    'Remove instance from ROT
    If OLEInstance <> 0 Then
      If RevokeActiveObject32(OLEInstance, 0) <> 0 Then
        Debug.Print "Could not revoke?"
      End If
    End If
    
    'Disconnect any remote process connections being maintained on behalf of the active object pointer.
    If CoDisconnectObject(obj, 0) <> 0 Then Debug.Print "Couldn't disconnect object " & TypeName(obj)
End Sub

'Obtain a new GUID
'@returns - GUID as string
Private Function getGUID() As String
  getGUID = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx"
  getGUID = Replace(getGUID, "y", Hex(Rnd() And &H3 Or &H8))
  Dim i As Long: For i = 1 To 30
    getGUID = Replace(getGUID, "x", Hex$(CLng(Rnd() * 15.9999)), 1, 1)
  Next
End Function

'Map between EKnownInterfaces and Interface IIDs
'@param iIID - Known interface to query for
'@returns - IID as string
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

'Deref pointer to string helper
'@param ptr - Pointer to string
'@returns - String
#If VBA7 Then
    Private Function GetStrFromPtrW(ByVal ptr As LongPtr) As String
#Else
    Private Function GetStrFromPtrW(ByVal ptr As Long) As String
#End If
    SysReAllocString VarPtr(GetStrFromPtrW), ptr
End Function


'Deref pointer to IUnknown helper
'@param memAddress - Pointer to IUnknown
'@returns - IUnknown
'@devRemark DO NOT RETURN `result`
#If VBA7 Then
  Private Function Deref(ByVal memAddress As LongPtr) As IUnknown
#Else
  Private Function Deref(ByVal memAddress As Long) As IUnknown
#End If
  Dim result As IUnknown
  CopyMemory result, ByVal VarPtr(memAddress), LenB(memAddress)
  Set Deref = result                             'Set and add ref
  ZeroMemory result, LenB(memAddress)
End Function

'IDEA: for ProgID based registration
'  RegisterProgID("...")
'  RevokeProgID()
'AHK Source
'===================
'RegisterIDs(CLSID, APPID)
'{
'  RegWrite, REG_SZ, HKCU, Software\Classes\%APPID%,, %APPID%
'  RegWrite, REG_SZ, HKCU, Software\Classes\%APPID%\CLSID,, %CLSID%
'  RegWrite, REG_SZ, HKCU, Software\Classes\CLSID\%CLSID%,, %APPID%
'}
'
'RevokeIDs(CLSID, APPID)
'{
'  RegDelete, HKCU, Software\Classes\%APPID%
'  RegDelete, HKCU, Software\Classes\CLSID\%CLSID%
'}

