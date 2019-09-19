
//these bits of IDL were taken from Edmundos olelib.tlb
//

typedef [public] LONG BOOL;
typedef unsigned char BYTE;

typedef struct UUID {
    LONG Data1;
    SHORT Data2;
    SHORT Data3;
    BYTE Data4[8];
} UUID;

typedef enum HRESULTS {
    S_OK = 0,
    S_FALSE = 1,

    E_NOTIMPL = 0x80004001,
    E_OUTOFMEMORY = 0x8007000E,
    E_INVALIDARG = 0x80070057,
    E_NOINTERFACE = 0x80004002,
    E_POINTER = 0x80004003,
    E_HANDLE = 0x80070006,
    E_ABORT = 0x80004004,
    E_FAIL = 0x80004005,
    E_ACCESSDENIED = 0x80070005,
} HRESULTS;

typedef enum CLSCTX {
    CLSCTX_INPROC_SERVER = 0x1,
    CLSCTX_INPROC_HANDLER = 0x2,
    CLSCTX_LOCAL_SERVER = 0x4,
    CLSCTX_INPROC_SERVER16 = 0x8,
    CLSCTX_REMOTE_SERVER = 0x10,
    CLSCTX_INPROC_HANDLER16 = 0x20,
    CLSCTX_INPROC_SERVERX86 = 0x40,
    CLSCTX_INPROC_HANDLERX86 = 0x80,
    CLSCTX_ESERVER_HANDLER = 0x100,
    CLSCTX_RESERVED = 0x200,
    CLSCTX_NO_CODE_DOWNLOAD = 0x400,
    CLSCTX_NO_WX86_TRANSLATION = 0x800,
    CLSCTX_NO_CUSTOM_MARSHAL = 0x1000,
    CLSCTX_ENABLE_CODE_DOWNLOAD = 0x2000,
    CLSCTX_NO_FAILURE_LOG = 0x4000
} CLSCTX;

typedef [helpstring("DISP_E_EXCEPTION Error info")] struct EXCEPINFO {
        [helpstring("Error number")]
        short wCode;
        short wReserved;
        [helpstring("Source")]
        BSTR Source;
        [helpstring("Description")]
        BSTR Description;
        [helpstring("Help file")]
        BSTR HelpFile;
        [helpstring("Help context")]
        BSTR dwHelpContext;
        LONG pvReserved;
        [helpstring("HRESULT (__stdcall *pfnDeferredFillIn)(struct EXCEPINFO *);")]
        LONG pfnDeferredFillIn;
        LONG scode;
} EXCEPINFO;

typedef struct DISPPARAMS {
        [helpstring("Parameters")]
        LONG rgPointerToVariantArray;
        [helpstring("Named parameters")]
        LONG rgPointerToLONGNamedArgs;
        [helpstring("Parameter count")]
        LONG cArgs;
        [helpstring("Named parameter count")]
        LONG cNamedArgs;
} DISPPARAMS;

[
    odl,
    uuid(00000000-0000-0000-C000-000000000046),
]
interface IUnknown{

    LONG QueryInterface(
        [in, out] UUID *riid,
        [in, out] void *ppvObject);

    LONG AddRef();
    LONG Release();
}
 


    [
      odl,
      uuid(00020400-0000-0000-C000-000000000046),
    ]
    interface IDispatch : IUnknown {

        HRESULT GetTypeInfoCount(
            [out, retval] int* pctinfo);

        HRESULT GetTypeInfo(
            [in, defaultvalue(0)] int itinfo,
            [in, defaultvalue(0)] long lcid,
            [out, retval] LONG **pptinfo);

        LONG GetIDsOfNames(
            [in] UUID* riid,
            [in] LPWSTR *rgszNames,
            [in] int cNames,
            [in] long lcid,
            [out] long *rgdispid);

        LONG Invoke(
            [in] long dispidMember,
            [in] UUID* riid,
            [in] long lcid,
            [in] short wFlags,
            [in] DISPPARAMS *pdispparams,
            [in] long pvarResult,
            [out] EXCEPINFO *pexcepinfo,
            [out] int *puArgErr);
    };




[
    dllname("OLE32.DLL"),
    helpstring("Storage functions")
]
module ole32 {

    [entry("CoCreateInstance")]
    LONG CoCreateInstance(
        [in] UUID *CLSID,
        [in] stdole.IUnknown *pUnkOuter,
        [in] CLSCTX dwClsContext,
        [in] UUID *IID,
        [out] void *ppv);


    [entry("CLSIDFromString")]
    HRESULT CLSIDFromString(
        [in] LPWSTR lpszProgID,
        [in, out] UUID *lpclsid);

    [entry("CLSIDFromProgID")]
    HRESULT CLSIDFromProgID(
        [in] LPWSTR lpszProgID,
        [in, out] UUID *lpclsid);
}

[
    dllname("KERNEL32.DLL")
]
module kernel32 {
    [entry("RtlMoveMemory")]
    void MoveMemory(
        [in] void *pDest,
        [in] void *pSource,
        [in] LONG ByteLen);

    [entry("RtlMoveMemory")]
    void MoveMemoryToLPSTR(
        [in] LPSTR pDest,
        [in] void *pSource,
        [in] LONG ByteLen);

    [entry("RtlMoveMemory")]
    void MoveMemoryFromLPSTR(
        [in] void *pDest,
        [in] LPSTR pSource,
        [in] LONG ByteLen);

    [entry("lstrlenW")]
    LONG lstrlenW (
        [in] LONG StringPtr);

    [entry("lstrlenA")]
    LONG lstrlenA (
        [in] LONG StringPtr);

    [entry("lstrcpyA")]
    LONG lstrcpyA (
        [in] LPSTR Dest,
        [in] LONG Src);
}


[
	dllname("USER32.DLL")
]
module user32 {

	[entry("SendMessageA")]
	LONG SendMessage(
                [in] LONG hWnd,
                [in] LONG wMsg,
                [in] LONG wParam,
                [in] void *lParam);

	[entry("SendMessageW")]
	LONG SendMessageW(
                [in] LONG hWnd,
                [in] LONG wMsg,
                [in] LONG wParam,
                [in] void *lParam);

}