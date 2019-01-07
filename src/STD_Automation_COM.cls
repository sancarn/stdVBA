VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "STD_Automation_COM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'OTHERS
'--------
'GetOpenFileNameA
'GetSaveFileNameA
'Using EbGetExecutingProj from msvbvm60.dll  - http://www.vbforums.com/showthread.php?799593&p=4901797#post4901797
'http://www.vbforums.com/showthread.php?254279-Return-an-object(late-bind)-by-giving-an-object-pointer-How
'https://stackoverflow.com/questions/1667397/accessing-dynamically-loaded-dll-with-loadlibrary-in-visual-basic-6
'https://github.com/Lexikos/ActiveScript.ahk/blob/master/ActiveScript.ahk

'EXAMPLES
'---------
''Get Excel Instance: (not sure why they use currency here as a GUID?):
''**********************************************************************
'Public Function GetExcelInstance(ByRef FileName As String) As Object
'  Const IID_DISPATCH = "{00020400-0000-0000-C000-000000000046}"
'  Dim IID(1) As Currency
'  IIDFromString StrPtr(IID_DISPATCH), IID(0&)
'  CoGetObject StrPtr(FileName), ByVal 0&, IID(0&), GetExcelInstance
'End Function
'
''Create COM Object on remote machine:
''**********************************************************************
''http://exceldevelopmentplatform.blogspot.com/2017/04/createobjectex-vba-to-create-remote-com.html
'---------

'GUID type for CLSID and IID
Private Type Guid
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(0 To 7) As Byte
End Type

'CLSIDFromProgID
'https://docs.microsoft.com/en-us/windows/desktop/api/combaseapi/nf-combaseapi-clsidfromprogid
Private Declare Function CLSIDFromProgID Lib "ole32.dll" (ByVal TSzProgID As String, ByVal t As Guid) As Long

'CLSIDFromString
'https://docs.microsoft.com/en-us/windows/desktop/api/combaseapi/nf-combaseapi-clsidfromstring
Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszProgID As Long, pCLSID As Guid) As Long

'IIDFromString
'https://docs.microsoft.com/en-us/windows/desktop/api/combaseapi/nf-combaseapi-iidfromstring
Private Declare Function IIDFromString Lib "ole32" (ByVal lpszIID As Long, IID As Any) As Long
 
'ProgIDFromCLSID
'https://docs.microsoft.com/en-us/windows/desktop/api/combaseapi/nf-combaseapi-progidfromclsid
Private Declare Function ProgIDFromCLSID Lib "ole32.dll" (pCLSID As Guid, lpszProgID As LongPtr) As Long

'StringFromCLSID
'https://docs.microsoft.com/en-us/windows/desktop/api/combaseapi/nf-combaseapi-stringfromclsid
Private Declare Function StringFromCLSID Lib "ole32.dll" (pCLSID As Any, lpszProgID As Long) As Long

'StringFromIID
'https://docs.microsoft.com/en-us/windows/desktop/api/combaseapi/nf-combaseapi-stringfromiid

'StringFromGUID2
'https://docs.microsoft.com/en-us/windows/desktop/api/combaseapi/nf-combaseapi-stringfromguid2

'DllGetClassObject
'https://docs.microsoft.com/en-us/windows/desktop/api/combaseapi/nf-combaseapi-dllgetclassobject

'CoLoadLibrary:
'https://docs.microsoft.com/en-us/windows/desktop/api/objbase/nf-objbase-coloadlibrary


'CoRegisterChannelHook
'https://docs.microsoft.com/en-us/windows/desktop/api/objbase/nf-objbase-coregisterchannelhook
Private Declare Function CoRegisterChannelHook Lib "ole32.dll" (ExtensionUUID As Guid, pChannelHook As LongPtr) As Long

'CoTreatAsClass
'https://docs.microsoft.com/en-us/windows/desktop/api/objbase/nf-objbase-cotreatasclass
'Converts an object into a class that can be instantiated
'Also allows for putty hacking of COM CLSIDs

'CoGetClassObject
'https://docs.microsoft.com/en-us/windows/desktop/api/combaseapi/nf-combaseapi-cogetclassobject
'Call CoGetClassObject directly to create multiple objects through a class object for which there is a CLSID in the system registry


'CreateBindCtx
'https://docs.microsoft.com/en-us/windows/desktop/api/objbase/nf-objbase-createbindctx


'CreateFileMoniker
'https://docs.microsoft.com/en-us/windows/desktop/api/objbase/nf-objbase-createfilemoniker

'CoCreateGuid
'https://docs.microsoft.com/en-us/windows/desktop/api/combaseapi/nf-combaseapi-cocreateguid

'GetClassFile
'https://docs.microsoft.com/en-us/windows/desktop/api/objbase/nf-objbase-getclassfile
'Returns the CLSID associated with the specified file name.

'CoGetInstanceFromFile
'https://docs.microsoft.com/en-us/windows/desktop/api/objbase/nf-objbase-cogetinstancefromfile
'Creates a new object and initializes it from a file using IPersistFile::Load.

'CoCreateInstance
'https://docs.microsoft.com/en-us/windows/desktop/api/combaseapi/nf-combaseapi-cocreateinstance
'Call CoCreateInstance when you want to create only one object on the local system. To create multiple objects based on a single CLSID, call the CoGetClassObject function.

'CoGetObject
'Converts a display name into a moniker that identifies the object named, and then binds to the object identified by the moniker.
'Not sure if this makes a call to MkParseDisplayName or not.
Private Declare Function CoGetObject Lib "ole32.dll" (ByVal pszName As Long, ByVal pBindOptions As Long, ByRef riid As Guid, ByVal ppv As Long) As Long

'GetRunningObjectTable
'https://docs.microsoft.com/en-us/windows/desktop/api/objbase/nf-objbase-getrunningobjecttable


'MkParseDisplayName
'https://docs.microsoft.com/en-us/windows/desktop/api/objbase/nf-objbase-mkparsedisplayname
'This function is the inverse of the IMoniker::GetDisplayName operation, which retrieves the display name associated with a moniker.

'CLSIDFromProgIDEx
'https://docs.microsoft.com/en-us/windows/desktop/api/objbase/nf-objbase-clsidfromprogidex

'CoGetCurrentProcess
'https://docs.microsoft.com/en-us/windows/desktop/api/combaseapi/nf-combaseapi-cogetcurrentprocess

'CoRegisterClassObject
'
