# cUniversalDLLCalls

This code can be used for creating classes which implement raw COM interfaces.

Code written by [LaVolpe](http://www.vbforums.com/member.php?92357-LaVolpe)

For more info see:

http://www.vbforums.com/showthread.php?781595-VB6-Call-Functions-By-Pointer-(Universall-DLL-Calls)

## From Post:

I was recently made aware that a function I've used from time to time for calling virtual functions of COM objects was perfectly adept at calling functions from just about any standard DLL out there. So, I whipped up a 'generic' class that can call both standard DLL functions & COM VTable functions. No thunks are used, just a couple of supporting API calls in the main class, including the low-level core API function: `DispCallFunc`

What does this mean for you? Well, it does allow you to call DLL functions from nearly 10 different calling conventions, including the two most common: StdCall & CDecl. It also allows you to call virtual functions from COM objects. And if you wish, it means you do not have to declare a single API function declaration in your VB project. Though, personally, I'd use it for calling DLL conventions other than StdCall.

I'd consider this topic advanced for one reason only. This is very low level. If you provide incorrect parameter information to the class, your project is likely to crash. For advanced coders, we have no problem doing the research to understand what parameter information is required, be it variable type, a pointer, a pointer to a pointer, function return types, etc, etc. Not-so-advanced coders just want to plug in values & play, but when playing at such a low level, that usually results in crashes and frustration.

The attachment includes very simple examples of calling DLL functions and calling a COM virtual function. You will notice that the form has no API function declarations, though several DLL functions are called & executed correctly. A sample call to the class might look like:

```vb
Debug.Print myClass.CallFunction_DLL("user32.dll", "IsWindowUnicode", STR_NONE, CR_LONG, CC_STDCALL, Me.hWnd)
```

For DLL calls, the class takes the DLL name and function name to be called. Technically, you aren't passing the function pointer to the class. However, the class does make the call to the pointer, not via declared API functions. Just thought I'd throw this comment in, should someone suggest we aren't really calling functions by pointer. The class is, the user calling the class is not, but can be if inclined to modify the code a bit.

### Limitations: Callbacks from non-StdCall DLLs/functions

If whatever function you are calling requires a callback pointer, then stack corruption is likely from all calling conventions where you pass a VB function pointer as the callback address. The exceptions are stdCall DLLs and also CDecl calls, if the thunk/patch option in the class is used. 

### One class per dll

If you really like this class, you may want to instantiate one for each DLL you will be calling quite often. This could speed things up a bit when making subsequent calls. As is, the class will load the requested DLL into memory if it isn't already. Once class is called again, for a different DLL, then the previous DLL is unloaded if needed & the new DLL loaded as needed. So, if you created `cUser32`, `cShell32`, `cKernel32` instances, less code is executed in the class if it doesn't have to drop & load DLLs.

```vb
' top of form
Private cUser32 As cUniversalDLLCalls
Private cKernel32 As cUniversalDLLCalls
Private cShell32 As cUniversalDLLCalls
 
' in form load
Set cUser32 = New cUniversalDLLCalls
Set cKernel32 = New cUniversalDLLCalls
Set cShell32 = New cUniversalDLLCalls
```

### String parameters

When using the STR_ANSI flag to indicate the passed parameters include string values destined for ANSI functions, the class will convert the passed string to ANSI before calling the function. Doing so, default Locale is used for string conversion. If this is a problem, you should ensure you convert the string(s) to ANSI before passing it to the class. If you do this conversion, use STR_NONE & pass the string via StrPtr(). FYI: strings used strictly as a buffer for return values should always be passed via StrPtr() and the flag STR_NONE used; regardless if destined for ANSI or unicode functions. ANSI strings are never passed to COM interfaces. Always use StrPtr(theString) for any string parameters to those COM methods.

```vb
' how to have a VB string contain ANSI vs Unicode
myString = StrConv(myString, vbFromUnicode, [Locale ID])
' how to convert the returned ANSI string to a proper VB string
myString = StrConv(myString, vbUnicode, [Locale ID])
```

### Calling private COM functions

If you ever need to call a private COM interface function by its pointer/address, post [#24](http://www.vbforums.com/showthread.php?781595-VB6-Call-Functions-By-Pointer-(Universall-DLL-Calls)&p=4818005&viewfull=1#post4818005) shows how that can be done. A slight modification to the attached class is required.

```vb
Debug.Print myClass.CallFunction_COM (0&, [private function address], STR_NONE, CR_LONG, CC_STDCALL, ObjPtr(ComInterface), [any private method parameters])
```

### Getting an IID

```vb
Private Declare Function IIDFromString Lib "ole32.dll" (ByVal lpszProgID As Long, piid As Any) As Long
' sample of getting the IDataObject GUID
Dim aGUID(0 To 3) As Long
Call IIDFromString(StrPtr("{0000010e-0000-0000-C000-000000000046}"), ByVal VarPtr(aGUID(0)))
```

### Query Interface:

```vb
Const IUnknownQueryInterface As Long = 0&   ' IUnknown vTable offset to Query implemented interfaces
Const IUnknownRelease As Long = 8&          ' IUnknown vTable offset to decrement reference count

'Get IID of IPicture
Dim IID_IPicture As Long, aGUID(0 To 3) As Long
Call IIDFromString(StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), ByVal VarPtr(aGUID(0)))

'ask if objToTest supports IPicture
Dim c As cUniversalDLLCalls: Set c = New cUniversalDLLCalls
c.CallFunction_COM ObjPtr(objToTest), IUnknownQueryInterface, CR_LONG, CC_STDCALL, VarPtr(aGUID(0)), VarPtr(IID_IPicture)
If IID_IPicture <> 0& Then
    ' do stuff
    ...

    ' Release the IPicture interface at some point. QueryInterface calls AddRef internally
    c.CallFunction_COM IID_IPicture, IUnknownRelease, CR_LONG, CC_STDCALL
End If
```

### Interfaces:

[IUnknown](https://docs.microsoft.com/en-us/windows/desktop/api/unknwn/nn-unknwn-iunknown): GUID {00000000-0000-0000-C000-000000000046}
VTable Order: QueryInterface, AddRef, Release

[IPicture](https://docs.microsoft.com/en-us/windows/desktop/api/ocidl/nn-ocidl-ipicture): GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
VTable Order: GetHandle, GetHPal, GetType, GetWidth, GetHeight, Render, SetHPal, GetCurDC, 
SelectPicture, GetKeepOriginalFormat, SetKeepOriginalFormat, PictureChanged, SaveAsFile, GetAttributes

[IDataObject](https://docs.microsoft.com/en-us/windows/desktop/api/objidl/nn-objidl-idataobject): GUID {0000010e-0000-0000-C000-000000000046}
VTable Order: GetData, GetDataHere, QueryGetData, GetCanonicalFormatEtc, SetData, 
EnumFormatEtc, DAdvise, DUnadvise, EnumDAdvise

#### Get the IDataObject from the Data parameter of VB's OLEDrag[...] events

```vb
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)

Dim IID_DataObject As Long
CopyMemory IID_DataObject, ByVal ObjPtr(Data) + 16&, 4&
' you now have an unreferenced pointer to the IDataObject
```

#### Get IDataObject of the clipboard

```vb
Private Declare Function OleGetClipboard Lib "ole32.dll" (ByRef ppDataObj As Long) As Long

Dim IID_DataObject As Long
OleGetClipboard IID_DataObject
' if IID_DataObject is non-null, you have a referenced pointer to the IDataObject
' Referenced pointers must call IUnknown.Release
```

[IOLEObject](http://msdn.microsoft.com/en-us/library/windows/desktop/dd542709%28v=vs.85%29.aspx): GUID {00000112-0000-0000-C000-000000000046}
VTable Order: SetClientSite, GetClientSite, SetHostNames, Close, SetMoniker, GetMoniker, 
InitFromData, GetClipboardData, DoVerb, EnumVerbs, Update, IsUpToDate, GetUserClassID, GetUserType, SetExtent, GetExtent, Advise, EnumAdvise, GetMiscStatus, SetColorScheme

[IStream](http://msdn.microsoft.com/en-us/library/windows/desktop/aa380034%28v=vs.85%29.aspx): inherits IUnknown:ISequentialStream. GUID {0000000C-0000-0000-C000-000000000046}
VTable Order: Read [from ISequentialStream], Write [from ISequentialStream], Seek, SetSize, 
CopyTo, Commit, Revert, LockRegion, UnlockRegion, Stat, Clone

[ITypeLib](http://msdn.microsoft.com/en-us/library/windows/desktop/ms221549%28v=vs.85%29.aspx): GUID {00020402-0000-0000-C000-000000000046}
VTable Order: GetTypeInfoCount, GetTypeInfo, GetTypeInfoType, GetLibAttr, 
GetTypeComp, GetDocumentation, IsName, FindName, ReleaseTLibAttr

[ITypeInfo](http://msdn.microsoft.com/en-us/library/windows/desktop/ms221696%28v=vs.85%29.aspx): GUID {00020401-0000-0000-C000-000000000046}
VTable Order: GetTypeAttr, GetTypeComp, GetFuncDesc, GetVarDesc, GetNames, 
GetRefTypeOfImplType, GetImplTypeFlags, GetIDsOfNames, Invoke, GetDocumentation, GetDLLEntry, GetRefTypeInfo, AddressOfMember, CreateInstance, GetMops, GetContainingTypeLib, ReleaseTypeAttr, ReleaseFuncDesc, RelaseVarDesc

### Implement VTable Orders:

An easy method to implement VTable order is using enums:

```vb
Enum EITypeLib
  IUnk_QueryInterface
  IUnk_AddRef
  IUnk_Release

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
  GetDLLEntry
  GetRefTypeInfo
  AddressOfMember
  CreateInstance
  GetMops
  GetContainingTypeLib
  ReleaseTypeAttr
  ReleaseFuncDesc
  RelaseVarDesc
End Enum
```

Now when calling e.g. `GetFuncDesc` you can easily do so like:

```vb
c.CallFunction_COM ObjPtr(objToTest), EITypeLib.GetFuncDesc, ...
```