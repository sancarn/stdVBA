This file will probably have a lot of cross over to other files, but will remain as a dumping ground for ideas of features for DLL related calls.

## Paramater type and struct

DLL calls are easy once you've called them a few times, but what really is frustrating is the unknown definition of [Win32 datatypes](https://docs.microsoft.com/en-us/windows/desktop/winprog/windows-data-types) e.g. `DWORD`,`DWORD32`,`DWORD64`,`DWORDLONG`, ... , `HANDLE`, `HBITMAP`, `HCURSOR`, `HDC`, `LONGLONG`, ... In all cases these are representable in VBA using Byte Arrays `Byte()`, but you still need to know the size of each parameter.

An idea is to have a parser which parses the `typedef`. E.G. 

```vb
'LONG SetWindowLongA(
'  HWND hWnd,
'  int  nIndex,
'  LONG dwNewLong
');

with stdDLL
  Debug.Print .call("User32","SetWindowLongA", _ 
    .type("HWND", hWnd), _ 
    .type("int", nIndex), _ 
    .type("LONG", dwNewLong), _ 
    .rtype("LONG")
  )
end with
```

Each parameter when parsed is converted to a byte array, casted to the correct type and then used as the parameter to call the dll function.

Structs should also be usable directly:

```vb
Dim lpRect as stdStruct
lpRect.add("LONG",0,"left")
lpRect.add("LONG",0,"top")
lpRect.add("LONG",0,"right")
lpRect.add("LONG",0,"bottom")

With stdDLL
  Debug.Print .call("User32", "GetWindowRect", _ 
    .type("HWND", hWnd), _ 
    .type("LPRECT", lpRect), _
    .rtype("BOOL")
  )
End With
```

Enums should also be usable directly: (not sure if this is howthey will work)

```vb
Dim GWL as stdEnum
GWL.add("GWL_EXSTYLE"  ,-20)
GWL.add("GWL_HINSTANCE",-6)
GWL.add("GWL_ID"       ,-12)
GWL.add("GWL_STYLE"    ,-16)
'...
GWL = "GWL_EXSTYLE"


with stdDLL
  Debug.Print .call("User32","SetWindowLongA", _ 
    .type("HWND", hWnd), _ 
    .type("int", GWL), _ 
    .type("LONG", dwNewLong), _ 
    .rtype("LONG")
  )
end with
```
