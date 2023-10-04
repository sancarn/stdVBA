# stdDLL

# Spec

* Can create with `Create()` passing path or shortened path for files in `C:\Windows\System32` directory.
* Can add function sigs using `addFunc`
* Function sigs are defined using cTypes
* Can run functions with `run` or `runUnsafe`. `run` will check that supplied params are correct before calling dll, casting if possible. `runUnsafe` will run function with parameters supplied as-is and may cause crashes if they are wrong.

```vb
set dll = stdDll.Create("user32")
Call dll.addFunc("MessageBoxA", ret:=dllTypeInt, dllTypeHWND, dllTypeLPCSTR, dllTypeLPCSTR, dllTypeUINT)
Debug.Print dll.run("MessageBoxA", 0, "Hello", "World", 0)
Debug.Print dll.runUnsafe("MessageBoxA", 0, "Hello", "World", 0)
```

* Can also add function sigs using `addCStubs` to add many functions in bulk using C stub code snippets.

```vb
Dim sSrc as string: sSrc = ""
sSrc = sSrc & vbCrLf & "int MessageBoxA("
sSrc = sSrc & vbCrLf & "  [in, optional] HWND   hWnd,"
sSrc = sSrc & vbCrLf & "  [in, optional] LPCSTR lpText,"
sSrc = sSrc & vbCrLf & "  [in, optional] LPCSTR lpCaption,"
sSrc = sSrc & vbCrLf & "  [in]           UINT   uType"
sSrc = sSrc & vbCrLf & ");"
sSrc = sSrc & vbCrLf & "HWND CreateWindowExA("
sSrc = sSrc & vbCrLf & "  [in]           DWORD     dwExStyle,"
sSrc = sSrc & vbCrLf & "  [in, optional] LPCSTR    lpClassName,"
sSrc = sSrc & vbCrLf & "  [in, optional] LPCSTR    lpWindowName,"
sSrc = sSrc & vbCrLf & "  [in]           DWORD     dwStyle,"
sSrc = sSrc & vbCrLf & "  [in]           int       X,"
sSrc = sSrc & vbCrLf & "  [in]           int       Y,"
sSrc = sSrc & vbCrLf & "  [in]           int       nWidth,"
sSrc = sSrc & vbCrLf & "  [in]           int       nHeight,"
sSrc = sSrc & vbCrLf & "  [in, optional] HWND      hWndParent,"
sSrc = sSrc & vbCrLf & "  [in, optional] HMENU     hMenu,"
sSrc = sSrc & vbCrLf & "  [in, optional] HINSTANCE hInstance,"
sSrc = sSrc & vbCrLf & "  [in, optional] LPVOID    lpParam"
sSrc = sSrc & vbCrLf & ");"
sSrc = sSrc & vbCrLf & "BOOL EnumChildWindows("
sSrc = sSrc & vbCrLf & "  [in, optional] HWND        hWndParent,"
sSrc = sSrc & vbCrLf & "  [in]           WNDENUMPROC lpEnumFunc,"
sSrc = sSrc & vbCrLf & "  [in]           LPARAM      lParam"
sSrc = sSrc & vbCrLf & ");"

set dll = stdDll.Create("user32")
Call dll.addCStubs(sSrc)
Debug.Print dll.run("MessageBoxA", 1) 'not sure if possible to optionally leave out `optional`
Debug.Print dll.run("MessageBoxA", 0, "Hello", "World", 1)

'Only way to get intellisense:
Debug.Print dll.exportVBXStubs
```

* Can export vbx stubs for intellisense purposes.

```vb
Public Function EnumChildWindows(ByVal hHWndParent as LongPtr, ByVal lpEnumFunc as LongPtr, ByVal lparam as Long) as Boolean
  static dll as stdDLL: if dll is nothing then set dll = stdDll.Create("user32") _ 
    .addFunc("EnumChildWindows", ret:=dllTypeBool, dllTypeHWND, dllTypeWndEnumProc, dllTypeLParam)
  EnumChildWindows = dll.run("EnumChildWindows", hHWNDParent, lpEnumFunc, lparam) = 1
End Function
```

* Casting of stdICallable into LongPtr
* Casting of parameters in safe call mode will look something like this:

```vb
Private Function safeCast(param, paramType) as variant
  select case paramType
    case dllTypeHandle, dllTypeHWND
      select case vartype(param)
        case vbLongPtr
          return param
        case else
          CriticalRaise "Parameter of incorrect type"
      end select
    case dllTypeLPCSTR
      select case vartype(param)
        case vbString
          return varptr(param)
        case vbLongPtr
          return param 
        case else
          CriticalRaise "Parameter of incorrect type"
      end select
    case dllTypeCallback, dllTypeWndEnumProc
      select case vartype(param)
        case vbObject
          if typeof param is stdICallable then
            return allocCallbackThunk(param)
          else
            CriticalRaise "Parameter of incorrect type"
          end if
        case vbLongPtr
          return param
        case else
          CriticalRaise "Parameter of incorrect type"
      end select
    case else
      CriticalRaise "Parameter type undefined"
  end select
end function
```



