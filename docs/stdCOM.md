# stdCOM spec

# [General Reference](https://docs.microsoft.com/en-us/windows/desktop/api/_automat/)
* [RegisterActiveObject](https://docs.microsoft.com/en-us/windows/desktop/api/oleauto/nf-oleauto-registeractiveobject) - Registers an instance as the active object for a clsid.
* [RevokeActiveObject](https://docs.microsoft.com/en-us/windows/desktop/api/oleauto/nf-oleauto-revokeactiveobject)
* [RegisterTypeLibForUser](https://docs.microsoft.com/en-us/windows/desktop/api/oleauto/nf-oleauto-registertypelibforuser)

## GetActiveObjects(Optional ByVal Prefix as string ="", Optional ByVal CaseSensitive as Boolean = false) as Object

Returns a Dictionary of active COM objects, where each key is the item moniker or suffix of the object. If Prefix is specified, only objects whose item monikers match the given prefix are returned, and the prefix is omitted from the returned keys.

AHK example:

```ahk
GetActiveObjects(Prefix:="", CaseSensitive:=false) {
    objects := {}
    DllCall("ole32\CoGetMalloc", "uint", 1, "ptr*", malloc) ; malloc: IMalloc
    DllCall("ole32\CreateBindCtx", "uint", 0, "ptr*", bindCtx) ; bindCtx: IBindCtx
    DllCall(NumGet(NumGet(bindCtx+0)+8*A_PtrSize), "ptr", bindCtx, "ptr*", rot) ; rot: IRunningObjectTable
    DllCall(NumGet(NumGet(rot+0)+9*A_PtrSize), "ptr", rot, "ptr*", enum) ; enum: IEnumMoniker
    while DllCall(NumGet(NumGet(enum+0)+3*A_PtrSize), "ptr", enum, "uint", 1, "ptr*", mon, "ptr", 0) = 0 ; mon: IMoniker
    {
        DllCall(NumGet(NumGet(mon+0)+20*A_PtrSize), "ptr", mon, "ptr", bindCtx, "ptr", 0, "ptr*", pname) ; GetDisplayName
        name := StrGet(pname, "UTF-16")
        DllCall(NumGet(NumGet(malloc+0)+5*A_PtrSize), "ptr", malloc, "ptr", pname) ; Free
        if InStr(name, Prefix, CaseSensitive) = 1 {
            DllCall(NumGet(NumGet(rot+0)+6*A_PtrSize), "ptr", rot, "ptr", mon, "ptr*", punk) ; GetObject
            ; Wrap the pointer as IDispatch if available, otherwise as IUnknown.
            if (pdsp := ComObjQuery(punk, "{00020400-0000-0000-C000-000000000046}"))
                obj := ComObject(9, pdsp, 1), ObjRelease(punk)
            else
                obj := ComObject(13, punk, 1)
            ; Store it in the return array by suffix.
            objects[SubStr(name, StrLen(Prefix) + 1)] := obj
        }
        ObjRelease(mon)
    }
    ObjRelease(enum)
    ObjRelease(rot)
    ObjRelease(bindCtx)
    ObjRelease(malloc)
    return objects
}
```

## [Loading external assemblies?](https://docs.microsoft.com/en-us/windows/desktop/SbsCs/microsoft-windows-actctx-object)

VBScript example:

```vbs
var actCtx = WScript.CreateObject("Microsoft.Windows.ActCtx");
actCtx.Manifest = "myregfree.manifest";
var obj =  actCtx.CreateObject("MyObj");   
```

Not sure if this is possible to use or not...?

## Registering objects for COM interrop:

A register and unregister method which allows registering active COM servers.

```vb
Declare Private Function RegisterActiveObject lib "oleaut32.dll" (obj as any, clsid as string, dwFlags as long, handle as long)
Declare Private Function RevokeActiveObject   lib "oleaut32.dll" (handle as long, pvReserved as long)

Private COMDictionary As Object
Private Sub Class_Initialize()
  Set COMDictionary = CreateObject("Scripting.Dictionary")
End Sub

Sub Register(ByRef obj As Object, ByVal CLSID As String)
  Dim iHandle As Long
  RegisterActiveObject obj,CLSID, 0, iHandle
  COMDictionary(CLSID) = iHandle
End Sub

'Revoke active objects
Private Sub Class_Terminate()
  Dim key As Variant
  For Each key In COMDictionary.keys()
    Call RevokeActiveObject(COMDictionary(iHandle),0)
  Next
End Sub
```
