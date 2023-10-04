# stdEnumProvider

## Spec

The whole point of this class is to allow for object instances to easily implement `IEnumVariant` and be enumeratable.

```vb
class ArrayList
  Function GetEnum() as IUnknown
    if not isObject(stdEnumProvider) then Err.Raise 1, "ArrayList::GetEnum()", "This feature requires stdEnumProvider"
    set GetEnum = stdEnumProvider.CreateServer("protEnumProviderNext", "protEnumProviderReset", "protEnumProviderSkip").getIEnumVariant()
  End Function

  Public Function protEnumProviderNext(ByVal al as ArrayList, ByRef bFinished as boolean, ByVal oMeta as Object, ByVal params as variant) as Variant
    Call CopyVariant(protEnumProviderNext, al.items(oMeta("index")))
    oMeta("index") = oMeta("index") + 1
    bFinished = oMeta("index") > al.length
  End Function
  
  Public Function protEnumProviderReset(ByVal al as ArrayList, ByVal oMeta as Object, ByRef bFinished as boolean) as Variant
    oMeta("index") = 1
  End Function
  
  Public Function protEnumProviderSkip(ByVal al as ArrayList, ByRef bFinished as boolean, ByVal oMeta as Object, ByVal params as variant) as Variant
    oMeta("index") = oMeta("index") + params(0)
    bFinished = oMeta("index") > al.length
  End Function
end class
```

Provision of `Skip` and `Reset` are Optional. If not provided they'll error with not implemented error.

This class will work by constructing a VTable:

```vb
VTable(0) = FncPtr(AddressOf IUnknown_QueryInterface)
VTable(1) = FncPtr(AddressOf IUnknown_AddRef)
VTable(2) = FncPtr(AddressOf IUnknown_Release)
VTable(3) = FncPtr(AddressOf IEnumVARIANT_Next)
VTable(4) = FncPtr(AddressOf IEnumVARIANT_Skip)
VTable(5) = FncPtr(AddressOf IEnumVARIANT_Reset)
VTable(6) = FncPtr(AddressOf IEnumVARIANT_Clone)
```

But instead of using `AddressOf` we'll have to use an object identifier (e.g. Last private methods(?)).

N.B. `CreateServer` will have to lock itself (using `CoLockObjectExternal`) post-creation in order that the `IEnumVariant` object survives and can call to it. The object will unlock itself again when `bFinished` flag returns `true`, else the VBA runtime will crash.


## Hurdles

* Obtain pointer of function in class - need a good example of this.
  * Ensure it works in both 32 and 64 bit.
* Alternative: Use M-Code thunks.
