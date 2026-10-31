# stdReg

## Creating keys and values

~~~vb
With stdReg.CreateFromKey("HKEY_CURRENT_USER\Software")
  With .addKey("MyApp")
    .value = "Open with My&App"
    With .addItem("icon")
      .value = "C:\Program Files\MyApp\MyIcon.ico"
    End With
    With .addKey("command")
      .value = """C:\Program Files\MyApp\MyApp.exe"" ""%1"""
    End With
  End With
End With
~~~

addKey accepts a nonempty relative subkey path and creates missing intermediate keys. An existing key and its values are preserved.

addItem creates a named value and returns a writable stdReg item. It defaults to an empty REG_SZ string. Empty names and "(Default)" select the unnamed value. Existing values retain their data and type even when a different creation type is requested.

~~~vb
Dim count As stdReg
Set count = stdReg.CreateFromKey("HKCU\Software\MyApp").addItem("Count", Value_DWORD)
count.value = 42
count.SetValue "forty-two", Value_String
~~~

Creation supports Value_String, Value_String_WithEnvVars, Value_DWORD, Value_Binary and Value_None. Other new-value types are rejected before writing; existing limitations of the class's other value encodings are unchanged.

## Default values

Reading or assigning value on a key accesses its unnamed value. Use SetValue(data, type) on a key to select the unnamed value's registry type explicitly.

## Access and naming

Creation observes the current user's Windows permissions. New wrappers inherit their parent's access mask and selected 32/64-bit registry view. Value writes acquire a write handle in that same view. Query and write handles are released when the wrapper is destroyed.

The class retains its existing ANSI registry API behavior. addItem rejects embedded nulls and backslashes because its path-based representation cannot distinguish a backslash in a value name from a subkey separator.

## Regression tests

stdRegTests.testAll exercises the chained example, unnamed/named strings, existing values, type changes, DWORDs, binary and expand strings. It compares stored data with WScript.Shell.RegRead and removes only its unique temporary HKCU\Software\stdVBA_RegistryCreation_* subtree.

