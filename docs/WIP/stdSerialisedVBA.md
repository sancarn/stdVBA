# SerialisedVBA:

General structure:

```vb
  Public memory as Dictionary 'stores all objects with GUID keys
  Public Function eval(input as string):...:End Function
  Public Sub evalAsync(input as string):...:End Function
```

## `eval(input)` / `evalAsync(input)`

Where input is:

```json
{
  "parent"  :"<<GUID-IF-EXISTS>>",
  "callType":"Get/Let/Set/Invoke",
  "args":[...]
}
```

### Return:

```json
{
  "value"  :"<<VALUE>>/<<GUID-IF-EXISTS>>",
  #IF isObject(Value)  'Use GetTypeInfo()
  "interface":[  
    {"type":"property","access":"r/w/rw"},
    {"type":"method", "args":[{"type":"string/double/...","optional":true/false}], "retType":"string/double/..."},
  ]
  #END IF
}
```
