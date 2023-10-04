# stdObject

## Spec

```vb
Dim o as object: set o = stdObject.Create("Key","Value","Hello",1,"Marge", stdObject.Create("Lisa","Hello"))
Debug.Print o.key        'Value
Debug.Print o.hello      '1
Debug.print o.Marge.Lisa 'hello
Call o.setField("poop","something")
Debug.Print o.poop       'something
```

Implements IDispatch interface with specified layout. I believe VBA calls GetIDsOfNames continually, every time a function is called, so can have a dynamic object in theory.