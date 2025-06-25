//Process description:

/* 
 1. Read all files in ../../src directory
 2. Append a function to the end of each files as follows:

```vb
Private Sub Err_Raise(ByVal number as Long, Optional ByVal source as string = "", Optional ByVal description as string = "")
  Call stdError.Raise(description)
End Sub
```
3. Replace all calls to `Err.Raise` to calls to `Err_Raise`
4. Perform the following mapping:

```vb
Public Function MyMethod(ByVal param1 as type1, ByVal param2 as type2, ...) as returnType
  ...
End Function
```

to

```vb
Public Function MyMethod(ByVal param1 as type1, ByVal param2 as type2, ...) as returnType
  With stdError.getSentry("MyMethod", "param1", param1, "param2", param2, ...)
    ...
  End With
End Function
```

5. Save and overwrite the files

This will be run on a every commit, and will publish to the `stdErrorWrapping` branch.
*/
















import { throws } from "assert";
import * as fs from "fs";
function main() {
  //Find all files in ../../src directory
  let files = fs.readdirSync(__dirname + "/../../src");
  files = files.filter((f) =>
    fs.lstatSync(__dirname + "/../../src/" + f).isFile()
  );

  //TBC
}

main();
