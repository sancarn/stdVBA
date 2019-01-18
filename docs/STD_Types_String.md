# `String`

Base class which not only contains useful string functions but also contains real interpolation functionality.

## Example

```vb
Debug.Print String.[This is my cool string! It can have quotes " and stuff in it too!]

Dim interpolation as string
interpolation = "magic!"
Debug.Print String.[It also has $(interpolation)]
```

## Interpolation Mechanism

### Background

First point to realise is that the variables cotnained within the current scope in VBA are stored in a table for each individual sub-routine / function:

```vb
Sub t()
    Dim a As Object: Set a = Nothing
    Dim b As Variant
    Dim c As String
    Dim d As Integer
    Dim e As Long

    'Memory Starts at     '31846740
    Debug.Print VarPtr(a) '31846736
    Debug.Print VarPtr(b) '31846720
    Debug.Print VarPtr(c) '31846716
    Debug.Print VarPtr(d) '31846714
    Debug.Print VarPtr(e) '31846708
    Debug.Print "-------------------------"
    Call t2
End Sub

Sub t2()
    Dim a As Object: Set a = Nothing
    Dim b As Variant
    Dim c As String
    Dim d As Integer
    Dim e As Long

    Debug.Print VarPtr(a) '31846556
    Debug.Print VarPtr(b) '31846540
    Debug.Print VarPtr(c) '31846536
    Debug.Print VarPtr(d) '31846534
    Debug.Print VarPtr(e) '31846528
End Sub
```

If we compare the numbers we can see the differing lengths of the data also:

```vb
    Dim a As Object  'Requires 6 bytes
    Dim b As Variant 'Requires 16 bytes
    Dim c As String  'Requires 4 bytes
    Dim d As Integer 'Requires 2 bytes
    Dim e As Long    'Requires 6 bytes
```

It is also important to note that the table that the data is found in is entirely dependant on the stack depth of the called function/sub i.e.

```vb
Call t()
```

Emits:

```
 31846736
 31846720
 31846716
 31846714
 31846708
-------------------------
 31846556
 31846540
 31846536
 31846534
 31846528
```

where as:

```vb
Call t2()
```

Emits:

```
 31846736
 31846720
 31846716
 31846714
 31846708
```

And as a more complete example:

```vb
Sub t0()
  Dim a as Integer
  Debug.Print VarPtr(a)
  Call t1()
End Sub
Sub t1()
  Dim a as Integer
  Debug.Print VarPtr(a)
  Call t2()
End Sub
Sub t2()
  Dim a as Integer
  Debug.Print VarPtr(a)
  Call t3()
End Sub
Sub t3()
  Dim a as Integer
  Debug.Print VarPtr(a)
  Call t4()
End Sub
Sub t4()
  Dim a as Integer
  Debug.Print VarPtr(a)
  Call t5()
End Sub
Sub t5()
  Dim a as Integer
  Debug.Print VarPtr(a)
End Sub
```

Emits:

```
 31846738
 31846586
 31846434
 31846282
 31846130
 31845978
```

And by calculating the difference between each table:

```vb
?31846738 - 31846586 '152
?31846586 - 31846434 '152
?31846434 - 31846282 '152
?31846282 - 31846130 '152
?31846130 - 31845978 '152
```

> Note: It is likely that each variable defined here is also just a pointer to the real variable, as strings for example can be longer than 152 characters long, Thus a fixed table size of 152 really wouldn't make sense otherwise.

> Note: This is also true between modules. I am yet to test between classes.


And thus it follows that:

```vb
'Stack depth is 0-based
Function varTablePointer(stackDepth as long) as Long
   varTablePointer = 31846740 - 152 * stackDepth
End Function
```

And conversly the inverse is true:

```vb
'Stack depth is 0-based
Function stackDepth() as Long
  Dim i as integer, varTablePointer as Long
  varTablePointer = VarPtr(i) + 2
  stackDepth = (31846740 - varTablePointer)/152
End Function
```

> Notice: `stackDepth()` function returns the depth of the `stackDepth` function, not of the function calling it. To get the caller's stack depth then `stackDepth()-1` is required.

Another important distinction that should be noticed is the difference in stack depth when calling the stackDepth function in different ways. Take the following module:

```vb
Sub printStackDepth()
  Debug.Print stackDepth()
End sub

'Stack depth is 0-based
Function stackDepth() as Long
  Dim i as integer, varTablePointer as Long
  varTablePointer = VarPtr(i) + 2
  stackDepth = (31846740 - varTablePointer)/152
End Function
```

When `stackDepth()` is called from the immediate window then the function returns `-1`.

When `stackDepth()` is called by running `printStackDepth()` sub from the VBE then the function returns `1`.

What is happening here? It is important to remember that `stackDepth()` is only really a measure of the offset in the stack from a certain position in table memory, not the measure of the stack itself depth itself. The fact that there is a `2` unit offset between `printStackDepth()` and calling the function from the immediate window however indicates to me that:

```
'When run button is clicked:
Run Button Pressed --> Application.Run "<name>" called --> Sub <name> called --> stackDepth() called (returning 1)

'When ran from immediate window calls `Sub <name>` directly:
Immediate window called --> Sub <name> called --> stackDepth() called (returning 0)

'When ran from immediate window calls `stackDepth()` directly:
Immediate window called --> stackDepth() called (returning -1)
```

### Getting names and types of variables

Currently I don't know any other way than to parse the code for these values. However I assume there will be a table of names somewhere as well, as 152 is really very little space to store data (10 variants would cause a stack overflow), thus either there is a garbage collector with a table of names and types OR there is an extended overflow stack.

### Implementation

Ultimately to get information about the previous stack `String.[]` has a call pattern akin to the following:

```vb
class String
  Static Function Create(str as string)
    Attribute Create.VB_UserMemId = -5
    Dim VBAVariableTable as STD_Types_VBAVariableTable
    Set VBAVariableTable = STD_Types_VBAVariableTable.Create(-1)
    For each key in VBAVariableTable.keys()
      replace(str,"#{" & key & "}",VBAVariableTable(key))
    next
  End Function
End Class
class VBAVariableTable
  Static Function Create(Optional ByVal stackOffset as integer = 0)
    Dim v as integer
    Dim tableStart as LongPtr
    tableStart = VarPtr(v) + 2

    Dim vbaTablePointer as LongPtr
    vbaTablePointer = tableStart - 152 * (stackOffset-1)

    'Determine stack pointer's name:
    '...

    'Determine variable names and types:
    '...

    'Build variable dictionary
    '...
  End Function
end class
class Pointer
  Static Function stackDepth() as LongPtr
    Dim i as integer, varTablePointer as Long
    varTablePointer = VarPtr(i) + 2
    stackDepth = (31846740 - varTablePointer)/152 - 1
  End Function
end class
```
