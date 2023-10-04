# JSON Candidates

Base code is quite important here are a few candidates which might be useful while creating stdJSON:

## [JSONBag by Dilettante](http://www.vbforums.com/showthread.php?738845-VB6-JsonBag-Another-JSON-Parser-Generator)

### Pros

#### API

Really handy feature of this library is the amazing use of the evaluate function:

```
json = "{""web-app"":{""servlet"":[{""init-param"":{""templateProcessorClass"":""MyClass""}}]}"

Debug.Print myBag![web-app]![servlet](1)![init-param]![templateProcessorClass] 'Prints MyClass
```

Does offer a neat API for producing JSON objects:

```vb
With JB
    .Clear
    .IsArray = False 'Actually the default after Clear.

    ![First] = 1
    ![Second] = Null
    With .AddNewArray("Third")
        .Item = "These"
        .Item = "Add"
        .Item = "One"
        .Item = "After"
        .Item = "The"
        .Item = "Next"
        .Item(1) = "*These*" 'Should overwrite 1st Item, without moving it.

        'Add a JSON "object" to this "array" (thus no name supplied):
        With .AddNewObject()
            .Item("A") = True
            !B = False
            !C = 3.14E+16
        End With
    End With
    With .AddNewObject("Fourth")
        .Item("Force Case") = 1 'Use quoted String form to force case of names.
        .Item("force Case") = 2
        .Item("force case") = 3

        'This syntax can be risky with case-sensitive JSON since the text is
        'treated like any other VB identifier, i.e. if such a symbol ("Force"
        'or "Case" here) is already defined in the language (VB) or in your
        'code the casing of that symbol will be enforced by the IDE:

        ![Force Case] = 666 'Should overwrite matching-case named item, which
                            'also moves it to the end.
        'Safer:
        .Item("Force Case") = 666
    End With
    'Can also use implied (default) property:
    JB("Fifth") = Null

    txtSerialized.Text = .JSON
End With
```

This means it is very easy to add stdIJSON implementation to custom classes/objects.

### Cons

* No automatic serialization to dictionaries and arrays/collections. This can likely be easily added.
* No iterator methods

## [VBA-JSON by Tim Hall](https://github.com/VBA-tools/VBA-JSON)

### Pros

#### Actively developed

This is actively developed which can't be said for JSONBags

#### Built for VB6

Fits in neatly with VBA, producing Dictionaries and Collections. Also consumes dictionaries and collections 

### Cons

#### API

* No Easy API to build JSON.
* No iterator methods


Ideally a combination of the 2 classes would be ideal.