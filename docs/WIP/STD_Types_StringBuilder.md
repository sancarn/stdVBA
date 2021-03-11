# `Std.Types.StringBuilder`

This class is used to build large VBA Strings. This includes features such as interpolation and employs the `DISPID_EVALUATE` attribute to provide a simple, easy to read and maintainable string building process.

## Usage Example:

### Test example:

```vb
Dim sb As Object
Set sb = StringBuilder.Create()
sb.JoinStr = "-"
sb.Str = "Start"
sb.TrimBehaviour = RTrim
sb.InjectionVariables.Add "@1", "cool"
sb.[This is a really cool multi-line    ]
sb.[string which can even include       ]
sb.[symbols like " ' # ! / \ without    ]
sb.[causing compiler errors!!           ]
sb.[also this has @1 variable injection!]
Test = sb.Str = "Start-This is a really cool multi-line-string which can even include-symbols like "" ' # ! / \ without-causing compiler errors!!-also this has cool variable injection!"
```

### Building HTML:

```vb
'IMPORTANT!!! Only Object (aka "IDispatch") can use square bracket syntax! Therefore must define sb as object!
Dim sb as Object
set sb = StringBuilder.Create()
sb.TrimBehaviour = RTrim

'Inject variables into string using the InjectionVariables dictionary:
sb.InjectionVariables.add "{this.handleChange}", handleChange
sb.InjectionVariables.add "{this.state.value}", state.value
sb.InjectionVariables.add "{this.getRawMarkup()}", getRawMarkup()

'Build string
sb.[<div className="MarkdownEditor">                 ]
sb.[  <h3>Input</h3>                                 ]
sb.[  <label htmlFor="markdown-content">             ]
sb.[    Enter some markdown                          ]
sb.[  </label>                                       ]
sb.[  <textarea                                      ]
sb.[    id="markdown-content"                        ]
sb.[    onChange="{this.handleChange}"               ]
sb.[    defaultValue="{this.state.value}"            ]
sb.[  />                                             ]
sb.[  <h3>Output</h3>                                ]
sb.[  <div                                           ]
sb.[    className="content"                          ]
sb.[    dangerouslySetInnerHTML={this.getRawMarkup()}]
sb.[  />                                             ]
sb.[</div>                                           ]
Call renderHTML(sb)
```
> Note: The default value of a `StringBuilder` object auto-evaluates to the propert `Str` property meaning that `renderHTML(sb)` means we are passing the entire interpolated string to the renderHTML function (assuming `renderHTML` has a definition like `Sub renderHTML(s as string).

# Properties and Descriptions

|Type    |Name                |Description|
|--------|--------------------|-----------|
|Property|`RawString`         |The underlying string before interpolation from InjectionVariables|
|Property|`JoinStr`           |The variable used to delimit each call to `StringBuilder::Append()`|
|Property|`TrimBehaviour`     | `NoTrim (Default)`, `LTrim`, `RTrim` or `Trim`
|Property|`InjectionVariables`|A [`Dictionary`](https://excelmacromastery.com/vba-dictionary/) of keys and values. The keys will be used as replacers, the values will be used as replacement.
|Property|`Str`               |This has `DISPID_VALUE` and thus is the default member of the class and is called automatically when `sb` is provided. `Str` will return the string build after full interpolation.
|Method  |`Create`            |Creates an instance of this object. Use this instead of `new`. This is mainly a standard used throughout my API, so you can take it or leave it.
|Method  |`Append`            |This has `DISPID_EVALUATE`. This appends the string passed to the `RawString` variable.
|Method  |`Test`              |A standard test created by me to make sure everything is working correctly.
