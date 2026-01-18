VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdJSON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@module
'@description A class which implements a JSON object. Use this class to
'create JSON objects, manipulate them, and convert them to and from JSON strings.
'@example Simple example of creating a JSON object and converting it to a string:
'```vb
'Dim o as stdJSON: set o = stdJSON.Create()
'o.Add "a", 1
'o.Add "b", "hello"
'o.AddArray("val", Array(1,2,3))
'o.add "serialized", stdJSON.CreateFromString("{""a"":1,""b"":""hello""}")
'Debug.Print o.ToString()
'```
'This will output:
'```json
'{
'  "a": 1,
'  "b": "hello",
'  "val": [
'    1,
'    "hello"
'  ],
'  "serialized": {
'    "a": 1,
'    "b": "hello"
'  }
'}
'```
'@example Simple example of creating a JSON object from a file:
'```vb
'Dim o as stdJSON: set o = stdJSON.CreateFromFile("H:\test.json")
'Debug.Print o.ToString()
'```
'@example Using chaining to build a JSON object:
'```vb
' With stdJSON.Create()                '{          
'   With .addObject("key")             '  "key": {
'     .add "a", 1                      '    "a": 1,
'     .add "b", "hello"                '    "b": "hello",
'     With .addArray("val")            '    "val": [
'       With .addArray()               '      [
'         .add 1                       '        1,
'         .add "hello"                 '        "hello",
'         .add "shit"                  '        "shit"
'       End With                       '      ]
'     End With                         '    ]
'   End With                           '  }
'                                      '}
'   Dim s as string: s = .ToString()
'End With
'```




#Const UseDictionaryLateBinding = True

Public Enum EStdJsonType
  eJSONObject
  eJSONArray
End Enum
Private Type TParser
  StoreLargeNumsAsString As Boolean
  AllowUnquotedKeys As Boolean
  AllowSingleQuoteStrings As Boolean
End Type
Private Type TThis
  iType As EStdJsonType
  children As Object
  parserProps As TParser
End Type
Private This As TThis

Private Type StringBuffer
  value As String
  index As Long
  count As Long
End Type

'Consider using ParserTarget for parsing instead for performance gains with FileSystem
' Private Enum EParserTarget
'   eParserTargetString
'   eParserTargetFile
' End Enum
' Private Type ParserTarget
'   targetType as EParserTarget
'   index as long
'   data as string
'   file as long
' End Type


'Create a JSON object
'@constructor
'@param iType - The type of json object to create. Default is JSON object, i.e. {}.
'@returns - stdJSON object of the specified type.
Public Function Create(Optional ByVal iType As EStdJsonType = eJSONObject) As stdJSON
  With stdError.getSentry("stdJSON#Create", "iType", iType)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Set Create = New stdJSON
      Call Create.protInit(iType)
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Create a JSON object from string
'@constructor
'@param sJSON                   - string to create json from
'@param AllowUnquotedKeys       - Allow unquoted keys? E.G. "{a:""hello""}"
'@param AllowSingleQuoteStrings - Allow single quoted strings? E.G. "{'a':'hello'}"
'@param StoreLargeNumsAsString  - Store Large numbers >16 chars as string instead of doubles. This may be useful if precision is important.
'@returns - `stdJSON` object from string.
'@example ```vb
'  Dim o as stdJSON: set o = stdJSON.CreateFromString("{""a"":1,""b"":""hello""}")
'  Debug.Print o.ToString()
'```
'@example ```vb
'  Dim o as stdJSON: set o = stdJSON.CreateFromString("{a: 1, b: 'hello'}", AllowUnquotedKeys:=true, AllowSingleQuoteStrings:=true)
'  Debug.Print o.ToString()
'```
Public Function CreateFromString(ByVal sJSON As String, Optional ByVal AllowUnquotedKeys As Boolean = False, Optional ByVal AllowSingleQuoteStrings As Boolean = False, Optional ByVal StoreLargeNumsAsString As Boolean = False) As stdJSON
  With stdError.getSentry("stdJSON#CreateFromString", "sJSON", sJSON, "AllowUnquotedKeys", AllowUnquotedKeys, "AllowSingleQuoteStrings", AllowSingleQuoteStrings, "StoreLargeNumsAsString", StoreLargeNumsAsString)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      'Save parser props...
      This.parserProps.AllowUnquotedKeys = AllowUnquotedKeys
      This.parserProps.StoreLargeNumsAsString = StoreLargeNumsAsString
      This.parserProps.AllowSingleQuoteStrings = AllowSingleQuoteStrings
    
      sJSON = VBA.Replace(VBA.Replace(VBA.Replace(sJSON, VBA.vbCr, ""), VBA.vbLf, ""), VBA.vbTab, "")
      Select Case Mid(sJSON, 1, 1)
        Case "{"
          Set CreateFromString = parseJSONObject(sJSON)
        Case "["
          Set CreateFromString = parseJSONArray(sJSON)
        Case Else
          Err_Raise 1, "stdJSON::CreateFromString()", "ERROR: Expecting '{' or '[' but found '" & Mid(sJSON, 1, 1) & "'"
      End Select
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Create a JSON object from file
'@constructor
'@param sJSONFile - Path to file containing json data
'@returns - stdJSON object from file data.
'@TODO: consider parsing open file num to `parseJSONObject` instead
Public Function CreateFromFile(ByVal sJSONFile As String) As stdJSON
  With stdError.getSentry("stdJSON#CreateFromFile", "sJSONFile", sJSONFile)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim sJSON As String
      Dim ff As Long: ff = FreeFile
      Open sJSONFile For Input As #ff
          sJSON = Input(LOF(ff), #ff)
      Close #ff
      Set CreateFromFile = CreateFromString(sJSON)
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Create a JSON object from VB variant
'@constructor
'@param vData - Variant to convert to JSON, e.g. Collection, Dictionary or Array
'@returns - stdJSON object from variant data.
Public Function CreateFromVariant(ByVal vData As Variant) As stdJSON
  With stdError.getSentry("stdJSON#CreateFromVariant", "vData", vData)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Set CreateFromVariant = New stdJSON
      Call CreateFromVariant.protInitFromVariant(vData)
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Create a JSON object from a set of params
'@constructor
'@param iType - The type of JSON to create, this affects how `params` is interpreted
'@param params - Either a set of `key`,`value` pairs, if creating an object, or a list of `value`s to add to an array.
'@returns - Created json object
Public Function CreateFromParams(ByVal iType As EStdJsonType, ParamArray params()) As stdJSON
  With stdError.getSentry("stdJSON#CreateFromParams", "iType", iType)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Set CreateFromParams = Create(iType)
      Dim i As Long
      Select Case iType
        Case eJSONObject
          For i = 0 To UBound(params) Step 2
            Call CreateFromParams.Add(params(i), params(i + 1))
          Next
        Case eJSONArray
          For i = 0 To UBound(params)
            Call CreateFromParams.Add(params(i))
          Next
      End Select
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Initialise stdJSON object
'@protected
'@constructor
'@param iType - The type of json object to create. Default is JSON object, i.e. {}.
Friend Sub protInit(ByRef iType As EStdJsonType)
  With stdError.getSentry("stdJSON#protInit", "iType", iType)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      This.iType = iType
      Select Case iType
        Case eJSONObject
          Set This.children = CreateDictionary()
        Case eJSONArray
          Set This.children = New Collection
      End Select
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Initialise stdJSON object from variant
'@protected
'@constructor
'@param vData - Variant to convert to JSON, e.g. `Collection`, `Dictionary` or `Array`
Friend Sub protInitFromVariant(ByRef vData As Variant)
  With stdError.getSentry("stdJSON#protInitFromVariant", "vData", vData)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim i As Long: i = 0
      Dim vItem As Variant
      Select Case VarType(vData)
        Case vbObject
          Select Case TypeName(vData)
            Case "Collection"
              This.iType = eJSONArray
              Set This.children = New Collection
    
              Dim oCollection As Collection: Set oCollection = vData
              With ChildrenCol
                For Each vItem In oCollection
                  If IsObject(vItem) Then
                    Call .Add(stdJSON.CreateFromVariant(vItem))
                  Else
                    Call .Add(vItem)
                  End If
                Next
              End With
            Case "Dictionary"
              This.iType = eJSONObject
              Set This.children = CreateDictionary()
    
              'Get keys and values...
              With vData
                Dim vKeys: vKeys = .keys()
                Dim vVals: vVals = .items()
              End With
    
              'Add to dict
              For i = 1 To vData.Count
                If IsObject(vVals(i - 1)) Then
                  Call This.children.add(vKeys(i - 1), stdJSON.CreateFromVariant(vVals(i - 1)))
                Else
                  Call This.children.add(vKeys(i - 1), vVals(i - 1))
                End If
              Next
          End Select
        Case Else
          If IsArray(vData) Then
            This.iType = eJSONArray
            Dim iLB As Long: iLB = LBound(vData)
            Dim iUB As Long: iUB = UBound(vData)
            Set This.children = New Collection
            With ChildrenCol
              For i = iLB To iUB
                If IsObject(vData(i)) Then
                  Call .Add(stdJSON.CreateFromVariant(vData(i)))
                Else
                  Call .Add(vData(i))
                End If
              Next
            End With
          End If
      End Select
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Add a key and value to an object, or a value to an array
'@param v - if This is an eJSONObject, the Key and Value to add. Else the value to add to the eJSONArray.
'@returns - if value was a stdJSON, this is returned, else `Nothing`.
'@remark Uses `protAddItem_JSONArray` when `This.iType == eJSONArray`; and `protAddItem_JSONObject` when `This.iType == eJSONObject`.
Public Function Add(ParamArray v()) As stdJSON
  With stdError.getSentry("stdJSON#Add")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Select Case This.iType
        Case eJSONObject
          Call This.children.add(v(0), v(1))
          If TypeOf v(1) Is stdJSON Then Set Add = v(1)
        Case eJSONArray
          Call This.children.add(v(0))
          If TypeOf v(0) Is stdJSON Then Set Add = v(0)
      End Select
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Create and add an array to the JSON object/array.
'@param key - If `this.iType == eJSONObject` then this is the key it should be assigned to. Else this param is unused.
'@param fromVar as Variant<Array|Collection> - If provided the object will be pre-filled with data
'@returns - The array created
'@example ```vb
' .AddArray(fromVar:=Array(1,2,3))
'```
Public Function AddArray(Optional ByVal key As String = "", Optional ByVal fromVar As Variant) As stdJSON
  With stdError.getSentry("stdJSON#AddArray", "key", key, "fromVar", fromVar)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      'Create array
      Dim oRet As stdJSON
      If IsArray(fromVar) Or IsObject(fromVar) Then
        Set oRet = stdJSON.CreateFromVariant(fromVar)
      Else
        Set oRet = stdJSON.Create(eJSONArray)
      End If
    
      'Add array to me
      Select Case This.iType
        Case eJSONObject
          Call Add(key, oRet)
        Case eJSONArray
          Call Add(oRet)
      End Select
    
      Set AddArray = oRet
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Create and add an object to the JSON object/array.
'@param key - If `this.iType == eJSONObject` then this is the key it should be assigned to. Else this param is unused.
'@returns - The object created
Public Function AddObject(Optional ByVal key As String = "") As stdJSON
  With stdError.getSentry("stdJSON#AddObject", "key", key)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      'Create JSObject
      Dim oRet As stdJSON: Set oRet = stdJSON.Create(eJSONObject)
      
      'Add JSObject
      Select Case This.iType
        Case eJSONObject
          Call Add(key, oRet)
        Case eJSONArray
          Call Add(oRet)
      End Select
      Set AddObject = oRet
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Obtain the JSON type
'@returns - The JSON type
Public Property Get JsonType() As EStdJsonType
  With stdError.getSentry("stdJSON#JsonType[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      JsonType = This.iType
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Obtain the length of this JSON object
'@returns - The length of the JSON object
Public Property Get Length() As Long
  With stdError.getSentry("stdJSON#Length[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Length = This.children.count
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Obtain the item for a key or index
'@param keyOrIndex as Variant<string|long> - Either a string key (if eJSONObject) or index as long (if eJSONArray)
'@returns - The item for the key or index
Public Property Get Item(ByVal keyOrIndex As Variant) As Variant
  With stdError.getSentry("stdJSON#Item[Get]", "keyOrIndex", keyOrIndex)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
    Attribute Item.VB_UserMemId = 0
      Select Case This.iType
        Case eJSONObject
          Call CopyVariant(Item, This.children.item(keyOrIndex))
        Case eJSONArray
          Call CopyVariant(Item, This.children.item(keyOrIndex))
      End Select
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property
Public Property Let Item(ByVal keyOrIndex As Variant, v As Variant)
  With stdError.getSentry("stdJSON#Item[Let]", "keyOrIndex", keyOrIndex, "v", v)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Select Case This.iType
        Case eJSONObject
          This.children.item(keyOrIndex) = v
        Case eJSONArray
          This.children.item(keyOrIndex) = v
      End Select
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property
Public Property Set Item(ByVal keyOrIndex As Variant, v As Variant)
  With stdError.getSentry("stdJSON#Item[Set]", "keyOrIndex", keyOrIndex, "v", v)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Select Case This.iType
        Case eJSONObject
          Set This.children.item(keyOrIndex) = v
        Case eJSONArray
          Set This.children.item(keyOrIndex) = v
      End Select
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Obtain the item for a key or index
'@param keyOrIndex as Variant<string|long> - Either a string key (if eJSONObject) or index as long (if eJSONArray)
'@returns - Whether the key or index exists
Public Property Get Exists(ByVal keyOrIndex As Variant) As Boolean
  With stdError.getSentry("stdJSON#Exists[Get]", "keyOrIndex", keyOrIndex)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Select Case This.iType
        Case eJSONObject
          Exists = This.children.exists(keyOrIndex)
        Case eJSONArray
          Exists = keyOrIndex <= This.children.count And keyOrIndex > 0
      End Select
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Obtain the children of this JSON object
'@returns - The children of this JSON object of the form [{key: string, value: variant, isJSON: boolean},...]
'@remark - Because of the nature of the JSON object, this will be a collection of dictionaries with the keys "key", "value" and "isJSON".
Public Function ChildrenInfo() As Collection
  With stdError.getSentry("stdJSON#ChildrenInfo")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim ItemParent As stdJSON: Set ItemParent = Me
      Set ChildrenInfo = New Collection
      Select Case This.iType
        Case eJSONObject
          Dim vKey As Variant
          For Each vKey In This.children.Keys()
            Call ChildrenInfo.Add(CreateDictionary("key", vKey, "value", This.children(vKey), "isJSON", TypeOf This.children(vKey) Is stdJSON, "parent", ItemParent))
          Next
        Case eJSONArray
          Dim vChild As Variant, i As Long: i = 0
          For Each vChild In This.children
            i = i + 1
            Call ChildrenInfo.Add(CreateDictionary("key", i, "value", vChild, "isJSON", TypeOf vChild Is stdJSON, "parent", ItemParent))
          Next
      End Select
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function


' 'Obtain the item for a key or index
' '@param {string} Query - JsonPath query to get/set variable
' '@remark see https://github.com/json-path/JsonPath for details
' '@TODO: implementation - might be better in it's own class
' Private Property Get ItemByPath(ByVal query as string) as Variant
  
' End Property
' Private Property Let ItemByPath(ByVal keyOrIndex as Variant, v as Variant)
  
' End Property
' Private Property Set ItemByPath(ByVal keyOrIndex as Variant, v as Variant)
  
' End Property

'Convert the JSON to string
'@param pretty - If true, pretty print the JSON
'@param prettyIndent - If pretty, the indent to use
'@param depth - If pretty, the depth to start at
'@returns - JSON encoded data.
Public Function ToString(Optional ByVal pretty As Boolean = False, Optional ByVal prettyIndent As Long = 2, Optional depth As Long = 1) As String
  With stdError.getSentry("stdJSON#ToString", "pretty", pretty, "prettyIndent", prettyIndent, "depth", depth)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim sRootIndent As String, sChildIndent As String, sNewLine As String
      If pretty Then
        sChildIndent = String$(depth * prettyIndent, " ")
        sRootIndent = String$((depth - 1) * prettyIndent, " ")
        sNewLine = vbCrLf
      End If
    
      Dim s As StringBuffer: s = StringBuffer_Create(64)
      Select Case This.iType
        Case eJSONObject
          Call StringBuffer_Append(s, sRootIndent & "{" & sNewLine)
        Case eJSONArray
          Call StringBuffer_Append(s, sRootIndent & "[" & sNewLine)
      End Select
    
      Dim i As Long, ub As Long, sValue As String
      Select Case This.iType
        Case eJSONObject
          Dim vVals: vVals = This.children.items()
          Dim vKeys: vKeys = This.children.keys()
          ub = UBound(vKeys)
          For i = 0 To ub
            If TypeOf vVals(i) Is stdJSON Then
              sValue = LTrim(vVals(i).ToString(pretty, prettyIndent, depth + 1))
            Else
              sValue = simpleStringify(vVals(i))
            End If
            'Concatenation of small strings is faster than calling Append, so optimise on this
            Call StringBuffer_Append(s, sChildIndent & """" & vKeys(i) & """: ")
            Call StringBuffer_Append(s, sValue)
            Call StringBuffer_Append(s, IIf(i < ub, ",", "") & sNewLine)
          Next
        Case eJSONArray
          i = 0
          ub = This.children.count
          Dim vItem
          For Each vItem In This.children
            i = i + 1
            If TypeOf vItem Is stdJSON Then
              sValue = LTrim(vItem.ToString(pretty, prettyIndent, depth + 1))
            Else
              sValue = simpleStringify(vItem)
            End If
            Call StringBuffer_Append(s, sChildIndent)
            Call StringBuffer_Append(s, sValue)
            Call StringBuffer_Append(s, IIf(i < ub, ",", "") & sNewLine)
          Next
      End Select
      
      Select Case This.iType
        Case eJSONObject
          Call StringBuffer_Append(s, sRootIndent & "}")
        Case eJSONArray
          Call StringBuffer_Append(s, sRootIndent & "]")
      End Select
      
      ToString = StringBuffer_ToString(s)
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Dump the JSON data to a file
'@param sPath - Path to file to dump JSON data to
'@param prettyIndent - If pretty, the indent to use
Public Sub ToFile(ByVal sPath As String, Optional ByVal prettyIndent As Long = 0)
  With stdError.getSentry("stdJSON#ToFile", "sPath", sPath, "prettyIndent", prettyIndent)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim ff As Long: ff = FreeFile
      Open sPath For Output As #ff
          Print #ff, ToString(pretty:=prettyIndent > 0, prettyIndent:=prettyIndent)
      Close #ff
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Convert to VBA Object
'@returns Object<Collection|Dictionary> - `Dictionary` if `eJSONObject`, `Collection` if `eJSONArray`
Public Function ToVBObject() As Object
  With stdError.getSentry("stdJSON#ToVBObject")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim ret As Object
      Select Case This.iType
        Case eJSONObject
          Set ret = CreateDictionary()
        Case eJSONArray
          Set ret = New Collection
      End Select
        
      Dim i As Long, ub As Long
      Select Case This.iType
        Case eJSONObject
          Dim vVals: vVals = This.children.items()
          Dim vKeys: vKeys = This.children.keys()
          ub = UBound(vKeys)
          For i = 0 To ub
            If TypeOf vVals(i) Is stdJSON Then
              ret.add vKeys(i), vVals(i).ToVBObject
            Else
              ret.add vKeys(i), vVals(i)
            End If
          Next
        Case eJSONArray
          Dim vItem
          For Each vItem In This.children
            If TypeOf vItem Is stdJSON Then
              ret.add vItem.ToVBObject
            Else
              ret.add vItem
            End If
          Next
      End Select
    
      Set ToVBObject = ret
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Returns the current json object
'@example ```vb
'With stdJSON.Create()
'  .add "X", "Y"
'  set j = .ToSelf()
'End With
'```
Public Function ToSelf() As stdJSON
  With stdError.getSentry("stdJSON#ToSelf")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Set ToSelf = Me
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Helper: For type safety, obtain this this.children as collection
'@returns {Collection} The children collection
Private Property Get ChildrenCol() As Collection
  With stdError.getSentry("stdJSON#ChildrenCol[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Set ChildrenCol = This.children
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Allow `for each x in json` syntax
'@returns - The enumerator to iterate over
Public Function NewEnum() As IUnknown
  With stdError.getSentry("stdJSON#NewEnum")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
    Attribute NewEnum.VB_UserMemId = -4
    Attribute NewEnum.VB_MemberFlags = "40"
      Select Case This.iType
        Case eJSONObject
          Set NewEnum = This.children.[_NewEnum]
        Case eJSONArray
          Set NewEnum = This.children.[_NewEnum]
      End Select
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'*******************
'* PRIVATE HELPERS *
'*******************

'Helper: Copy variant from one location to another
'@param dest - Destination variant
'@param src - Source variant
Private Sub CopyVariant(ByRef dest As Variant, ByVal src As Variant)
  With stdError.getSentry("stdJSON#CopyVariant", "dest", dest, "src", src)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If IsObject(src) Then
        Set dest = src
      Else
        Let dest = src
      End If
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Helper: stringify a simple construct
'@param value - The value to stringify
'@returns {string} The stringified value
Private Function simpleStringify(ByVal value As Variant) As String
  With stdError.getSentry("stdJSON#simpleStringify", "value", value)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Select Case VarType(value)
        Case vbNull, vbEmpty
          simpleStringify = "null"
        Case vbDate
          simpleStringify = """" & Format(value, "dd/mm/yyyy") & """"
        Case vbString
          value = Replace(value, "\", "\\")
          value = Replace(value, """", "\""")
          value = Replace(value, vbTab, "\t")
          value = Replace(value, vbCr, "\r")
          value = Replace(value, vbLf, "\n")
          simpleStringify = """" & value & """"
        Case vbBoolean
          simpleStringify = IIf(value, "true", "false")
        Case vbInteger
          simpleStringify = value
        Case vbDouble, vbSingle
          'Handle different locales where "." is not the decimal separator
          Dim intPart As Long: intPart = Fix(value)
          If value - intPart = 0 Then
            simpleStringify = value
          Else
            simpleStringify = intPart & "." & Mid(CStr(value - intPart), IIf(value < 0, 4, 3))
          End If
        Case Else
          If IsArray(value) Then
            simpleStringify = stdJSON.CreateFromVariant(value).ToString()
          Else
            simpleStringify = value
          End If
      End Select
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'***********
'* PARSING *
'***********

'Parser: Parse a JSON object
'@param sJSON - The JSON string to parse
'@param index - The index to start parsing from
'@returns - The parsed JSON object
Private Function parseJSONObject(ByRef sJSON As String, Optional ByRef index As Long = 1) As stdJSON
  With stdError.getSentry("stdJSON#parseJSONObject", "sJSON", sJSON, "index", index)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim json As stdJSON: Set json = stdJSON.Create(eJSONObject)
      Call parserConsume("{", sJSON, index)
      Call parserIgnoreSpaces(sJSON, index)
      While Not parserPeek("}", sJSON, index)
        Dim key As String: key = parseJSONObjectKey(sJSON, index)
        Call parserIgnoreSpaces(sJSON, index)
        Call parserConsume(":", sJSON, index)
        Call parserIgnoreSpaces(sJSON, index)
        Dim value: Call parseValue(value, sJSON, index)
        json.Add key, value
        Call parserIgnoreSpaces(sJSON, index)
        If parserOptConsume(",", sJSON, index) Then Call parserIgnoreSpaces(sJSON, index)
      Wend
      Call parserConsume("}", sJSON, index)
      Set parseJSONObject = json
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Parser: Parse a JSON array
'@param sJSON - The JSON string to parse
'@param index - The index to start parsing from
'@returns - The parsed JSON array
Private Function parseJSONArray(ByRef sJSON As String, Optional ByRef index As Long = 1) As stdJSON
  With stdError.getSentry("stdJSON#parseJSONArray", "sJSON", sJSON, "index", index)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim json As stdJSON: Set json = stdJSON.Create(eJSONArray)
      Call parserConsume("[", sJSON, index)
      Call parserIgnoreSpaces(sJSON, index)
      While Not parserPeek("]", sJSON, index)
        Dim value: Call parseValue(value, sJSON, index)
        json.Add value
        Call parserIgnoreSpaces(sJSON, index)
        If parserOptConsume(",", sJSON, index) Then Call parserIgnoreSpaces(sJSON, index)
      Wend
      Call parserConsume("]", sJSON, index)
      Set parseJSONArray = json
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Parser: Parse a JSON object key
'@param sJSON - The JSON string to parse
'@param index - The index to start parsing from
'@returns - The parsed JSON object key
Private Function parseJSONObjectKey(ByRef sJSON As String, ByRef index As Long) As String
  With stdError.getSentry("stdJSON#parseJSONObjectKey", "sJSON", sJSON, "index", index)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim char As String: char = parserPeekAny(sJSON, index)
      If char = """" Then
        parseJSONObjectKey = parseString(sJSON, index)
      ElseIf This.parserProps.AllowUnquotedKeys Then
        Dim sKey As String: sKey = ""
        Do While True
          char = parserPeekAny(sJSON, index)
          If (char <> " ") And (char <> ":") Then
            Call parserConsumeAny(sJSON, index)
            sKey = sKey & char
          Else
            parseJSONObjectKey = sKey
            Exit Do
          End If
        Loop
      End If
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Parser: Parse a JSON string
'@param sJSON - The JSON string to parse
'@param index - The index to start parsing from
'@returns - The parsed JSON string
Private Function parseString(ByRef sJSON As String, Optional ByRef index As Long = 1) As String
  With stdError.getSentry("stdJSON#parseString", "sJSON", sJSON, "index", index)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim buffer As StringBuffer: buffer = StringBuffer_Create()
      Dim char As String
      
      Dim sQuoteType As String
      If parserOptConsume("""", sJSON, index) Then
        sQuoteType = """"
      ElseIf This.parserProps.AllowSingleQuoteStrings Then
        Call parserConsume("'", sJSON, index)
        sQuoteType = "'"
      End If
      
      While Not parserPeek(sQuoteType, sJSON, index)
        char = parserConsumeAny(sJSON, index)
        Select Case char
          Case "\"
            char = parserConsumeAny(sJSON, index)
            Select Case char
              Case """", "\", "/", sQuoteType: Call StringBuffer_Append(buffer, char)
              Case "b": Call StringBuffer_Append(buffer, vbBack)
              Case "f": Call StringBuffer_Append(buffer, vbFormFeed)
              Case "n": Call StringBuffer_Append(buffer, vbLf)
              Case "r": Call StringBuffer_Append(buffer, vbCr)
              Case "t": Call StringBuffer_Append(buffer, vbTab)
              Case "u"
                Dim WCode As String: WCode = parserConsumeAny(sJSON, index, 4)
                Call StringBuffer_Append(buffer, ChrW("&H" & WCode))
            End Select
          Case Else
            Call StringBuffer_Append(buffer, char)
        End Select
      Wend
      Call parserConsume(sQuoteType, sJSON, index)
      parseString = StringBuffer_ToString(buffer)
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Parser: Parse a JSON value
'@param value - The value to parse
'@param sJSON - The JSON string to parse
'@param index - The index to start parsing from
Private Sub parseValue(ByRef value As Variant, ByRef sJSON As String, ByRef index As Long)
  With stdError.getSentry("stdJSON#parseValue", "value", value, "sJSON", sJSON, "index", index)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim char As String: char = parserPeekAny(sJSON, index)
      Select Case char
        Case "{"
          Set value = parseJSONObject(sJSON, index)
        Case "["
          Set value = parseJSONArray(sJSON, index)
        Case """", "'"
          value = parseString(sJSON, index)
        Case "t"
          Call parserConsume("true", sJSON, index, 4)
          value = True
        Case "f"
          Call parserConsume("false", sJSON, index, 5)
          value = False
        Case "n"
          Call parserConsume("null", sJSON, index, 4)
          value = Null
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "-"
          value = parseNumber(sJSON, index)
      End Select
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Parser: Parse a JSON number
'@param sJSON - The JSON string to parse
'@param index - The index to start parsing from
'@returns - The parsed JSON number
Private Function parseNumber(ByRef sJSON As String, ByRef index As Long) As Double
  With stdError.getSentry("stdJSON#parseNumber", "sJSON", sJSON, "index", index)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim char As String: char = parserPeekAny(sJSON, index)
      Dim sNum As String: sNum = ""
      While InStr("+-0123456789.eE", char) 'Test char
        index = index + 1                  'Move to next char
        sNum = sNum & char                 'Generate number
        char = parserPeekAny(sJSON, index) 'Prepare next char
      Wend
      
      'Handle large numbers losing precision
      If This.parserProps.StoreLargeNumsAsString And Len(sNum) >= 16 Then
        parseNumber = sNum
      Else
        parseNumber = Val(sNum)
      End If
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function


'Parser helper: Ignore spaces in JSON string
'@param sJSON - The JSON string to parse
'@param index - The index to start parsing from
Private Sub parserIgnoreSpaces(ByRef sJSON As String, ByRef index As Long)
  With stdError.getSentry("stdJSON#parserIgnoreSpaces", "sJSON", sJSON, "index", index)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      While Mid(sJSON, index, 1) = " "
        index = index + 1
      Wend
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Parser helper: Consumes some characters from the JSON string only if they match the string
'@param sChar - The string to consume
'@param sJSON - The JSON string to parse
'@param index - The index to start parsing from
'@param iLength - The length of the string to consume
Private Sub parserConsume(ByVal sChar As String, ByRef sJSON As String, ByRef index As Long, Optional ByVal iLength As Long = 1)
  With stdError.getSentry("stdJSON#parserConsume", "sChar", sChar, "sJSON", sJSON, "index", index, "iLength", iLength)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
        If Mid(sJSON, index, iLength) = sChar Then
          index = index + iLength
        Else
          Err_Raise 1, "", "JSONParseError: Unexpected char, found: " & Mid(sJSON, index, iLength) & " but expected: " & sChar
        End If
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Parser helper: Consumes some characters from the JSON string
'@param sJSON - The JSON string to parse
'@param index - The index to start parsing from
'@param iLength - The length of the string to consume
'@returns - The consumed string
Private Function parserConsumeAny(ByRef sJSON As String, ByRef index As Long, Optional ByVal iLength As Long = 1) As String
  With stdError.getSentry("stdJSON#parserConsumeAny", "sJSON", sJSON, "index", index, "iLength", iLength)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      parserConsumeAny = Mid(sJSON, index, iLength)
      index = index + iLength
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Parser helper: Checks whether the token at iTokenIndex is of the given type
'@param sChar - The string to check
'@param sJSON - The JSON string to parse
'@param index - The index to start parsing from
'@param iLength - The length of the string to check
'@returns - Whether the token at iTokenIndex is of the given type
Private Function parserPeek(ByVal sChar As String, ByRef sJSON As String, ByRef index As Long, Optional ByVal iLength As Long = 1) As Boolean
  With stdError.getSentry("stdJSON#parserPeek", "sChar", sChar, "sJSON", sJSON, "index", index, "iLength", iLength)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
        parserPeek = Mid(sJSON, index, iLength) = sChar
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Parser helper: Obtains a character without consuming the token
'@param sJSON - The JSON string to parse
'@param index - The index to start parsing from
'@returns - The character at iTokenIndex
Private Function parserPeekAny(ByRef sJSON As String, ByRef index As Long) As String
  With stdError.getSentry("stdJSON#parserPeekAny", "sJSON", sJSON, "index", index)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
        parserPeekAny = Mid(sJSON, index, 1)
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Parser helper: Combines peek and consume, consuming a token only if matched, without throwing an error if not
'@param sChar - The string to consume
'@param sJSON - The JSON string to parse
'@param index - The index to start parsing from
'@returns - Whether the token at iTokenIndex is of the given type
Private Function parserOptConsume(ByVal sChar As String, ByRef sJSON As String, ByRef index As Long) As Boolean
  With stdError.getSentry("stdJSON#parserOptConsume", "sChar", sChar, "sJSON", sJSON, "index", index)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
        Dim matched As Boolean: matched = parserPeek(sChar, sJSON, index)
        If matched Then
            Call parserConsume(sChar, sJSON, index)
        End If
        parserOptConsume = matched
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function


'Create a stringbuffer
'@param Initial - The initial size of the buffer
'@returns - The stringbuffer
Private Function StringBuffer_Create(Optional ByVal Initial As Long = 16) As StringBuffer
  With stdError.getSentry("stdJSON#StringBuffer_Create", "Initial", Initial)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      StringBuffer_Create.value = String$(Initial, " ")
      StringBuffer_Create.count = Initial
      StringBuffer_Create.index = 1
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Convert a stringbuffer to a string
'@param buffer - The stringbuffer to convert
'@returns - The stringbuffer as a string
Private Function StringBuffer_ToString(ByRef buffer As StringBuffer) As String
  With stdError.getSentry("stdJSON#StringBuffer_ToString")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      StringBuffer_ToString = Left(buffer.value, buffer.index - 1)
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Append a string to a stringbuffer
'@param buffer - The stringbuffer to append to
'@param sToAppend - The string to append
'@remark If the stringbuffer is not large enough, it will be upsized
Private Sub StringBuffer_Append(ByRef buffer As StringBuffer, ByVal sToAppend As String)
  With stdError.getSentry("stdJSON#StringBuffer_Append", "sToAppend", sToAppend)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim lenToAppend As Long: lenToAppend = Len(sToAppend)
      If buffer.index + lenToAppend > buffer.count Then
        Call StringBuffer_Upsize(buffer, lenToAppend)
      End If
      Mid(buffer.value, buffer.index, lenToAppend) = sToAppend
      buffer.index = buffer.index + lenToAppend
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Upsize a stringbuffer
'@param buffer - The stringbuffer to upsize
'@param minIncrease - The minimum amount to increase the buffer by
Private Sub StringBuffer_Upsize(ByRef buffer As StringBuffer, Optional ByVal minIncrease As Long)
  With stdError.getSentry("stdJSON#StringBuffer_Upsize", "minIncrease", minIncrease)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      minIncrease = IIf(buffer.count > minIncrease, buffer.count, minIncrease)
      buffer.count = buffer.count + minIncrease
      buffer.value = buffer.value & String$(minIncrease, " ")
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Remove characters from the end of a stringbuffer
'@param buffer - The stringbuffer to pop from
'@param iCount - The number of characters to remove
Private Sub StringBuffer_Pop(ByRef buffer As StringBuffer, Optional ByVal iCount As Long = 1)
  With stdError.getSentry("stdJSON#StringBuffer_Pop", "iCount", iCount)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      buffer.index = buffer.index - iCount
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Create a dictionary
'@returns - The dictionary
Private Function CreateDictionary(ParamArray children()) As Object
  With stdError.getSentry("stdJSON#CreateDictionary")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      #If UseDictionaryLateBinding Then
        Set CreateDictionary = CreateObject("Scripting.Dictionary")
      #Else
        Set CreateDictionary = New Scripting.Dictionary
      #End If
      CreateDictionary.CompareMode = vbTextCompare
    
      Dim i As Long
      For i = LBound(children) To UBound(children) Step 2
        Call CreateDictionary.add(children(i), children(i + 1))
      Next
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function


Private Sub Err_Raise(ByVal number as Long, Optional ByVal source as string = "", Optional ByVal description as string = "")
  Call stdError.Raise(description)
End Sub
