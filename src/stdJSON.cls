VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdJSON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Spec
' Create JSON:
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
'   'Get json object ref               '}
'  Dim o as stdJSON: set o = .obj
'   'Get as dictionary of collections
'   Dim o as object: set o = .toVBON()
'   'Get json string
'   Dim s as string: s = .toString()
' End With

Public Enum EStdJsonType
  eJSONObject
  eJSONArray
End Enum
Private Type TParser
  StoreLargeNumsAsString as boolean
  AllowUnquotedKeys as boolean
  AllowSingleQuoteStrings as boolean
End Type
Private Type TThis
  iType as EStdJsonType
  children as Object
  parserProps as TParser
End Type
Private This as TThis

Private Type StringBuffer
  value as string
  index as Long
  count as long
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
Public Function Create(Optional ByVal iType as EStdJsonType = eJSONObject) as stdJSON
  set Create = new stdJSON
  Call Create.protInit(iType)
End Function

'Create a JSON object from string
'@constructor
'@param sJSON                   - string to create json from
'@param AllowUnquotedKeys       - Allow unquoted keys? E.G. "{a:""hello""}"
'@param AllowSingleQuoteStrings - Allow single quoted strings? E.G. "{'a':'hello'}"
'@param StoreLargeNumsAsString  - Store Large numbers >16 chars as string instead of doubles. This may be useful if precision is important.
'@returns - `stdJSON` object from string.
Public Function CreateFromString(ByVal sJSON as string, optional ByVal AllowUnquotedKeys as boolean = false, optional ByVal AllowSingleQuoteStrings as boolean = false, optional ByVal StoreLargeNumsAsString as boolean = false) as stdJSON
  'Save parser props...
  this.parserProps.AllowUnquotedKeys = AllowUnquotedKeys
  this.parserProps.StoreLargeNumsAsString = StoreLargeNumsAsString
  this.parserProps.AllowSingleQuoteStrings = AllowSingleQuoteStrings

  sJSON = VBA.Replace(VBA.Replace(VBA.Replace(sJSON, VBA.vbCr, ""), VBA.vbLf, ""), VBA.vbTab, "")
  select case mid(sJSON, 1, 1)
    case "{"
      set CreateFromString = parseJSONObject(sJSON)
    case "["
      set CreateFromString = parseJSONArray(sJSON)
    case else
      Err.Raise 1, "stdJSON::CreateFromString()", "ERROR: Expecting '{' or '[' but found '" & mid(sJson,1,1) & "'"
  end select
End Function

'Create a JSON object from file
'@constructor
'@param sJSONFile - Path to file containing json data
'@returns - stdJSON object from file data.
'@TODO: consider parsing open file num to `parseJSONObject` instead
Public Function CreateFromFile(ByVal sJSONFile as string) as stdJSON
  Dim sJSON as string
  Dim ff As Long: ff = FreeFile
  Open sJSONFile For Input As #ff
      sJSON = Input(LOF(ff), #ff)
  Close #ff
  set CreateFromFile = CreateFromString(sJSON)
End Function

'Create a JSON object from VB variant
'@constructor
'@param vData - Variant to convert to JSON, e.g. Collection, Dictionary or Array
'@returns - stdJSON object from variant data.
Public Function CreateFromVariant(ByVal vData as Variant) as stdJSON
  set CreateFromVariant = new stdJSON
  Call CreateFromVariant.protInitFromVariant(vData)
End Function

'Initialise stdJSON object
'@protected
'@constructor
'@param iType - The type of json object to create. Default is JSON object, i.e. {}.
Friend Sub protInit(ByRef iType as EStdJsonType)
  This.iType = iType
  select case iType
    case eJSONObject
      set this.children = CreateObject("Scripting.Dictionary")
    case eJSONArray
      set this.children = new Collection
  end select
End Sub

'Initialise stdJSON object from variant
'@protected
'@constructor
'@param vData - Variant to convert to JSON, e.g. `Collection`, `Dictionary` or `Array`
Friend Sub protInitFromVariant(ByRef vData as Variant)
  Dim i as Long: i = 0
  Dim vItem as variant
  select case vartype(vData)
    case vbObject
      select case typename(vData)
        case "Collection"
          this.iType = eJSONArray
          set this.children = new Collection

          Dim oCollection as Collection: set oCollection = vData
          With ChildrenCol
            For each vItem in oCollection
              Call .add(vItem)
            next
          End With
        case "Dictionary"
          this.iType = eJSONObject
          set this.children = CreateObject("Scripting.Dictionary")

          'Get keys and values...
          With vData
            Dim vKeys: vKeys = .keys()
            Dim vVals: vVals = .items()
          End with

          'Add to dict
          For i = 1 to oDictionary.Count
            Call this.children.add(vKeys(i-1), vVals(i-1))
          next
      end select
    case else
      if isArray(vData) then
        this.iType = eJSONArray
        Dim iLB as Long: iLB = LBound(vData)
        Dim iUB as Long: iUB = UBound(vData)
        set this.children = new Collection
        With ChildrenCol
          For i = iLB to iUB
            Call .add(vData(i))
          next
        End With
      end if
  end select
End Sub

'Add a key and value to an object, or a value to an array
'@param v - if This is an eJSONObject, the Key and Value to add. Else the value to add to the eJSONArray.
'@returns - if value was a stdJSON, this is returned, else `Nothing`.
'@remark Uses `protAddItem_JSONArray` when `This.iType == eJSONArray`; and `protAddItem_JSONObject` when `This.iType == eJSONObject`.
Public Function Add(ParamArray v()) as stdJSON
  select case This.iType
    case eJSONObject
      Call this.children.add(v(0), v(1))
      if typeof v(1) is stdJSON then set Add = v(1)
    case eJSONArray
      Call this.children.add(v(0))
      if typeof v(0) is stdJSON then set Add = v(0)
  end select
End Function

'Create and add an array to the JSON object/array.
'@param key - If `this.iType == eJSONObject` then this is the key it should be assigned to. Else this param is unused.
'@param fromVar as Variant<Array|Collection> - If provided the object will be pre-filled with data
'@returns - The array created
'@example ` .AddArray(fromVar:=Array(1,2,3))`
Public Function AddArray(Optional ByVal key as string = "", optional ByVal fromVar as variant) as stdJSON
  'Create array
  Dim oRet as stdJSON
  if isArray(fromVar) or isObject(fromVar) then
    set oRet = stdJSON.CreateFromVariant(fromVar)
  else
    set oRet = stdJSON.Create(eJSONArray)
  end if

  'Add array to me
  select case This.iType
    case eJSONObject
      Call Add(key, oRet)
    case eJSONArray
      Call Add(oRet)
  end select

  set AddArray = oRet
End Function

'Create and add an object to the JSON object/array.
'@param key - If `this.iType == eJSONObject` then this is the key it should be assigned to. Else this param is unused.
'@returns - The object created
Public Function AddObject(Optional ByVal key as string = "") as stdJSON
  'Create JSObject
  Dim oRet as stdJSON: set oRet = stdJSON.Create(eJSONObject)
  
  'Add JSObject
  select case This.iType
    case eJSONObject
      Call Add(key, oRet)
    case eJSONArray
      Call Add(oRet)
  end select
  set AddObject = oRet
End Function

'Obtain the length of this JSON object
'@returns - The length of the JSON object
Public Property Get Length() as Long
  Length = this.children.count
End Property

'Obtain the item for a key or index
'@param keyOrIndex as Variant<string|long> - Either a string key (if eJSONObject) or index as long (if eJSONArray)
Public Property Get Item(ByVal keyOrIndex as Variant) as Variant
Attribute Item.VB_UserMemId = 0
  select case this.iType
    case eJSONObject
      Call CopyVariant(Item, this.children.item(keyOrIndex))
    case eJSONArray
      Call CopyVariant(Item, this.children.item(keyOrIndex))
  end select
End Property
Public Property Let Item(ByVal keyOrIndex as Variant, v as Variant)
  select case this.iType
    case eJSONObject
      this.children.item(keyOrIndex) = v
    case eJSONArray
      this.children.item(keyOrIndex) = v
  end select
End Property
Public Property Set Item(ByVal keyOrIndex as Variant, v as Variant)
  select case this.iType
    case eJSONObject
      set this.children.item(keyOrIndex) = v
    case eJSONArray
      set this.children.item(keyOrIndex) = v
  end select
End Property

'Obtain the item for a key or index
'@param keyOrIndex as Variant<string|long> - Either a string key (if eJSONObject) or index as long (if eJSONArray)
'@returns - Whether the key or index exists
Public Property Get Exists(ByVal keyOrIndex as Variant) as Boolean
  select case this.iType
    case eJSONObject
      Exists = this.children.exists(keyOrIndex)
    case eJSONArray
      Exists = keyOrIndex <= this.children.count and keyOrIndex > 0
  end select
End Property

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
Public Function ToString(optional byval pretty as boolean = false, optional byval prettyIndent as long = 2, optional depth as long = 1) as String
  Dim sRootIndent as string, sChildIndent as string, sNewLine as string
  if pretty then 
    sChildIndent = string$(depth*prettyIndent, " ")
    sRootIndent = string$((depth-1)*prettyIndent," ")
    sNewLine = vbCrLf
  end if

  Dim s as StringBuffer: s = StringBuffer_Create(64)
  select case This.iType
    case eJSONObject
      Call StringBuffer_Append(s, sRootIndent & "{" & sNewLine) 
    case eJSONArray
      Call StringBuffer_Append(s, sRootIndent & "[" & sNewLine)
  end select

  Dim i as Long, ub as Long
  select case this.iType
    case eJSONObject
      Dim vVals: vVals = this.children.items()
      Dim vKeys: vKeys = this.children.keys()
      ub = ubound(vKeys)
      For i = 0 to ub
        if typeof vVals(i) is stdJSON then
          sValue = vVals(i).ToString(pretty, prettyIndent*depth, depth+1)
        else
          sValue = simpleStringify(vVals(i))
        end if
        'Concatenation of small strings is faster than calling Append, so optimise on this
        Call StringBuffer_Append(s, sChildIndent & """" & vKeys(i) & """:")
        Call StringBuffer_Append(s, sValue)
        Call StringBuffer_Append(s, iif(i<ub, ",", "") & sNewLine)
      next
    case eJSONArray
      i = 0
      ub = this.children.count
      Dim vItem
      For each vItem in this.children
        i=i+1
        if typeof vItem is stdJSON then
          sValue = vItem.ToString(pretty, prettyIndent*depth, depth+1)
        else
          sValue = simpleStringify(vItem)
        end if
        Call StringBuffer_Append(s, sChildIndent)
        Call StringBuffer_Append(s, sValue)
        Call StringBuffer_Append(s, iif(i<ub, ",", "") & sNewLine)
      next
  end select
  
  select case This.iType
    case eJSONObject
      Call StringBuffer_Append(s, sRootIndent & "}" & sNewLine)
    case eJSONArray
      Call StringBuffer_Append(s, sRootIndent & "]" & sNewLine)
  end select

  ToString = StringBuffer_ToString(s)
End Function

'Dump the JSON data to a file
'@param sPath - Path to file to dump JSON data to
'@param prettyIndent - If pretty, the indent to use
Public Sub ToFile(ByVal sPath as string, Optional ByVal prettyIndent as Long = 0)
  Dim ff As Long: ff = FreeFile
  Open sPath For Output As #ff
      Print #ff, data
  Close #ff
End Sub

'Convert to VBA Object
'@returns Object<Collection|Dictionary> - `Dictionary` if `eJSONObject`, `Collection` if `eJSONArray`
Public Function ToVBObject() as Object
  Dim ret as Object
  select case This.iType
    case eJSONObject
      set ret = CreateObject("Scripting.Dictionary")
    case eJSONArray
      set ret = new Collection
  end select
    
  Dim i as Long, ub as Long
  select case this.iType
    case eJSONObject
      Dim vVals: vVals = this.children.items()
      Dim vKeys: vKeys = this.children.keys()
      ub = ubound(vKeys)
      For i = 0 to ub
        if typeof vVals(i) is stdJSON then
          ret.add vKeys(i), vVals(i).ToVBObject
        else
          ret.add vKeys(i), vVals(i)
        end if
      next
    case eJSONArray
      Dim vItem
      For each vItem in this.children
        if typeof vItem is stdJSON then
          ret.add vItem.ToVBObject
        else
          ret.add vItem
        end if
      next
  end select

  set ToVBObject = ret
End Function

'Helper: For type safety, obtain this this.children as collection
'@returns {Collection} The children collection
Private Property Get ChildrenCol() as Collection
  set ChildrenCol = this.Children
End Property

'Allow `for each x in json` syntax
'@returns {IEnumVariant} The enumerator to iterate over
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  select case This.iType
    case eJSONObject
      Set NewEnum = This.children.[_NewEnum]
    case eJSONArray
      set NewEnum = This.children.[_NewEnum]
  end select
End Function

'*******************
'* PRIVATE HELPERS *
'*******************

'Helper: Copy variant from one location to another
'@param dest - Destination variant
'@param src - Source variant
Private Sub CopyVariant(ByRef dest as Variant, ByVal src as Variant)
  if isObject(src) then
    set dest = src
  else
    let dest = src
  end if
End Sub

'Helper: stringify a simple construct
'@param value - The value to stringify
'@returns {string} The stringified value
Private Function simpleStringify(ByVal value as variant) as string
  select case vartype(value)
    case vbNull
      simpleStringify = "null"
    case vbDate
      simpleStringify = """" & Format(value, "dd/mm/yyyy") & """"
    case vbString
      simpleStringify = """" & value & """"
    Case vbBoolean
      simpleStringify = iif(value, "true", "false")
    case vbInteger, vbDouble, vbFloat
      simpleStringify = value
    case else
      if isArray(value) then
        simpleStringify = stdJSON.CreateFromVariant(value).toString()
      else
        simpleStringify = value
      end if
  end select
End Function

'***********
'* PARSING *
'***********

'Parser: Parse a JSON object
'@param sJSON - The JSON string to parse
'@param index - The index to start parsing from
'@returns - The parsed JSON object
Private Function parseJSONObject(ByRef sJSON as string, Optional ByRef index as Long = 1) as stdJSON
  set json = stdJSON.Create(eJSONObject)
  Call parserConsume("{", sJSON, index)
  Call parserIgnoreSpaces(sJSON, index)
  While not parserPeek("}", sJSON, index)
    Dim key as string: key = parseJSONObjectKey(sJSON, index)
    Call parserIgnoreSpaces(sJSON, index)
    Call parserConsume(":", sJSON, index)
    Call parserIgnoreSpaces(sJSON, index)
    Dim value: Call parseValue(value, sJSON, index)
    json.add key, value
    Call parserIgnoreSpaces(sJSON, index)
    if parserOptConsume(",", sJSON, index) then Call parserIgnoreSpaces(sJSON, index)
  Wend
  Call parserConsume("}", sJSON, index)
  set parseJSONObject = json
End Function

'Parser: Parse a JSON array
'@param sJSON - The JSON string to parse
'@param index - The index to start parsing from
'@returns - The parsed JSON array
Private Function parseJSONArray(ByRef sJSON as string, Optional ByRef index as Long = 1) as stdJSON
  set json = stdJSON.Create(eJSONArray)
  Call parserConsume("[", sJSON, index)
  Call parserIgnoreSpaces(sJSON, index)
  While not parserPeek("]", sJSON, index)
    Dim value: Call parseValue(value, sJSON, index)
    json.add value
    Call parserIgnoreSpaces(sJSON, index)
    if parserOptConsume(",", sJSON, index) then Call parserIgnoreSpaces(sJSON, index)
  Wend
  Call parserConsume("]", sJSON, index)
  set parseJSONArray = json
End Function

'Parser: Parse a JSON object key
'@param sJSON - The JSON string to parse
'@param index - The index to start parsing from
'@returns - The parsed JSON object key
Private Function parseJSONObjectKey(ByRef sJSON as string, ByRef index as Long) as string
  Dim char as string: char = parserPeekStr(sJSON, index)
  if char = """" then 
    parseJSONObjectKey = parseString(sJSON, index)
  elseif this.parserProps.AllowUnquotedKeys then
    Dim sKey as string: sKey = ""
    Do While True
      char = parserConsumeAny(sJSON, index)
      if (char <> " ") And (char <> ":") then
        sKey = sKey & char
      else
        parseJSONObjectKey = sKey
        Exit Do
      end if
    Loop
  end if
End Function

'Parser: Parse a JSON string
'@param sJSON - The JSON string to parse
'@param index - The index to start parsing from
'@returns - The parsed JSON string
Private Function parseString(ByRef sJSON as string, Optional ByRef index as Long = 1) as String
  Dim buffer as StringBuffer: buffer = StringBuffer_Create()
  Dim char as string
  
  Dim sQuoteType as string
  if parserOptConsume("""", sJSON, index) then 
    sQuoteType = """"
  elseif this.parserProps.AllowSingleQuoteStrings then
    Call parserConsume("'", sJSON, index)
    sQuoteType = "'"
  end if
  
  While not parserPeek(sQuoteType, sJSON, index)
    char = parserConsumeAny(sJSON,index)
    select case char
      case "\"
        char = parserConsumeAny(sJSON, index)
        select case char
          case """", "\", "/", sQuoteType: Call StringBuffer_Append(buffer, char)
          case "b": Call StringBuffer_Append(buffer, vbBack)
          case "f": Call StringBuffer_Append(buffer, vbFormFeed)
          case "n": Call StringBuffer_Append(buffer, vbLf)
          case "r": Call StringBuffer_Append(buffer, vbCr)
          case "t": Call StringBuffer_Append(buffer, vbTab)
          case "u"
            Dim WCode as string: WCode = parserConsumeAny(sJSON, index, 4)
            Call StringBuffer_Append(buffer, ChrW("&H" & WCode))
        end select
      case else
        Call StringBuffer_Append(buffer, char)
    end select
  Wend
  Call parserConsume(sQuoteType, sJSON, index)
  parseString = StringBuffer_ToString(buffer)
End Function

'Parser: Parse a JSON value
'@param value - The value to parse
'@param sJSON - The JSON string to parse
'@param index - The index to start parsing from
Private Sub parseValue(ByRef value as Variant, ByRef sJSON as string, ByRef index as Long)
  Dim char as string: char = parserPeekStr(sJSON, index)
  select case char
    case "{"
      set value = parseJSONObject(sJSON, index)
    case "["
      set value = parseJSONArray(sJSON, index)
    case """","'"
      value = parseString(sJSON, index)
    case "t"
      Call parserConsume("true", sJSON, index, 4)
      value = true
    case "f"
      Call parserConsume("false", sJSON, index, 5)
      value = false
    case "n"
      Call parserConsume("null", sJSON, index, 4)
      value = null
    case "1", "2", "3", "4", "5", "6", "7", "8", "9", "-"
      value = parseNumber(sJSON, index)
  end select
End Sub

'Parser: Parse a JSON number
'@param sJSON - The JSON string to parse
'@param index - The index to start parsing from
'@returns - The parsed JSON number
Private Function parseNumber(ByRef sJSON as string, ByRef index as Long) as Double
  Dim char as string: char = parserPeekStr(sJSON, index)
  Dim sNum as string: sNum = ""
  while InStr("+-0123456789.eE", char) 'Test char
    index = index + 1                  'Move to next char
    sNum = sNum & char                 'Generate number
    char = parserPeekStr(sJSON, index) 'Prepare next char
  wend
  
  'Handle large numbers losing precision
  if this.parserProps.StoreLargeNumsAsString and len(sNum) >= 16 then
    parseNumber = sNum
  else
    parseNumber = Val(sNum)
  end if
End Function


'Parser helper: Ignore spaces in JSON string
'@param sJSON - The JSON string to parse
'@param index - The index to start parsing from
Private Sub parserIgnoreSpaces(ByRef sJSON as string, ByRef index as Long)
  While Mid(sJSON, index, 1) = " "
    index = index + 1
  Wend
End Sub

'Parser helper: Consumes some characters from the JSON string only if they match the string
'@param sChar - The string to consume
'@param sJSON - The JSON string to parse
'@param index - The index to start parsing from
'@param iLength - The length of the string to consume
Private Sub parserConsume(ByVal sChar As String, ByRef sJSON as string, ByRef index as Long, Optional ByVal iLength as Long = 1)
    If mid(sJSON,index,iLength) = sChar Then
      index = index + iLength
    Else
      Err.Raise 1, "", "JSONParseError: Unexpected char, found: " & mid(sJSON,index,iLength) & " but expected: " & sChar
    End If
End Sub

'Parser helper: Consumes some characters from the JSON string
'@param sJSON - The JSON string to parse
'@param index - The index to start parsing from
'@param iLength - The length of the string to consume
'@returns - The consumed string
Private Function parserConsumeAny(ByRef sJSON as string, ByRef index as Long, Optional ByVal iLength as Long = 1) as string
  parserConsumeAny = mid(sJSON,index,iLength)
  index = index + iLength
End Function

'Parser helper: Checks whether the token at iTokenIndex is of the given type
'@param sChar - The string to check
'@param sJSON - The JSON string to parse
'@param index - The index to start parsing from
'@param iLength - The length of the string to check
'@returns - Whether the token at iTokenIndex is of the given type
Private Function parserPeek(ByVal sChar As String, ByRef sJSON as string, ByRef index as Long, Optional ByVal iLength as Long = 1) As Boolean
    parserPeek = mid(sJSON,index,iLength) = sChar
End Function

'Parser helper: Obtains a character without consuming the token
'@param sJSON - The JSON string to parse
'@param index - The index to start parsing from
'@returns - The character at iTokenIndex
Private Function parserPeekStr(ByRef sJSON as string, ByRef index as Long) As String
    parserPeekStr = mid(sJSON,index,1)
End Function

'Parser helper: Combines peek and consume, consuming a token only if matched, without throwing an error if not
'@param sChar - The string to consume
'@param sJSON - The JSON string to parse
'@param index - The index to start parsing from
'@returns - Whether the token at iTokenIndex is of the given type
Private Function parserOptConsume(ByVal sChar As String, ByRef sJSON as string, ByRef index as Long) As Boolean
    Dim matched as boolean: matched = parserPeek(sChar, sJSON, index)
    If matched Then
        Call parserConsume(sChar, sJSON, index)
    End If
    parserOptConsume = matched
End Function


'Create a stringbuffer
'@param Initial - The initial size of the buffer
'@returns - The stringbuffer
Private Function StringBuffer_Create(Optional ByVal Initial as Long = 16) as StringBuffer
  StringBuffer_Create.value = String$(Initial, " ")
  StringBuffer_Create.count = Initial
  StringBuffer_Create.index = 1
End Function

'Convert a stringbuffer to a string
'@param buffer - The stringbuffer to convert
'@returns - The stringbuffer as a string
Private Function StringBuffer_ToString(ByRef buffer as StringBuffer) as string
  StringBuffer_ToString = left(buffer.value, buffer.index - 1)
End Function

'Append a string to a stringbuffer
'@param buffer - The stringbuffer to append to
'@param sToAppend - The string to append
'@remark If the stringbuffer is not large enough, it will be upsized
Private Sub StringBuffer_Append(ByRef buffer as StringBuffer, ByVal sToAppend as string)
  Dim lenToAppend as Long: lenToAppend = Len(sToAppend)
  if buffer.index + lenToAppend > buffer.count then
    Call StringBuffer_Upsize(buffer, lenToAppend)
  end if
  mid(buffer.value, buffer.index, lenToAppend) = sToAppend
  buffer.index = buffer.index + lenToAppend
End Sub

'Upsize a stringbuffer
'@param buffer - The stringbuffer to upsize
'@param minIncrease - The minimum amount to increase the buffer by
Private Sub StringBuffer_Upsize(ByRef buffer as StringBuffer, Optional ByVal minIncrease as Long)
  minIncrease = iif(buffer.count > minIncrease, buffer.count, minIncrease)
  buffer.count = buffer.count + minIncrease
  buffer.value = buffer.value & string$(minIncrease, " ")
End Sub