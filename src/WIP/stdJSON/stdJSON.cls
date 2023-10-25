VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdCallback2"
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
'	Dim o as stdJSON: set o = .obj
'   'Get as dictionary of collections
'   Dim o as object: set o = .toVBON()
'   'Get json string
'   Dim s as string: s = .toString()
' End With

Private Enum EStdJsonType
	eJSONObject
	eJSONArray
End Enum
Private Type TChild
	key as string
	value as variant
End Type
Private Type TThis
	iType as EType
	children() as TChild
End Type
Private This as TThis

Public Function Create(Optional ByVal iType as EStdJsonType = eJSONObject) as stdJSON
    set Create = new stdJSON
    Call Create.protInit(iType)
End Function

Public Function CreateFromString(ByVal sJSON as string) as stdJSON
	Err.Raise 1, "stdJSON::CreateFromParse()", "TODO: Not implemented"
End Function

Public Function CreateFromFile(ByVal sJSONFile as string) as stdJSON
	Err.Raise 1, "stdJSON::CreateFromFile()", "TODO: Not implemented"
End Function

Public Function CreateFromVBObject(ByVal obj as Object) as stdJSON
	Err.Raise 1, "stdJSON::CreateFromVBObject()", "TODO: Not implemented"
End Function

Friend Sub protInit(ByRef iType as EStdJsonType)
    This.iType = iType
End Sub

'Add a key and value to an object, or a value to an array
'@param {ParamArray}
'@returns {stdJSON} Either `Me` or if value was a stdJSON, this is returned.
'@remark Uses `AddForArray` when `This.iType == eJSONArray`; and `AddForObject` when `This.iType == eJSONObject`.
Public Function Add(ParamArray v()) as stdJSON
	select case This.iType
		case eJSONObject
			set Add = AddForObject(v(0), v(1))
		case eJSONArray
			set Add = AddForArray(v(0))
	end select
End Function

'Add a value to the JSON array
'@param {variant} Value - The value to add to the array
'@returns {stdJSON} if the Value is a stdJSON, this value is returned. Else the current stdJSON object is returned.
Public Function AddForArray(ByVal value as variant) as stdJSON
	Dim iNew as Long: iNew = Length+1
    Redim preserve children(1 to iNew)
    Call CopyVariant(children(iNew).value, value)
	if typeof value is stdJSON then
		set AddForArray = value
	else
		set AddForArray = me
	end if
End Function

'Add a key and value to the JSON object
'@param {string} Key - The key to add the object under
'@param {variant} Value - The value to assign to this key
'@returns {stdJSON} if the Value is a stdJSON, this value is returned. Else the current stdJSON object is returned.
Public Function AddForObject(ByVal key as string, ByVal value as variant) as stdJSON
	Dim iNew as Long: iNew = Length+1
    Redim preserve children(1 to iNew)
    With children(iNew)
        .key = key
        Call CopyVariant(.value, value)
    End With
	
	if typeof value is stdJSON then
		set AddForArray = value
	else
		set AddForArray = me
	end if
End Function

'Create and add an array to the JSON object/array.
'@param {string} If `this.iType == eJSONObject` then this is the key it should be assigned to. Else this param is unused.
'@returns {stdJSON} The array created
Public Function AddArray(Optional ByVal key as string = "") as stdJSON
	Dim oRet as stdJSON: set oRet = stdJSON.Create(eJSONArray)
	select case This.iType
		case eJSONObject
			Call AddForObject(key, oRet)
		case eJSONArray
			Call AddForArray(oRet)
	end select
	set AddArray = oRet
End Function

'Create and add an object to the JSON object/array.
'@param {string} If `this.iType == eJSONObject` then this is the key it should be assigned to. Else this param is unused.
'@returns {stdJSON} The object created
Public Function AddObject(Optional ByVal key as string = "") as stdJSON
	Dim oRet as stdJSON: set oRet = stdJSON.Create(eJSONObject)
	select case This.iType
		case eJSONObject
			Call AddForObject(key, oRet)
		case eJSONArray
			Call AddForArray(oRet)
	end select
	set AddObject = oRet
End Function

'Obtain the length of this JSON object
'@returns {Long} The length of the JSON object
Public Property Get Length() as Long
    Length = ubound(this.children) - lbound(this.children) + 1
End Property

'Obtain the item for a key or index
'@param {Variant} Either a string key (if eJSONObject) or index as long (if eJSONArray)
Public Property Get Item(ByVal keyOrIndex as Variant) as Variant
    select case this.iType
        case eJSONObject
            Call CopyVariant(Item, ItemForObject(keyOrIndex))
        case eJSONArray
            Call CopyVariant(Item, ItemForArray(keyOrIndex))
    end select
End Property

'Obtain the item for a key in a eJSONObject
'@param {string} A string key
'@returns {Variant} The item for this key
'TODO: Optimise as HashMap?
Public Property Get ItemForObject(ByVal key as string) as Variant
    if this.iType = eJSONArray then Err.Raise 1, "stdJSON::ItemForObject", "ERROR: This is an array, use ItemForArray"
    Dim i as long
    For i = 1 to Length
        if this.children(i).key = key then
            Call CopyVariant(ItemForObject, this.children(i).value)
        end if
    next
End Property

'Obtain the item for a index in an eJSONArray
'@param {Long} 1-based Index of item
'@returns {Variant} The item at this index
Public Property Get ItemForArray(ByVal index as long) as Variant
    if this.iType = eJSONObject then Err.Raise 1, "stdJSON::ItemForObject", "ERROR: This is an array, use ItemForArray"
    Call CopyVariant(ItemForArray, this.children(index).value)
End Property

'Convert the JSON to string
'@returns {String} JSON encoded data.
Public Function ToString() as String
	Dim s as string
	select case This.iType
		case eJSONObject
			s = s & "{"
		case eJSONArray
			s = s & "["
	end select
	
	Dim ub as Long: ub = ubound(this.children)
	For i = 1 to ub
		With this.children(i)
			Dim sValue as string
			select case vartype(.value)
				case vbObject
					select case typename(.value)
						case "stdJSON"
							sValue = .value.toString()
                        'TODO: add dictionary?
                        'TODO: add collection?
						case else
							Err.Raise 1, "stdJSON::ToString()", "Unsupported type"
					end select
				case else
					sValue = simpleStringify(.value)
			end select
			
			Dim sLine as string: sLine = """" & .key & """:" & sValue
			
			s = s & sLine & iif(i<ub, ",", "")
		End With
	next
	
	select case This.iType
		case eJSONObject
			s = s & "}"
		case eJSONArray
			s = s & "]"
	end select
End Function

'Convert the JSON to pretty string
'@returns {String} JSON encoded data.
Public Function ToPrettyString(optional byval indent as long) as String
	Err.Raise 1, "stdJSON::ToPrettyString()", "TODO: Not implemented"
End Function

'Convert the JSON to pretty string
'@returns {String} JSON encoded data.
Public Sub ToFile(ByVal sPath as string, Optional ByVal prettyIndent as Long = 0)
	Err.Raise 1, "stdJSON::ToFile()", "TODO: Not implemented"
End Function

'Convert to VBA Object
'@returns {Object<Collection|Dictionary>} Dictionary if eJSONObject, Collection if eJSONArray
Public Function ToVBObject() as Object
    Dim ret as Object
    select case This.iType
		case eJSONObject
			set ret = CreateObject("Scripting.Dictionary")
		case eJSONArray
			set ret = new Collection
	end select
    
    Dim i as Long
    For i = 1 to Length
        With this.children(i)
            select case This.iType
                case eJSONObject
                    if typeof .value is stdJSON then
                        ret.add .key, .value.ToVBObject()
                    else
                        ret.add .key, .value
                    end if
                case eJSONArray
                    if typeof .value is stdJSON then
                        ret.add .value.ToVBObject()
                    else
                        ret.add .value
                    end if
            end select
        End With
    next
End Function


Private Sub CopyVariant(ByRef dest as Variant, ByVal src as Variant)
    if isObject(src) then
        set dest = src
    else
        let dest = src
    end if
End Sub

Private Function simpleStringify(ByVal value as variant) as string
	select case vartype(value)
		case vbNull
			sValue = "null"
		case vbDate
			sValue = """" & Format(.value, "dd/mm/yyyy") & """"
		case vbString
			sValue = """" & value & """"
        'TODO: add Array?
		case else
			sValue = .value
	end select
End Function

