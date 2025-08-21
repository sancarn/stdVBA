VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Enum EElementType
  ENone
  EElement
  ELiteral
End Enum
Private Type TElement
  TagName As String
  Attributes As Variant()
  Children As stdHTML()
  isSelfClosing As Boolean
End Type
Private Type TLiteral
  content as string
End Type

Private Type TThis
  ElementType As EElementType
  Element as TElement
  Literal as TLiteral
End Type
Private This As TThis

'Creates a new stdHTML object. This is the main entry point for creating HTML document objects.
Public Function Create() as stdHTML
  set Create = New stdHTML
  Call Create.protInit("html", Array())
End Function

' Creates a new stdHTML object with the specified tag name and attributes, and adds it to the current element.
'@param tagName - Name of the tag to create
'@param attributes - Attributes to add to the tag
'@param isSelfClosing - Whether the tag is self-closing
Public Function CreateChild(ByVal tagName as string, ByVal attributes as variant, optional byval isSelfClosing as Boolean = false) as stdHTML
  set CreateChild = new stdHTML
  Call CreateChild.protInit(tagName, attributes, isSelfClosing)

  'Theoretically someone could call this function on a literal element, or the predeclared stdHTML object.
  If this.ElementType = EElementType.EElement Then
    Dim iUB as long: iUB = UBound(This.Element.Children)
    ReDim Preserve This.Element.Children(iUB + 1)
    Set This.Element.Children(iUB + 1) = CreateChild
  else
    Err.Raise vbObjectError + 1, "Cannot add child to a non-element context."
  End if
End Function

'Creates a literal element and adds it to the current element.
'@param text - Text to add
Public Sub CreateLiteral(ByVal text as string)
  Dim newLit as stdHTML: set newLit = new stdHTML
  Call newLit.protInitLiteral(text)
  
  'Theoretically someone could call this function on a literal element, or the predeclared stdHTML object.
  If this.ElementType = EElementType.EElement Then
    Dim iUB as long: iUB = UBound(This.Element.Children)
    ReDim Preserve This.Element.Children(iUB + 1)
    Set This.Element.Children(iUB + 1) = newLit
  else
    Err.Raise vbObjectError + 1, "Cannot add literal to a non-element context."
  End If
End Sub


'Initialise an element object.
'@param tagName - Name of the tag to create
'@param attributes - Attributes to add to the tag
'@param isSelfClosing - Whether the tag is self-closing
Public Sub protInitElement(tagName as string, attributes as variant, isSelfClosing as Boolean)
    This.ElementType = EElementType.EElement
    This.Element.TagName = tagName
    This.Element.Attributes = attributes
    This.Element.isSelfClosing = isSelfClosing
End Sub

'Initialise a literal object.
'@param content - Content of the literal element
Public Sub protInitLiteral(content as string)
    This.ElementType = EElementType.ELiteral
    This.Literal.content = content
End Sub



Public Property Get TagName() as string
  Select Case this.ElementType
    Case EElementType.EElement
      TagName = This.Element.TagName
    Case EElementType.ELiteral
      TagName = ""
    case else
      Err.Raise vbObjectError + 1, "Invalid element type."
  End Select
End Property

Public Property Get Attr(name as string) as variant
  Dim i as Long
  For i = 0 to UBound(This.Element.Attributes) step 2
    if This.Element.Attributes(i) = name then
      Attr = This.Element.Attributes(i + 1)
      exit for
    end if
  next i
  Attr = Empty 'If the attribute doesn't exist, return empty
End Property
Public Property Let Attr(name as string, value as variant)
  Dim i as Long
  For i = 0 to UBound(This.Element.Attributes) step 2
    if This.Element.Attributes(i) = name then
      This.Element.Attributes(i + 1) = value
      Exit Property
    end if
  next i

  'If the attribute doesn't exist, add it
  Dim iUB as Long: iUB = UBound(This.Element.Attributes)
  ReDim Preserve This.Element.Attributes(iUB + 2)
  This.Element.Attributes(iUB + 1) = name
  This.Element.Attributes(iUB + 2) = value
End Property


Public Property Get Id() as string
  ID = Attr("id")
End Property
Public Property Let ID(value As String)
  Attr("id") = value
End Property

Public Property Get Title() as string
  Title = Attr("title")
End Property
Public Property Let Title(value As String)
  Attr("title") = value
End Property

Public Property Get Children() As stdHTML()
  If this.ElementType <> EElementType.EElement then Err.Raise vbObjectError + 1, "Cannot get children of a non-element context."
  Children = This.Element.Children
End Property

'Adds the specified classes to the element
'@param classes - Classes to add
Public Sub AddClasses(ParamArray classes() as string)
  If this.ElementType <> EElementType.EElement then Err.Raise vbObjectError + 1, "Cannot add classes to a non-element context."

  Dim classes as string: classes = Attr("class")
  If isEmpty(classes) then
    Attr("class") = Join(classes, " ")
  Else
    Dim newClass as variant
    For each newClass in classes
      If InStr(1, classes, newClass, vbTextCompare) = 0 then
        classes = classes & " " & newClass
      end if
    next
    Attr("class") = classes
  end if
End Sub

'Removes the specified classes from the element
'@param classes - Classes to remove
Public Sub RemoveClasses(ParamArray classes() as string)
  If this.ElementType <> EElementType.EElement then Err.Raise vbObjectError + 1, "Cannot remove classes from a non-element context."
  
  Dim classes as string: classes = Attr("class")
  If not isEmpty(classes) then
    Dim remClass as variant
    For each remClass in classes
      If InStr(1, classes, remClass, vbTextCompare) <> 0 then
        classes = Replace(classes, newClass, "")
      end if
    next

    Attr("class") = Replace(classes, "  ", " ")
  end if
End Sub

'Obtain the HTML string representation of the element.
'@param indent - Indentation string to use for nested elements
Public Function ToString(Optional indent As String = "") As String
  select case this.ElementType
    case EElementType.EElement
      Dim output As String: output = indent & "<" & this.Element.TagName
      
      Dim i As Long
      For i = 0 To UBound(this.Element.Attributes) Step 2
        output = output & " " & this.Element.Attributes(i) & serializeAttribute(this.Element.Attributes(i + 1))
      Next i
      
      If this.Element.isSelfClosing Then
        output = output & " />"
      Else
        output = output & ">"
        
        Dim child As stdHTML
        For Each child In this.Element.Children
          output = output & vbCrLf & child.ToString(indent & "  ")
        Next child
        
        output = output & vbCrLf & indent & "</" & this.Element.TagName & ">"
      End If
      
      ToString = output
    case EElementType.ELiteral
      ToString = this.Literal.content
  end Select
End Function

'Serializes an attribute to a string
'@param attr - Attribute to serialize
'@returns Serialized attribute
Private Function serializeAttribute(ByRef attr as Variant) as string
  select case varType(attr)
    case vbNull
      serializeAttribute = ""
    case vbString
      serializeAttribute = "='" & attr & "'"
    case vbBoolean
      serializeAttribute = "=" & IIf(attr, "true", "false")
    case vbLong, vbInteger
      serializeAttribute = "=" & CStr(attr)
    case vbDouble, vbSingle
      serializeAttribute = "=" & CStr(attr)
    case vbDate
      serializeAttribute = "='" & format(attr, "yyyy-mm-dd") & "'"
    case vbCurrency
      serializeAttribute = "=" & CStr(attr)
    case vbByte
      serializeAttribute = "=" & CStr(attr)
    case else
      Err.Raise vbObjectError + 2, "Unsupported attribute type: " & VarType(attr) & ". For: " & attr
  end select
End Function
