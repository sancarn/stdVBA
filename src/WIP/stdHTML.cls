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
  ERaw
End Enum
Private Type TElement
  TagName As String
  Attributes() As Variant
  Children() As stdHTML
  isSelfClosing As Boolean
End Type
Private Type TLiteral
  content as string
End Type
Private Type TRaw
  content as string
End Type

Private Type TThis
  ElementType As EElementType
  Element as TElement
  Literal as TLiteral
  Raw as TRaw
End Type
Private This As TThis

Public Enum ECDNType
  javascript
  css
End Enum

'Creates a new stdHTML object. This is the main entry point for creating HTML document objects.
Public Function Create() as stdHTML
  set Create = New stdHTML
  Call Create.protInit("html", Array(), False)
End Function

' Creates a new stdHTML object with the specified tag name and attributes, and adds it to the current element.
'@param tagName - Name of the tag to create
'@param attributes - Attributes to add to the tag
'@param isSelfClosing - Whether the tag is self-closing
Public Function CreateChild(ByVal tagName as string, ByVal attributes as variant, optional byval isSelfClosing as Boolean = false) as stdHTML
  set CreateChild = new stdHTML
  Call CreateChild.protInit(tagName, attributes, isSelfClosing)
  Call AppendChild(CreateChild)
End Function

'Creates a literal element and adds it to the current element.
'@param text - Text to add
Public Sub CreateLiteral(ByVal text as string)
  Dim newLit as stdHTML: set newLit = new stdHTML
  Call newLit.protInitLiteral(text)
  Call AppendChild(newLit)
End Sub

'Creates a raw literal element and adds it to the current element.
'@param text - Raw text to add
Public Sub CreateRaw(ByVal text as string)
  Dim newLit as stdHTML: set newLit = new stdHTML
  Call newLit.protInitRaw(text)
  Call AppendChild(newLit)
End Sub

'Alias for CreateRaw.
'@param text - Raw text to add
Public Sub Raw(ByVal text as string)
  Call CreateRaw(text)
End Sub

'Creates a CDN resource element and adds it to the current element.
'@param cdnType - Type of CDN element to create
'@param link - URL of the CDN resource
Public Function CreateCDNElement(ByVal cdnType As ECDNType, ByVal link As String) As stdHTML
  If Len(link) = 0 Then
    Err.Raise vbObjectError + 3, "stdHTML", "CDN link cannot be empty."
  End If

  Select Case cdnType
    Case ECDNType.javascript
      Set CreateCDNElement = CreateChild("script", Array("src", link))
    Case ECDNType.css
      Set CreateCDNElement = CreateChild("link", Array("rel", "stylesheet", "href", link), True)
    Case Else
      Err.Raise vbObjectError + 3, "stdHTML", "Unsupported CDN type: " & CStr(cdnType)
  End Select
End Function

'Internal initializer shared by factory methods.
Friend Sub protInit(Optional ByVal tagName As String = "html", Optional ByVal attributes As Variant, Optional ByVal isSelfClosing As Boolean = False)
  Dim normalizedAttributes As Variant
  If IsMissing(attributes) Then
    normalizedAttributes = Array()
  Else
    normalizedAttributes = attributes
  End If
  Call protInitElement(tagName, normalizedAttributes, isSelfClosing)
End Sub

'Initialise an element object.
'@param tagName - Name of the tag to create
'@param attributes - Attributes to add to the tag
'@param isSelfClosing - Whether the tag is self-closing
Public Sub protInitElement(tagName as string, attributes as variant, isSelfClosing as Boolean)
    This.ElementType = EElementType.EElement
    This.Element.TagName = tagName
    This.Element.Attributes = NormalizeAttributes(attributes)
    Erase This.Element.Children
    This.Element.isSelfClosing = isSelfClosing
End Sub

'Initialise a literal object.
'@param content - Content of the literal element
Public Sub protInitLiteral(content as string)
    This.ElementType = EElementType.ELiteral
    This.Literal.content = content
End Sub

'Initialise a raw object.
'@param content - Content of the raw element
Public Sub protInitRaw(content as string)
    This.ElementType = EElementType.ERaw
    This.Raw.content = content
End Sub



Public Property Get TagName() as string
  Select Case this.ElementType
    Case EElementType.EElement
      TagName = This.Element.TagName
    Case EElementType.ELiteral
      TagName = ""
    Case EElementType.ERaw
      TagName = ""
    case else
      Err.Raise vbObjectError + 1, "Invalid element type."
  End Select
End Property

Public Property Get Attr(name as string) as variant
  If this.ElementType <> EElementType.EElement Then Err.Raise vbObjectError + 1, "stdHTML", "Cannot get attributes on a non-element context."

  Dim attrUB As Long: attrUB = UBoundAttributes()
  If attrUB < 1 Then
    Attr = Empty
    Exit Property
  End If

  Dim i as Long
  For i = 0 to attrUB step 2
    if This.Element.Attributes(i) = name then
      Attr = This.Element.Attributes(i + 1)
      Exit Property
    end if
  next i
  Attr = Empty 'If the attribute doesn't exist, return empty
End Property
Public Property Let Attr(name as string, value as variant)
  If this.ElementType <> EElementType.EElement Then Err.Raise vbObjectError + 1, "stdHTML", "Cannot set attributes on a non-element context."

  Dim attrUB As Long: attrUB = UBoundAttributes()
  Dim i as Long
  For i = 0 to attrUB step 2
    if This.Element.Attributes(i) = name then
      This.Element.Attributes(i + 1) = value
      Exit Property
    end if
  next i

  'If the attribute doesn't exist, add it
  If attrUB < 0 Then
    ReDim This.Element.Attributes(0 To 1)
    This.Element.Attributes(0) = name
    This.Element.Attributes(1) = value
  Else
    ReDim Preserve This.Element.Attributes(0 To attrUB + 2)
    This.Element.Attributes(attrUB + 1) = name
    This.Element.Attributes(attrUB + 2) = value
  End If
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
'@param classes as Variant<string>() - Classes to add
Public Sub AddClasses(ParamArray classes())
  If this.ElementType <> EElementType.EElement then Err.Raise vbObjectError + 1, "Cannot add classes to a non-element context."

  Dim classList As String
  classList = NormalizeWhitespace(CoalesceString(Attr("class")))

  Dim newClass As Variant
  For Each newClass In classes
    Dim className As String
    className = NormalizeWhitespace(CStr(newClass))
    If Len(className) > 0 Then
      If Not ContainsClass(classList, className) Then
        If Len(classList) = 0 Then
          classList = className
        Else
          classList = classList & " " & className
        End If
      End If
    End If
  Next

  Attr("class") = classList
End Sub

'Removes the specified classes from the element
'@param classes as Variant<string>() - Classes to remove
Public Sub RemoveClasses(ParamArray classes())
  If this.ElementType <> EElementType.EElement then Err.Raise vbObjectError + 1, "Cannot remove classes from a non-element context."

  Dim classList As String
  classList = NormalizeWhitespace(CoalesceString(Attr("class")))
  If Len(classList) = 0 Then Exit Sub

  Dim remClass as variant
  For each remClass in classes
    classList = RemoveClass(classList, NormalizeWhitespace(CStr(remClass)))
  next

  Attr("class") = NormalizeWhitespace(classList)
End Sub

'Obtain the HTML string representation of the element.
'@param indent - Indentation string to use for nested elements
Public Function ToString(Optional indent As String = "") As String
  select case this.ElementType
    case EElementType.EElement
      Dim output As String: output = indent & "<" & this.Element.TagName

      Dim i As Long
      Dim attrUB As Long: attrUB = UBoundAttributes()
      For i = 0 To attrUB Step 2
        output = output & " " & this.Element.Attributes(i) & serializeAttribute(this.Element.Attributes(i + 1))
      Next i

      If this.Element.isSelfClosing Then
        output = output & " />"
      Else
        output = output & ">"

        Dim childUB As Long: childUB = UBoundChildren()
        If childUB >= 0 Then
          For i = 0 To childUB
            If Not this.Element.Children(i) Is Nothing Then
              output = output & vbCrLf & this.Element.Children(i).ToString(indent & "  ")
            End If
          Next i
          output = output & vbCrLf & indent
        End If

        output = output & "</" & this.Element.TagName & ">"
      End If

      If Len(indent) = 0 And StrComp(this.Element.TagName, "html", vbTextCompare) = 0 Then
        ToString = "<!DOCTYPE html>" & vbCrLf & output
      Else
        ToString = output
      End If
    case EElementType.ELiteral
      ToString = indent & EscapeHtmlText(this.Literal.content)
    case EElementType.ERaw
      ToString = indent & This.Raw.content
    case else
      Err.Raise vbObjectError + 1, "stdHTML", "Invalid element type."
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
      serializeAttribute = "=""" & EscapeHtmlAttribute(CStr(attr)) & """"
    case vbBoolean
      serializeAttribute = "=""" & IIf(attr, "true", "false") & """"
    case vbLong, vbInteger
      serializeAttribute = "=" & CStr(attr)
    case vbDouble, vbSingle
      serializeAttribute = "=" & CStr(attr)
    case vbDate
      serializeAttribute = "=""" & EscapeHtmlAttribute(format$(attr, "yyyy-mm-dd")) & """"
    case vbCurrency
      serializeAttribute = "=" & CStr(attr)
    case vbByte
      serializeAttribute = "=" & CStr(attr)
    case else
      Err.Raise vbObjectError + 2, "stdHTML", "Unsupported attribute type: " & VarType(attr)
  end select
End Function

Public Sub AppendChild(ByVal child As stdHTML)
  If this.ElementType <> EElementType.EElement Then
    Err.Raise vbObjectError + 1, "stdHTML", "Cannot add child to a non-element context."
  End If

  Dim childUB As Long: childUB = UBoundChildren()
  If childUB < 0 Then
    ReDim This.Element.Children(0 To 0)
    Set This.Element.Children(0) = child
  Else
    ReDim Preserve This.Element.Children(0 To childUB + 1)
    Set This.Element.Children(childUB + 1) = child
  End If
End Sub

Private Function NormalizeAttributes(ByVal attributes As Variant) As Variant
  If IsArray(attributes) Then
    Dim sourceUB As Long: sourceUB = SafeUBoundVariant(attributes)
    If sourceUB < 0 Then
      NormalizeAttributes = Array()
      Exit Function
    End If

    Dim targetUB As Long: targetUB = sourceUB
    If (targetUB Mod 2) = 0 Then targetUB = targetUB + 1

    Dim normalized() As Variant
    ReDim normalized(0 To targetUB)

    Dim i As Long
    For i = 0 To sourceUB
      normalized(i) = attributes(i)
    Next i

    If sourceUB <> targetUB Then normalized(targetUB) = Empty
    NormalizeAttributes = normalized
    Exit Function
  End If

  If IsEmpty(attributes) Or IsNull(attributes) Then
    NormalizeAttributes = Array()
    Exit Function
  End If

  Err.Raise vbObjectError + 2, "stdHTML", "Attributes must be key/value array pairs."
End Function

Private Function CoalesceString(ByVal value As Variant) As String
  If IsEmpty(value) Or IsNull(value) Then Exit Function
  CoalesceString = CStr(value)
End Function

Private Function ContainsClass(ByVal classList As String, ByVal className As String) As Boolean
  If Len(classList) = 0 Or Len(className) = 0 Then Exit Function

  Dim parts() As String: parts = Split(classList, " ")
  Dim i As Long
  For i = LBound(parts) To UBound(parts)
    If StrComp(parts(i), className, vbTextCompare) = 0 Then
      ContainsClass = True
      Exit Function
    End If
  Next i
End Function

Private Function RemoveClass(ByVal classList As String, ByVal className As String) As String
  If Len(classList) = 0 Or Len(className) = 0 Then
    RemoveClass = classList
    Exit Function
  End If

  Dim parts() As String: parts = Split(classList, " ")
  Dim i As Long
  Dim output As String

  For i = LBound(parts) To UBound(parts)
    If Len(parts(i)) > 0 Then
      If StrComp(parts(i), className, vbTextCompare) <> 0 Then
        If Len(output) = 0 Then
          output = parts(i)
        Else
          output = output & " " & parts(i)
        End If
      End If
    End If
  Next i

  RemoveClass = output
End Function

Private Function NormalizeWhitespace(ByVal value As String) As String
  value = Replace(value, vbTab, " ")
  value = Replace(value, vbCr, " ")
  value = Replace(value, vbLf, " ")
  Do While InStr(value, "  ") > 0
    value = Replace(value, "  ", " ")
  Loop
  NormalizeWhitespace = Trim$(value)
End Function

Private Function EscapeHtmlText(ByVal value As String) As String
  value = Replace(value, "&", "&amp;")
  value = Replace(value, "<", "&lt;")
  value = Replace(value, ">", "&gt;")
  EscapeHtmlText = value
End Function

Private Function EscapeHtmlAttribute(ByVal value As String) As String
  value = EscapeHtmlText(value)
  value = Replace(value, """", "&quot;")
  value = Replace(value, "'", "&#39;")
  EscapeHtmlAttribute = value
End Function

Private Function SafeUBoundVariant(ByVal values As Variant) As Long
  On Error GoTo noValues
  SafeUBoundVariant = UBound(values)
  Exit Function

noValues:
  SafeUBoundVariant = -1
End Function

Private Function UBoundAttributes() As Long
  On Error GoTo noAttributes
  UBoundAttributes = UBound(This.Element.Attributes)
  Exit Function

noAttributes:
  UBoundAttributes = -1
End Function

Private Function UBoundChildren() As Long
  On Error GoTo noChildren
  UBoundChildren = UBound(This.Element.Children)
  Exit Function

noChildren:
  UBoundChildren = -1
End Function
