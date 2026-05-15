VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@module
'@description `stdHTML` provides a fluent, object-based API for building, querying, and transforming HTML in VBA.
'`stdHTML.Create()` creates a new document root (`<html>`), and child nodes can be composed via `CreateChild`.
'Use `CreateLiteral` for escaped text content and `Raw` for unescaped content when embedding trusted markup or script.
'The selector APIs (`QuerySelector`/`QuerySelectorAll`) support CSS-style traversal for reading or mutating node trees.
'@example **1. Build and render HTML**:
'```vb
'Sub mBuildHtml()
'  Dim doc As stdHTML
'  Set doc = stdHTML.Create()
'
'  With doc.CreateChild("body")
'    .CreateChild("h1").CreateLiteral "Hello"
'    .CreateChild("p").CreateLiteral "Generated with stdHTML"
'  End With
'
'  Debug.Print doc.ToString()
'End Sub
'```
'@example **2. Add raw script and query nodes**:
'```vb
'Sub mTest()
'  With stdHTML.Create()
'    With .CreateChild("head")
'      With .CreateChild("script")
'        .Raw "function hello(){"
'        .Raw "  console.log('hello world')"
'        .Raw "}"
'      End With
'    End With
'    With .CreateChild("body")
'      With .CreateChild("div")
'        .ID = "hi"
'        Call .CreateLiteral("hello world")
'      End With
'    End With
'    Call .QuerySelector("body > #hi").replaceWith(...)
'  End With
'End Sub
'```
'@example **3. Using `Self` to pass the current node**:
'```vb
'Sub mTestSelf()
'  With stdHTML.Create()
'    With .CreateChild("body")
'      With .CreateChild("div")
'        Call createCustomElement(.Self)
'      End With
'    End With
'  End With
'End Sub
'
'Private Sub createCustomElement(ByVal node As stdHTML)
'  With node
'    With .CreateChild("div")
'      .ClassName = "custom-card"
'      .CreateChild("span").CreateLiteral "Reusable content"
'    End With
'  End With
'End Sub
'```
#If Win64 Then
  Private Const PTR_SZ As Long = 8
  Private Const NULL_PTR As LongLong = 0^
#Else
  Private Const PTR_SZ As Long = 4
  Private Const NULL_PTR As Long = 0&
#End If

#If VBA7 = 0 Then
  Private Enum LongPtr
    [_]
  End Enum
#End If

#If VBA7 Then
  Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
#Else
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
#End If


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
Private Type TSelectorAtom
  TagName As String
  IsUniversal As Boolean
  Id As String
  Classes() As String
  AttrNames() As String
  AttrHasValue() As Boolean
  AttrValues() As String
End Type
Private Type TSelectorStep
  Combinator As String
  Atom As TSelectorAtom
End Type
Private Type TSelectorGroup
  Steps() As TSelectorStep
End Type

Private Type TThis
  ElementType As EElementType
  Element as TElement
  Literal as TLiteral
  Raw as TRaw
  ParentPtr As LongPtr
End Type
Private Type THTMLParseOptions
  preserveComments As Boolean
End Type
Private This As TThis

Public Enum ECDNType
  javascript
  css
End Enum

'Creates a new stdHTML object. This is the main entry point for creating HTML document objects.
'@constructor
'@param tagName - Name of the tag to create
'@param attributes - Attributes to add to the tag
'@param isSelfClosing - Whether the tag is self-closing
'@returns - New stdHTML object
Public Function Create(Optional ByVal tagName as string = "html", Optional ByVal attributes as variant = Empty, optional byval isSelfClosing as Boolean = false) as stdHTML
  With stdError.getSentry("stdHTML#Create", "tagName", tagName, "attributes", attributes, "isSelfClosing", isSelfClosing)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      set Create = New stdHTML
      if tagName = "html" then
        Call Create.protInit(tagName, attributes, isSelfClosing)
      else
        Call Create.protInitElement(tagName, attributes, isSelfClosing)
      end if
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Creates stdHTML nodes from a HTML string.
'@constructor
'@param html - HTML markup to parse into stdHTML nodes
'@returns - New stdHTML object
Public Function CreateFromHTML(ByVal html As String) As stdHTML
  With stdError.getSentry("stdHTML#CreateFromHTML", "html", html)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim options As THTMLParseOptions
      options = DefaultHTMLParseOptions()
    
      Dim parsed() As stdHTML
      parsed = ParseHTMLNodes(html, options)
    
      Dim parsedUB As Long: parsedUB = SafeUBoundHTMLArray(parsed)
      If parsedUB < 0 Then
        Set CreateFromHTML = Create()
        Exit Function
      End If
    
      If parsedUB = 0 Then
        If Not parsed(0) Is Nothing Then
          If parsed(0).ElementType = EElementType.EElement Then
            Set CreateFromHTML = parsed(0)
            Exit Function
          End If
        End If
      End If
    
      Set CreateFromHTML = Create()
      Dim i As Long
      For i = 0 To parsedUB
        If Not parsed(i) Is Nothing Then
          CreateFromHTML.ChildAppend parsed(i)
        End If
      Next i
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

' Creates a new stdHTML object with the specified tag name and attributes, and adds it to the current element.
'@constructor
'@param tagName - Name of the tag to create
'@param attributes - Attributes to add to the tag
'@param isSelfClosing - Whether the tag is self-closing
'@returns - New stdHTML object
Public Function CreateChild(ByVal tagName as string, Optional ByVal attributes as variant = Empty, optional byval isSelfClosing as Boolean = false) as stdHTML
  With stdError.getSentry("stdHTML#CreateChild", "tagName", tagName, "attributes", attributes, "isSelfClosing", isSelfClosing)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      set CreateChild = new stdHTML
      if isEmpty(attributes) then attributes = Array()
      Call CreateChild.protInit(tagName, attributes, isSelfClosing)
      Call ChildAppend(CreateChild)
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Creates a literal element and adds it to the current element.
'@constructor
'@param text - Text to add
'@returns - New stdHTML object
Public Sub CreateLiteral(ByVal text as string)
  With stdError.getSentry("stdHTML#CreateLiteral", "text", text)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim newLit as stdHTML: set newLit = new stdHTML
      Call newLit.protInitLiteral(text)
      Call ChildAppend(newLit)
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Creates a raw literal element and adds it to the current element.
'@constructor
'@param text - Raw text to add
'@returns - New stdHTML object
Public Sub CreateRaw(ByVal text as string)
  With stdError.getSentry("stdHTML#CreateRaw", "text", text)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim newLit as stdHTML: set newLit = new stdHTML
      Call newLit.protInitRaw(text)
      Call ChildAppend(newLit)
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Alias for CreateRaw.
'@constructor
'@param text - Raw text to add
'@returns - New stdHTML object
Public Sub Raw(ByVal text as string)
  With stdError.getSentry("stdHTML#Raw", "text", text)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Call CreateRaw(text)
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Creates a CDN resource element and adds it to the current element.
'@constructor
'@param cdnType - Type of CDN element to create
'@param link - URL of the CDN resource
'@returns - New stdHTML object
Public Function CreateCDNElement(ByVal cdnType As ECDNType, ByVal link As String) As stdHTML
  With stdError.getSentry("stdHTML#CreateCDNElement", "cdnType", cdnType, "link", link)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If Len(link) = 0 Then
        Err_Raise vbObjectError + 3, "stdHTML", "CDN link cannot be empty."
      End If
    
      Select Case cdnType
        Case ECDNType.javascript
          Set CreateCDNElement = CreateChild("script", Array("src", link))
        Case ECDNType.css
          Set CreateCDNElement = CreateChild("link", Array("rel", "stylesheet", "href", link), True)
        Case Else
          Err_Raise vbObjectError + 3, "stdHTML", "Unsupported CDN type: " & CStr(cdnType)
      End Select
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Internal initializer shared by factory methods.
'@protected
'@param tagName - Name of the tag to create
'@param attributes - Attributes to add to the tag
'@param isSelfClosing - Whether the tag is self-closing
Friend Sub protInit(Optional ByVal tagName As String = "html", Optional ByVal attributes As Variant, Optional ByVal isSelfClosing As Boolean = False)
  With stdError.getSentry("stdHTML#protInit", "tagName", tagName, "attributes", attributes, "isSelfClosing", isSelfClosing)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim normalizedAttributes As Variant
      If IsMissing(attributes) Then
        normalizedAttributes = Array()
      Else
        normalizedAttributes = attributes
      End If
      Call protInitElement(tagName, normalizedAttributes, isSelfClosing)
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Initialise an element object.
'@protected
'@param tagName - Name of the tag to create
'@param attributes - Attributes to add to the tag
'@param isSelfClosing - Whether the tag is self-closing
Public Sub protInitElement(tagName as string, attributes as variant, isSelfClosing as Boolean)
  With stdError.getSentry("stdHTML#protInitElement", "tagName", tagName, "attributes", attributes, "isSelfClosing", isSelfClosing)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
        This.ElementType = EElementType.EElement
        This.Element.TagName = tagName
        This.Element.Attributes = NormalizeAttributes(attributes)
        Erase This.Element.Children
        This.Element.isSelfClosing = isSelfClosing
        This.ParentPtr = 0
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Initialise a literal object.
'@protected
'@param content - Content of the literal element
Public Sub protInitLiteral(content as string)
  With stdError.getSentry("stdHTML#protInitLiteral", "content", content)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
        This.ElementType = EElementType.ELiteral
        This.Literal.content = content
        This.ParentPtr = 0
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Initialise a raw object.
'@protected
'@param content - Content of the raw element
Public Sub protInitRaw(content as string)
  With stdError.getSentry("stdHTML#protInitRaw", "content", content)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
        This.ElementType = EElementType.ERaw
        This.Raw.content = content
        This.ParentPtr = 0
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Returns the current instance for helper calls inside fluent `With` chains.
'@returns - Current instance
Public Property Get Self() as stdHTML
  With stdError.getSentry("stdHTML#Self[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Set Self = Me
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Gets the underlying node type (`EElement`, `ELiteral`, or `ERaw`).
'@returns - Element type
Public Property Get ElementType() As Long
  With stdError.getSentry("stdHTML#ElementType[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      ElementType = This.ElementType
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Gets the tag name for element nodes.
'@returns - Tag name
Public Property Get TagName() as string
  With stdError.getSentry("stdHTML#TagName[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Select Case this.ElementType
        Case EElementType.EElement
          TagName = This.Element.TagName
        Case EElementType.ELiteral
          TagName = ""
        Case EElementType.ERaw
          TagName = ""
        case else
          Err_Raise vbObjectError + 1, "Invalid element type."
      End Select
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Gets an attribute value on an element node.
'@param name - Attribute name
'@returns - Attribute value, or Empty when the attribute is missing
Public Property Get Attr(name as string) as variant
  With stdError.getSentry("stdHTML#Attr[Get]", "name", name)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If this.ElementType <> EElementType.EElement Then Err_Raise vbObjectError + 1, "stdHTML", "Cannot get attributes on a non-element context."
    
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
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property
Public Property Let Attr(name as string, value as variant)
  With stdError.getSentry("stdHTML#Attr[Let]", "name", name, "value", value)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If this.ElementType <> EElementType.EElement Then Err_Raise vbObjectError + 1, "stdHTML", "Cannot set attributes on a non-element context."
    
      Dim attrUB As Long: attrUB = UBoundAttributes()
      Dim i as Long
      For i = 0 to attrUB step 2
        if This.Element.Attributes(i) = name then
          If IsEmpty(value) Then
            RemoveAttributePairAt i
          Else
            This.Element.Attributes(i + 1) = value
          End If
          Exit Property
        end if
      next i
    
      If IsEmpty(value) Then Exit Property
    
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
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

Private Sub RemoveAttributePairAt(ByVal pairStartIndex As Long)
  With stdError.getSentry("stdHTML#RemoveAttributePairAt", "pairStartIndex", pairStartIndex)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim attrUB As Long: attrUB = UBoundAttributes()
      If pairStartIndex < 0 Or pairStartIndex > attrUB - 1 Then Exit Sub
    
      If attrUB <= 1 Then
        Erase This.Element.Attributes
        Exit Sub
      End If
    
      Dim newValues() As Variant
      ReDim newValues(0 To attrUB - 2)
    
      Dim src As Long
      Dim dst As Long
      For src = 0 To attrUB Step 2
        If src <> pairStartIndex Then
          newValues(dst) = This.Element.Attributes(src)
          newValues(dst + 1) = This.Element.Attributes(src + 1)
          dst = dst + 2
        End If
      Next src
    
      This.Element.Attributes = newValues
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub


'Get/Set the `id` attribute.
'@returns - `id` attribute value
Public Property Get Id() as string
  With stdError.getSentry("stdHTML#Id[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      ID = Attr("id")
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property
Public Property Let ID(value As String)
  With stdError.getSentry("stdHTML#ID[Let]", "value", value)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Attr("id") = value
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Get/Set the `title` attribute.
'@returns - `title` attribute value
Public Property Get Title() as string
  With stdError.getSentry("stdHTML#Title[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Title = Attr("title")
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property
Public Property Let Title(value As String)
  With stdError.getSentry("stdHTML#Title[Let]", "value", value)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Attr("title") = value
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Gets child nodes from an element node.
'@returns - Child nodes
Public Property Get Children() As stdHTML()
  With stdError.getSentry("stdHTML#Children[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If this.ElementType <> EElementType.EElement then Err_Raise vbObjectError + 1, "Cannot get children of a non-element context."
      Children = This.Element.Children
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Returns the parent node when available via weak pointer.
'@returns - Parent node
Public Property Get Parent() As stdHTML
  With stdError.getSentry("stdHTML#Parent[Get]")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If This.ParentPtr = 0 Then Exit Property
      Set Parent = ObjFromPtr(This.ParentPtr)
    
    Exit Property
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Property

'Sets the parent pointer for the current node.
'@protected
'@param parentPtr - Pointer to the parent node
Public Sub protSetParentPtr(ByVal parentPtr As LongPtr)
  With stdError.getSentry("stdHTML#protSetParentPtr", "parentPtr", parentPtr)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      This.ParentPtr = parentPtr
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Adds the specified classes to the element
'@param classes as Variant<string>() - Classes to add
Public Sub ClassesAdd(ParamArray classes())
  With stdError.getSentry("stdHTML#ClassesAdd")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If this.ElementType <> EElementType.EElement then Err_Raise vbObjectError + 1, "Cannot add classes to a non-element context."
    
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
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Removes the specified classes from the element
'@param classes as Variant<string>() - Classes to remove
Public Sub ClassesRemove(ParamArray classes())
  With stdError.getSentry("stdHTML#ClassesRemove")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If this.ElementType <> EElementType.EElement then Err_Raise vbObjectError + 1, "Cannot remove classes from a non-element context."
    
      Dim classList As String
      classList = NormalizeWhitespace(CoalesceString(Attr("class")))
      If Len(classList) = 0 Then Exit Sub
    
      Dim remClass as variant
      For each remClass in classes
        classList = RemoveClass(classList, NormalizeWhitespace(CStr(remClass)))
      next
    
      Attr("class") = NormalizeWhitespace(classList)
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Obtain the HTML string representation of the element.
'@param indent - Indentation string to use for nested elements
'@returns - HTML string representation
Public Function ToString(Optional indent As String = "") As String
  With stdError.getSentry("stdHTML#ToString", "indent", indent)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      select case this.ElementType
        case EElementType.EElement
          Dim output As String: output = indent & "<" & this.Element.TagName
    
          Dim i As Long
          Dim attrUB As Long: attrUB = UBoundAttributes()
          For i = 0 To attrUB Step 2
            Dim attrValue As Variant
            attrValue = This.Element.Attributes(i + 1)
    
            ' Empty means the attribute is removed; do not emit name or value.
            If Not IsEmpty(attrValue) Then
              output = output & " " & this.Element.Attributes(i) & serializeAttribute(attrValue)
            End If
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
          Err_Raise vbObjectError + 1, "stdHTML", "Invalid element type."
      end Select
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Serializes an attribute to a string
'@param attr - Attribute to serialize
'@returns Serialized attribute
Private Function serializeAttribute(ByRef attr as Variant) as string
  With stdError.getSentry("stdHTML#serializeAttribute", "attr", attr)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      select case varType(attr)
        case vbEmpty
          serializeAttribute = ""
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
          Err_Raise vbObjectError + 2, "stdHTML", "Unsupported attribute type: " & VarType(attr)
      end select
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Appends a child node to the current element.
'@param child - Child node to append
Public Sub ChildAppend(ByVal child As stdHTML)
  With stdError.getSentry("stdHTML#ChildAppend", "child", child)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If this.ElementType <> EElementType.EElement Then
        Err_Raise vbObjectError + 1, "stdHTML", "Cannot add child to a non-element context."
      End If
      If child Is Nothing Then
        Err_Raise vbObjectError + 5, "stdHTML", "Cannot add a Nothing child."
      End If
    
      Dim childUB As Long: childUB = UBoundChildren()
      If childUB < 0 Then
        ReDim This.Element.Children(0 To 0)
        Set This.Element.Children(0) = child
      Else
        ReDim Preserve This.Element.Children(0 To childUB + 1)
        Set This.Element.Children(childUB + 1) = child
      End If
      child.protSetParentPtr ObjPtr(Me)
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Replaces the current node contents with another node's contents.
'@param replacement - Replacement node
'@returns True when replacement data was applied
Public Function ReplaceWith(ByVal replacement As stdHTML) As Boolean
  With stdError.getSentry("stdHTML#ReplaceWith", "replacement", replacement)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If replacement Is Nothing Then
        Err_Raise vbObjectError + 5, "stdHTML#ReplaceWith", "Replacement node cannot be Nothing."
      End If
    
      If replacement Is Me Then
        ReplaceWith = True
        Exit Function
      End If
    
      Dim replacementState As Object
      Set replacementState = replacement.protGetStateAsDictionary()
      Call protInitFromStateDictionary(replacementState)
    
      ReplaceWith = True
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function


'Creates a serializable snapshot of the current node state.
'@protected
'@returns Dictionary containing node data
Public Function protGetStateAsDictionary() As Object
  With stdError.getSentry("stdHTML#protGetStateAsDictionary")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim state As Object
      Set state = CreateObject("Scripting.Dictionary")
    
      state("ElementType") = CLng(This.ElementType)
    
      Select Case This.ElementType
        Case EElementType.EElement
          state("TagName") = This.Element.TagName
          state("Attributes") = This.Element.Attributes
          state("Children") = This.Element.Children
          state("IsSelfClosing") = This.Element.isSelfClosing
        Case EElementType.ELiteral
          state("Content") = This.Literal.content
        Case EElementType.ERaw
          state("Content") = This.Raw.content
        Case Else
          Err_Raise vbObjectError + 1, "stdHTML#protGetStateAsDictionary", "Invalid element type."
      End Select
    
      Set protGetStateAsDictionary = state
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function


'Initializes the current node from a serialized state dictionary.
'@protected
'@param state - Dictionary produced by `protGetStateAsDictionary`
Public Sub protInitFromStateDictionary(ByVal state As Object)
  With stdError.getSentry("stdHTML#protInitFromStateDictionary", "state", state)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      This.ElementType = CLng(state("ElementType"))
    
      Select Case This.ElementType
        Case EElementType.EElement
          This.Element.TagName = CStr(state("TagName"))
          This.Element.Attributes = state("Attributes")
          This.Element.Children = state("Children")
          This.Element.isSelfClosing = CBool(state("IsSelfClosing"))
          Dim childUB As Long: childUB = UBoundChildren()
          Dim i As Long
          For i = 0 To childUB
            If Not This.Element.Children(i) Is Nothing Then
              This.Element.Children(i).protSetParentPtr ObjPtr(Me)
            End If
          Next i
        Case EElementType.ELiteral
          This.Literal.content = CStr(state("Content"))
          Erase This.Element.Attributes
          Erase This.Element.Children
          This.Element.TagName = vbNullString
          This.Element.isSelfClosing = False
        Case EElementType.ERaw
          This.Raw.content = CStr(state("Content"))
          Erase This.Element.Attributes
          Erase This.Element.Children
          This.Element.TagName = vbNullString
          This.Element.isSelfClosing = False
        Case Else
          Err_Raise vbObjectError + 1, "stdHTML#protInitFromStateDictionary", "Invalid element type."
      End Select
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Returns the first node that matches the selector.
'@param selector - CSS selector string
'@returns First matching node, or Nothing if no nodes match
Public Function QuerySelector(ByVal selector As String) As stdHTML
  With stdError.getSentry("stdHTML#QuerySelector", "selector", selector)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim matches() As stdHTML
      matches = QuerySelectorAll(selector)
    
      Dim ub As Long: ub = SafeUBoundHTMLArray(matches)
      If ub >= 0 Then Set QuerySelector = matches(0)
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Returns all nodes that match the selector.
'@param selector - CSS selector string
'@returns Array of matching nodes
Public Function QuerySelectorAll(ByVal selector As String) As stdHTML()
  With stdError.getSentry("stdHTML#QuerySelectorAll", "selector", selector)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If this.ElementType <> EElementType.EElement Then
        Err_Raise vbObjectError + 1, "stdHTML#QuerySelectorAll", "Cannot query a non-element context."
      End If
    
      Dim trimmedSelector As String
      trimmedSelector = Trim$(selector)
      If Len(trimmedSelector) = 0 Then
        Err_Raise vbObjectError + 4, "stdHTML#QuerySelectorAll", "Selector cannot be empty."
      End If
    
      Dim groups() As TSelectorGroup
      groups = ParseSelectorGroups(trimmedSelector)
    
      Dim output() As stdHTML
      Dim seen As Collection: Set seen = New Collection
    
      Dim groupUB As Long: groupUB = SafeUBoundSelectorGroupArray(groups)
      Dim i As Long
      For i = 0 To groupUB
        Dim groupMatches() As stdHTML
        groupMatches = EvaluateSelectorGroup(groups(i))
        AppendHTMLArrayUnique output, groupMatches, seen
      Next i
    
      QuerySelectorAll = output
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function EvaluateSelectorGroup(ByRef group As TSelectorGroup) As stdHTML()
  With stdError.getSentry("stdHTML#EvaluateSelectorGroup")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim stepUB As Long: stepUB = SafeUBoundSelectorStepArray(group.Steps)
      If stepUB < 0 Then Exit Function
    
      Dim current() As stdHTML
      Dim currentSeen As Collection: Set currentSeen = New Collection
      Dim firstAtom As TSelectorAtom
      firstAtom = group.Steps(0).Atom
      CollectMatchingWithinScope Me, firstAtom, current, currentSeen
    
      Dim stepIndex As Long
      For stepIndex = 1 To stepUB
        Dim nextMatches() As stdHTML
        Dim nextSeen As Collection: Set nextSeen = New Collection
        Dim currentUB As Long: currentUB = SafeUBoundHTMLArray(current)
        Dim i As Long
        Dim atom As TSelectorAtom
        atom = group.Steps(stepIndex).Atom
        For i = 0 To currentUB
          ApplyCombinator group.Steps(stepIndex).Combinator, current(i), atom, nextMatches, nextSeen
        Next i
        current = nextMatches
      Next stepIndex
    
      EvaluateSelectorGroup = current
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Sub ApplyCombinator(ByVal combinator As String, ByVal baseNode As stdHTML, ByRef atom As TSelectorAtom, ByRef output() As stdHTML, ByVal seen As Collection)
  With stdError.getSentry("stdHTML#ApplyCombinator", "combinator", combinator, "baseNode", baseNode, "seen", seen)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Select Case combinator
        Case " "
          CollectDescendantMatches baseNode, atom, output, seen
        Case ">"
          CollectChildMatches baseNode, atom, output, seen
        Case "+"
          CollectAdjacentSiblingMatch baseNode, atom, output, seen
        Case "~"
          CollectGeneralSiblingMatches baseNode, atom, output, seen
        Case Else
          Err_Raise vbObjectError + 4, "stdHTML#QuerySelectorAll", "Unsupported selector combinator: " & combinator
      End Select
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

Private Sub CollectMatchingWithinScope(ByVal root As stdHTML, ByRef atom As TSelectorAtom, ByRef output() As stdHTML, ByVal seen As Collection)
  With stdError.getSentry("stdHTML#CollectMatchingWithinScope", "root", root, "seen", seen)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If NodeMatchesAtom(root, atom) Then AppendHTMLUnique output, root, seen
      CollectDescendantMatches root, atom, output, seen
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

Private Sub CollectDescendantMatches(ByVal root As stdHTML, ByRef atom As TSelectorAtom, ByRef output() As stdHTML, ByVal seen As Collection)
  With stdError.getSentry("stdHTML#CollectDescendantMatches", "root", root, "seen", seen)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim children As Variant: children = GetChildrenArray(root)
      Dim childUB As Long: childUB = SafeUBoundVariant(children)
      Dim i As Long
      For i = 0 To childUB
        Dim child As stdHTML
        Set child = children(i)
        If Not child Is Nothing Then
          If IsElementNode(child) Then
            If NodeMatchesAtom(child, atom) Then AppendHTMLUnique output, child, seen
            CollectDescendantMatches child, atom, output, seen
          End If
        End If
      Next i
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

Private Sub CollectChildMatches(ByVal root As stdHTML, ByRef atom As TSelectorAtom, ByRef output() As stdHTML, ByVal seen As Collection)
  With stdError.getSentry("stdHTML#CollectChildMatches", "root", root, "seen", seen)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim children As Variant: children = GetChildrenArray(root)
      Dim childUB As Long: childUB = SafeUBoundVariant(children)
      Dim i As Long
      For i = 0 To childUB
        Dim child As stdHTML
        Set child = children(i)
        If Not child Is Nothing Then
          If IsElementNode(child) Then
            If NodeMatchesAtom(child, atom) Then AppendHTMLUnique output, child, seen
          End If
        End If
      Next i
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

Private Sub CollectAdjacentSiblingMatch(ByVal baseNode As stdHTML, ByRef atom As TSelectorAtom, ByRef output() As stdHTML, ByVal seen As Collection)
  With stdError.getSentry("stdHTML#CollectAdjacentSiblingMatch", "baseNode", baseNode, "seen", seen)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim parent As stdHTML
      Dim childIndex As Long
      If Not FindParentAndChildIndex(Me, baseNode, parent, childIndex) Then Exit Sub
    
      Dim siblings As Variant: siblings = GetChildrenArray(parent)
      Dim siblingUB As Long: siblingUB = SafeUBoundVariant(siblings)
      Dim i As Long
      For i = childIndex + 1 To siblingUB
        Dim sibling As stdHTML
        Set sibling = siblings(i)
        If Not sibling Is Nothing Then
          If IsElementNode(sibling) Then
            If NodeMatchesAtom(sibling, atom) Then AppendHTMLUnique output, sibling, seen
            Exit Sub
          End If
        End If
      Next i
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

Private Sub CollectGeneralSiblingMatches(ByVal baseNode As stdHTML, ByRef atom As TSelectorAtom, ByRef output() As stdHTML, ByVal seen As Collection)
  With stdError.getSentry("stdHTML#CollectGeneralSiblingMatches", "baseNode", baseNode, "seen", seen)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim parent As stdHTML
      Dim childIndex As Long
      If Not FindParentAndChildIndex(Me, baseNode, parent, childIndex) Then Exit Sub
    
      Dim siblings As Variant: siblings = GetChildrenArray(parent)
      Dim siblingUB As Long: siblingUB = SafeUBoundVariant(siblings)
      Dim i As Long
      For i = childIndex + 1 To siblingUB
        Dim sibling As stdHTML
        Set sibling = siblings(i)
        If Not sibling Is Nothing Then
          If IsElementNode(sibling) Then
            If NodeMatchesAtom(sibling, atom) Then AppendHTMLUnique output, sibling, seen
          End If
        End If
      Next i
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

Private Function FindParentAndChildIndex(ByVal root As stdHTML, ByVal target As stdHTML, ByRef parentOut As stdHTML, ByRef indexOut As Long) As Boolean
  With stdError.getSentry("stdHTML#FindParentAndChildIndex", "root", root, "target", target, "parentOut", parentOut, "indexOut", indexOut)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim children As Variant: children = GetChildrenArray(root)
      Dim childUB As Long: childUB = SafeUBoundVariant(children)
      Dim i As Long
      For i = 0 To childUB
        Dim child As stdHTML
        Set child = children(i)
        If Not child Is Nothing Then
          If child Is target Then
            Set parentOut = root
            indexOut = i
            FindParentAndChildIndex = True
            Exit Function
          End If
          If IsElementNode(child) Then
            If FindParentAndChildIndex(child, target, parentOut, indexOut) Then
              FindParentAndChildIndex = True
              Exit Function
            End If
          End If
        End If
      Next i
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function NodeMatchesAtom(ByVal node As stdHTML, ByRef atom As TSelectorAtom) As Boolean
  With stdError.getSentry("stdHTML#NodeMatchesAtom", "node", node)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If Not IsElementNode(node) Then Exit Function
    
      If Len(atom.TagName) > 0 Then
        If StrComp(node.TagName, atom.TagName, vbTextCompare) <> 0 Then Exit Function
      ElseIf Not atom.IsUniversal Then
        'No tag selector specified, continue with id/class/attributes.
      End If
    
      If Len(atom.Id) > 0 Then
        Dim nodeId As Variant: nodeId = node.Attr("id")
        If IsEmpty(nodeId) Or IsNull(nodeId) Then Exit Function
        If StrComp(CStr(nodeId), atom.Id, vbTextCompare) <> 0 Then Exit Function
      End If
    
      Dim classUB As Long: classUB = SafeUBoundStringArray(atom.Classes)
      If classUB >= 0 Then
        Dim classList As String
        classList = NormalizeWhitespace(CoalesceString(node.Attr("class")))
        Dim classIndex As Long
        For classIndex = 0 To classUB
          If Not ContainsClass(classList, atom.Classes(classIndex)) Then Exit Function
        Next classIndex
      End If
    
      Dim attrUB As Long: attrUB = SafeUBoundStringArray(atom.AttrNames)
      If attrUB >= 0 Then
        Dim nodeAttributes As Variant
        nodeAttributes = GetNodeAttributes(node)
    
        Dim attrIndex As Long
        For attrIndex = 0 To attrUB
          If atom.AttrHasValue(attrIndex) Then
            If Not ContainsAttributeNameValue(nodeAttributes, atom.AttrNames(attrIndex), atom.AttrValues(attrIndex)) Then Exit Function
          Else
            If Not ContainsAttributeName(nodeAttributes, atom.AttrNames(attrIndex)) Then Exit Function
          End If
        Next attrIndex
      End If
    
      NodeMatchesAtom = True
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function ParseSelectorGroups(ByVal selector As String) As TSelectorGroup()
  With stdError.getSentry("stdHTML#ParseSelectorGroups", "selector", selector)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim idx As Long: idx = 1
      Dim length As Long: length = Len(selector)
      Dim groups() As TSelectorGroup
    
      SkipSelectorWhitespace selector, idx, length
      If idx > length Then
        Err_Raise vbObjectError + 4, "stdHTML#QuerySelectorAll", "Selector cannot be empty."
      End If
    
      Do While idx <= length
        Dim group As TSelectorGroup
        ParseSelectorGroup selector, idx, length, group
        AppendSelectorGroup groups, group
    
        SkipSelectorWhitespace selector, idx, length
        If idx <= length Then
          If Mid$(selector, idx, 1) <> "," Then
            Err_Raise vbObjectError + 4, "stdHTML#QuerySelectorAll", "Unexpected token in selector at position " & CStr(idx) & "."
          End If
          idx = idx + 1
          SkipSelectorWhitespace selector, idx, length
          If idx > length Then
            Err_Raise vbObjectError + 4, "stdHTML#QuerySelectorAll", "Selector cannot end with ','."
          End If
        End If
      Loop
    
      ParseSelectorGroups = groups
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Sub ParseSelectorGroup(ByVal selector As String, ByRef idx As Long, ByVal length As Long, ByRef group As TSelectorGroup)
  With stdError.getSentry("stdHTML#ParseSelectorGroup", "selector", selector, "idx", idx, "length", length)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim step As TSelectorStep
      step.Combinator = ""
      step.Atom = ParseSelectorAtom(selector, idx, length)
      AppendSelectorStep group.Steps, step
    
      Do
        Dim hadWhitespace As Boolean
        hadWhitespace = SkipSelectorWhitespace(selector, idx, length)
        If idx > length Then Exit Do
    
        Dim ch As String: ch = Mid$(selector, idx, 1)
        If ch = "," Then Exit Do
    
        Dim combinator As String
        Select Case ch
          Case ">", "+", "~"
            combinator = ch
            idx = idx + 1
            SkipSelectorWhitespace selector, idx, length
          Case Else
            If hadWhitespace Then
              combinator = " "
            Else
              Err_Raise vbObjectError + 4, "stdHTML#QuerySelectorAll", "Expected combinator at position " & CStr(idx) & "."
            End If
        End Select
    
        If idx > length Then
          Err_Raise vbObjectError + 4, "stdHTML#QuerySelectorAll", "Selector cannot end with combinator."
        End If
        If Mid$(selector, idx, 1) = "," Then
          Err_Raise vbObjectError + 4, "stdHTML#QuerySelectorAll", "Selector cannot end with combinator."
        End If
    
        Dim nextStep As TSelectorStep
        nextStep.Combinator = combinator
        nextStep.Atom = ParseSelectorAtom(selector, idx, length)
        AppendSelectorStep group.Steps, nextStep
      Loop
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

Private Function ParseSelectorAtom(ByVal selector As String, ByRef idx As Long, ByVal length As Long) As TSelectorAtom
  With stdError.getSentry("stdHTML#ParseSelectorAtom", "selector", selector, "idx", idx, "length", length)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim atom As TSelectorAtom
      Dim hasPart As Boolean
    
      If idx > length Then
        Err_Raise vbObjectError + 4, "stdHTML#QuerySelectorAll", "Unexpected end of selector."
      End If
    
      Dim ch As String: ch = Mid$(selector, idx, 1)
      If ch = "*" Then
        atom.IsUniversal = True
        hasPart = True
        idx = idx + 1
      ElseIf IsIdentifierStart(ch) Then
        atom.TagName = ParseIdentifier(selector, idx, length, "tag")
        hasPart = True
      End If
    
      Do While idx <= length
        ch = Mid$(selector, idx, 1)
        Select Case ch
          Case "#"
            idx = idx + 1
            If Len(atom.Id) > 0 Then Err_Raise vbObjectError + 4, "stdHTML#QuerySelectorAll", "Duplicate id selector in one segment."
            atom.Id = ParseIdentifier(selector, idx, length, "id")
            hasPart = True
          Case "."
            idx = idx + 1
            AppendString atom.Classes, ParseIdentifier(selector, idx, length, "class")
            hasPart = True
          Case "["
            idx = idx + 1
            ParseAttributeSelector selector, idx, length, atom
            hasPart = True
          Case Else
            Exit Do
        End Select
      Loop
    
      If Not hasPart Then
        Err_Raise vbObjectError + 4, "stdHTML#QuerySelectorAll", "Expected selector segment at position " & CStr(idx) & "."
      End If
    
      ParseSelectorAtom = atom
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Sub ParseAttributeSelector(ByVal selector As String, ByRef idx As Long, ByVal length As Long, ByRef atom As TSelectorAtom)
  With stdError.getSentry("stdHTML#ParseAttributeSelector", "selector", selector, "idx", idx, "length", length)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      SkipSelectorWhitespace selector, idx, length
      Dim attrName As String
      attrName = ParseIdentifier(selector, idx, length, "attribute name")
    
      SkipSelectorWhitespace selector, idx, length
      Dim hasValue As Boolean
      Dim attrValue As String
      If idx <= length And Mid$(selector, idx, 1) = "=" Then
        hasValue = True
        idx = idx + 1
    
        SkipSelectorWhitespace selector, idx, length
        attrValue = ParseAttributeValue(selector, idx, length)
    
        SkipSelectorWhitespace selector, idx, length
        If idx > length Or Mid$(selector, idx, 1) <> "]" Then
          Err_Raise vbObjectError + 4, "stdHTML#QuerySelectorAll", "Missing closing ']' for attribute selector."
        End If
        idx = idx + 1
      ElseIf idx <= length And Mid$(selector, idx, 1) = "]" Then
        idx = idx + 1
      Else
        Err_Raise vbObjectError + 4, "stdHTML#QuerySelectorAll", "Expected '=' or closing ']' in attribute selector."
      End If
    
      AppendString atom.AttrNames, attrName
      AppendBoolean atom.AttrHasValue, hasValue
      AppendString atom.AttrValues, attrValue
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

Private Function ParseAttributeValue(ByVal selector As String, ByRef idx As Long, ByVal length As Long) As String
  With stdError.getSentry("stdHTML#ParseAttributeValue", "selector", selector, "idx", idx, "length", length)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If idx > length Then
        Err_Raise vbObjectError + 4, "stdHTML#QuerySelectorAll", "Missing attribute value."
      End If
    
      Dim ch As String: ch = Mid$(selector, idx, 1)
      If ch = """" Or ch = "'" Then
        Dim quoteChar As String: quoteChar = ch
        idx = idx + 1
        Dim startPos As Long: startPos = idx
        Do While idx <= length And Mid$(selector, idx, 1) <> quoteChar
          idx = idx + 1
        Loop
        If idx > length Then
          Err_Raise vbObjectError + 4, "stdHTML#QuerySelectorAll", "Unterminated quoted attribute value."
        End If
        ParseAttributeValue = Mid$(selector, startPos, idx - startPos)
        idx = idx + 1
        Exit Function
      End If
    
      Dim valueStart As Long: valueStart = idx
      Do While idx <= length
        ch = Mid$(selector, idx, 1)
        If ch = "]" Or ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then Exit Do
        idx = idx + 1
      Loop
    
      ParseAttributeValue = Trim$(Mid$(selector, valueStart, idx - valueStart))
      If Len(ParseAttributeValue) = 0 Then
        Err_Raise vbObjectError + 4, "stdHTML#QuerySelectorAll", "Missing attribute value."
      End If
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function ParseIdentifier(ByVal selector As String, ByRef idx As Long, ByVal length As Long, ByVal identifierName As String) As String
  With stdError.getSentry("stdHTML#ParseIdentifier", "selector", selector, "idx", idx, "length", length, "identifierName", identifierName)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If idx > length Then
        Err_Raise vbObjectError + 4, "stdHTML#QuerySelectorAll", "Missing " & identifierName & " at end of selector."
      End If
    
      If Not IsIdentifierStart(Mid$(selector, idx, 1)) Then
        Err_Raise vbObjectError + 4, "stdHTML#QuerySelectorAll", "Invalid " & identifierName & " at position " & CStr(idx) & "."
      End If
    
      Dim startPos As Long: startPos = idx
      idx = idx + 1
      Do While idx <= length
        If Not IsIdentifierChar(Mid$(selector, idx, 1)) Then Exit Do
        idx = idx + 1
      Loop
    
      ParseIdentifier = Mid$(selector, startPos, idx - startPos)
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function IsIdentifierStart(ByVal ch As String) As Boolean
  With stdError.getSentry("stdHTML#IsIdentifierStart", "ch", ch)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      IsIdentifierStart = IsIdentifierChar(ch)
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function IsIdentifierChar(ByVal ch As String) As Boolean
  With stdError.getSentry("stdHTML#IsIdentifierChar", "ch", ch)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If Len(ch) = 0 Then Exit Function
      Dim code As Long: code = AscW(ch)
      IsIdentifierChar = _
          (code >= 48 And code <= 57) Or _
          (code >= 65 And code <= 90) Or _
          (code >= 97 And code <= 122) Or _
          ch = "_" Or ch = "-"
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function SkipSelectorWhitespace(ByVal selector As String, ByRef idx As Long, ByVal length As Long) As Boolean
  With stdError.getSentry("stdHTML#SkipSelectorWhitespace", "selector", selector, "idx", idx, "length", length)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Do While idx <= length
        Dim ch As String: ch = Mid$(selector, idx, 1)
        If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then Exit Do
        idx = idx + 1
        SkipSelectorWhitespace = True
      Loop
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function GetChildrenArray(ByVal node As stdHTML) As Variant
  With stdError.getSentry("stdHTML#GetChildrenArray", "node", node)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      On Error GoTo noChildren
      Dim children() As stdHTML
      children = node.Children
      GetChildrenArray = children
      Exit Function
    
    noChildren:
      GetChildrenArray = Array()
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function IsElementNode(ByVal node As stdHTML) As Boolean
  With stdError.getSentry("stdHTML#IsElementNode", "node", node)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If node Is Nothing Then Exit Function
      IsElementNode = (node.ElementType = EElementType.EElement)
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Sub AppendHTMLArrayUnique(ByRef target() As stdHTML, ByRef source() As stdHTML, ByVal seen As Collection)
  With stdError.getSentry("stdHTML#AppendHTMLArrayUnique", "seen", seen)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim sourceUB As Long: sourceUB = SafeUBoundHTMLArray(source)
      Dim i As Long
      For i = 0 To sourceUB
        AppendHTMLUnique target, source(i), seen
      Next i
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

Private Sub AppendHTMLUnique(ByRef values() As stdHTML, ByVal item As stdHTML, ByVal seen As Collection)
  With stdError.getSentry("stdHTML#AppendHTMLUnique", "item", item, "seen", seen)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If item Is Nothing Then Exit Sub
    
      Dim key As String: key = CStr(ObjPtr(item))
      On Error Resume Next
      seen.Add key, key
      If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo stdErrorWrapper_ErrorOccurred
        Exit Sub
      End If
      On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim ub As Long: ub = SafeUBoundHTMLArray(values)
      If ub < 0 Then
        ReDim values(0 To 0)
        Set values(0) = item
      Else
        ReDim Preserve values(0 To ub + 1)
        Set values(ub + 1) = item
      End If
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

Private Sub AppendString(ByRef values() As String, ByVal value As String)
  With stdError.getSentry("stdHTML#AppendString", "value", value)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim ub As Long: ub = SafeUBoundStringArray(values)
      If ub < 0 Then
        ReDim values(0 To 0)
        values(0) = value
      Else
        ReDim Preserve values(0 To ub + 1)
        values(ub + 1) = value
      End If
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

Private Sub AppendBoolean(ByRef values() As Boolean, ByVal value As Boolean)
  With stdError.getSentry("stdHTML#AppendBoolean", "value", value)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim ub As Long: ub = SafeUBoundBooleanArray(values)
      If ub < 0 Then
        ReDim values(0 To 0)
        values(0) = value
      Else
        ReDim Preserve values(0 To ub + 1)
        values(ub + 1) = value
      End If
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

Private Sub AppendSelectorStep(ByRef values() As TSelectorStep, ByRef value As TSelectorStep)
  With stdError.getSentry("stdHTML#AppendSelectorStep")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim ub As Long: ub = SafeUBoundSelectorStepArray(values)
      If ub < 0 Then
        ReDim values(0 To 0)
        values(0) = value
      Else
        ReDim Preserve values(0 To ub + 1)
        values(ub + 1) = value
      End If
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

Private Sub AppendSelectorGroup(ByRef values() As TSelectorGroup, ByRef value As TSelectorGroup)
  With stdError.getSentry("stdHTML#AppendSelectorGroup")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim ub As Long: ub = SafeUBoundSelectorGroupArray(values)
      If ub < 0 Then
        ReDim values(0 To 0)
        values(0) = value
      Else
        ReDim Preserve values(0 To ub + 1)
        values(ub + 1) = value
      End If
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

Private Function DefaultHTMLParseOptions() As THTMLParseOptions
  With stdError.getSentry("stdHTML#DefaultHTMLParseOptions")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim options As THTMLParseOptions
      options.preserveComments = False
      DefaultHTMLParseOptions = options
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function ParseHTMLNodes(ByVal html As String, ByRef options As THTMLParseOptions) As stdHTML()
  With stdError.getSentry("stdHTML#ParseHTMLNodes", "html", html)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim roots() As stdHTML
      Dim stack As New Collection
    
      Dim idx As Long: idx = 1
      Dim length As Long: length = Len(html)
    
      Do While idx <= length
        If Mid$(html, idx, 1) = "<" Then
          If StartsWithAt(html, idx, "<!--") Then
            idx = SkipHTMLComment(html, idx, options)
          ElseIf StartsWithAt(html, idx, "</") Then
            idx = ParseHTMLEndTag(html, idx, stack)
          ElseIf StartsWithAt(html, idx, "<!") Then
            idx = SkipHTMLDeclaration(html, idx)
          ElseIf StartsWithAt(html, idx, "<?") Then
            idx = SkipProcessingInstruction(html, idx)
          Else
            idx = ParseHTMLStartTag(html, idx, stack, roots)
          End If
        Else
          idx = ParseHTMLText(html, idx, stack, roots)
        End If
      Loop
    
      ParseHTMLNodes = roots
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function ParseHTMLText(ByVal html As String, ByVal idx As Long, ByVal stack As Collection, ByRef roots() As stdHTML) As Long
  With stdError.getSentry("stdHTML#ParseHTMLText", "html", html, "idx", idx, "stack", stack)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim nextTag As Long: nextTag = InStr(idx, html, "<")
      Dim content As String
      If nextTag = 0 Then
        content = Mid$(html, idx)
        ParseHTMLText = Len(html) + 1
      Else
        content = Mid$(html, idx, nextTag - idx)
        ParseHTMLText = nextTag
      End If
    
      If Len(content) = 0 Then Exit Function
    
      ' ToString always inserts vbCrLf before each child; pretty-printed source also puts newlines
      ' after '>' and before '</tag>' inside the text node. Strip those edges so we do not double up.
      If stack.Count > 0 Then
        If ShouldTrimParsedEdgeWhitespace(stack(stack.Count).TagName) Then
          content = TrimHTMLAsciiWhitespace(content)
        End If
      End If
      If Len(content) = 0 Then Exit Function
    
      ' Drop text nodes that are only HTML "ASCII whitespace" when they are just markup formatting
      ' between tags (e.g. pretty-printed ToString output). Use Ascii-whitespace-only, not
      ' NormalizeWhitespace, so spaces inside <pre> etc. are not collapsed away.
      If IsHTMLAsciiWhitespaceOnly(content) Then
        If stack.Count = 0 Then Exit Function
        If IsHTMLIgnorableInterelementWhitespaceParent(stack(stack.Count).TagName) Then Exit Function
      End If
    
      Dim textNode As stdHTML
      Set textNode = New stdHTML
      textNode.protInitRaw content
      AppendParsedNode textNode, stack, roots
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function ParseHTMLStartTag(ByVal html As String, ByVal idx As Long, ByVal stack As Collection, ByRef roots() As stdHTML) As Long
  With stdError.getSentry("stdHTML#ParseHTMLStartTag", "html", html, "idx", idx, "stack", stack)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim length As Long: length = Len(html)
      Dim cursor As Long: cursor = idx + 1
      SkipHTMLWhitespace html, cursor, length
    
      Dim tagName As String
      tagName = ReadHTMLName(html, cursor, length)
      If Len(tagName) = 0 Then
        ParseHTMLStartTag = idx + 1
        Exit Function
      End If
    
      Dim attributes As Variant: attributes = Array()
      Dim isSelfClosing As Boolean
    
      Do While cursor <= length
        SkipHTMLWhitespace html, cursor, length
        If cursor > length Then Exit Do
    
        Dim ch As String: ch = Mid$(html, cursor, 1)
        If ch = ">" Then
          cursor = cursor + 1
          Exit Do
        End If
    
        If ch = "/" Then
          Dim lookahead As Long: lookahead = cursor + 1
          SkipHTMLWhitespace html, lookahead, length
          If lookahead <= length And Mid$(html, lookahead, 1) = ">" Then
            isSelfClosing = True
            cursor = lookahead + 1
            Exit Do
          End If
        End If
    
        ParseHTMLAttribute html, cursor, length, attributes
      Loop
    
      If IsHTMLVoidElement(tagName) Then isSelfClosing = True
    
      Dim elementNode As stdHTML
      Set elementNode = New stdHTML
      elementNode.protInitElement tagName, attributes, isSelfClosing
    
      If (Not isSelfClosing) And (StrComp(tagName, "svg", vbTextCompare) = 0) Then
        Dim closingStart As Long: closingStart = InStr(cursor, html, "</svg", vbTextCompare)
        If closingStart > 0 Then
          Dim closingEnd As Long: closingEnd = InStr(closingStart, html, ">")
          If closingEnd = 0 Then closingEnd = length
    
          Dim svgRaw As String: svgRaw = Mid$(html, cursor, closingStart - cursor)
          If Len(svgRaw) > 0 Then
            Dim rawNode As stdHTML
            Set rawNode = New stdHTML
            rawNode.protInitRaw svgRaw
            elementNode.ChildAppend rawNode
          End If
    
          AppendParsedNode elementNode, stack, roots
          ParseHTMLStartTag = closingEnd + 1
          Exit Function
        End If
      End If
    
      AppendParsedNode elementNode, stack, roots
      If Not isSelfClosing Then stack.Add elementNode
    
      ParseHTMLStartTag = cursor
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Sub ParseHTMLAttribute(ByVal html As String, ByRef idx As Long, ByVal length As Long, ByRef attributes As Variant)
  With stdError.getSentry("stdHTML#ParseHTMLAttribute", "html", html, "idx", idx, "length", length, "attributes", attributes)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      SkipHTMLWhitespace html, idx, length
      If idx > length Then Exit Sub
    
      Dim attrName As String
      attrName = ReadHTMLName(html, idx, length)
      If Len(attrName) = 0 Then
        idx = idx + 1
        Exit Sub
      End If
    
      SkipHTMLWhitespace html, idx, length
    
      Dim attrValue As Variant
      If idx <= length And Mid$(html, idx, 1) = "=" Then
        idx = idx + 1
        SkipHTMLWhitespace html, idx, length
    
        If idx <= length Then
          Dim quote As String: quote = Mid$(html, idx, 1)
          If quote = """" Or quote = "'" Then
            idx = idx + 1
            Dim valueStart As Long: valueStart = idx
            Do While idx <= length And Mid$(html, idx, 1) <> quote
              idx = idx + 1
            Loop
            attrValue = Mid$(html, valueStart, idx - valueStart)
            If idx <= length Then idx = idx + 1
          Else
            Dim rawStart As Long: rawStart = idx
            Do While idx <= length
              Dim ch As String: ch = Mid$(html, idx, 1)
              Select Case ch
                Case " ", vbTab, vbCr, vbLf, ">"
                  Exit Do
              End Select
              idx = idx + 1
            Loop
            attrValue = Mid$(html, rawStart, idx - rawStart)
          End If
        Else
          attrValue = vbNullString
        End If
      Else
        ' Name-only attributes are minimized/presence attributes (e.g. disabled, checked).
        attrValue = Null
      End If
    
      AppendVariantPair attributes, attrName, attrValue
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

Private Function ParseHTMLEndTag(ByVal html As String, ByVal idx As Long, ByVal stack As Collection) As Long
  With stdError.getSentry("stdHTML#ParseHTMLEndTag", "html", html, "idx", idx, "stack", stack)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim length As Long: length = Len(html)
      Dim cursor As Long: cursor = idx + 2
    
      SkipHTMLWhitespace html, cursor, length
      Dim tagName As String: tagName = ReadHTMLName(html, cursor, length)
    
      Dim closePos As Long: closePos = InStr(cursor, html, ">")
      If closePos = 0 Then
        ParseHTMLEndTag = length + 1
      Else
        ParseHTMLEndTag = closePos + 1
      End If
    
      If Len(tagName) > 0 Then
        CloseHTMLStackToTag stack, tagName
      End If
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function SkipHTMLComment(ByVal html As String, ByVal idx As Long, ByRef options As THTMLParseOptions) As Long
  With stdError.getSentry("stdHTML#SkipHTMLComment", "html", html, "idx", idx)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim endPos As Long: endPos = InStr(idx + 4, html, "-->")
      If endPos = 0 Then
        SkipHTMLComment = Len(html) + 1
      Else
        SkipHTMLComment = endPos + 3
      End If
    
      ' TODO: Implement optional comment-node retention behind preserveComments option.
      If options.preserveComments Then
        ' Intentionally not implemented in v1.
      End If
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function SkipHTMLDeclaration(ByVal html As String, ByVal idx As Long) As Long
  With stdError.getSentry("stdHTML#SkipHTMLDeclaration", "html", html, "idx", idx)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim endPos As Long: endPos = InStr(idx + 2, html, ">")
      If endPos = 0 Then
        SkipHTMLDeclaration = Len(html) + 1
      Else
        SkipHTMLDeclaration = endPos + 1
      End If
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function SkipProcessingInstruction(ByVal html As String, ByVal idx As Long) As Long
  With stdError.getSentry("stdHTML#SkipProcessingInstruction", "html", html, "idx", idx)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim endPos As Long: endPos = InStr(idx + 2, html, "?>")
      If endPos > 0 Then
        SkipProcessingInstruction = endPos + 2
      Else
        SkipProcessingInstruction = SkipHTMLDeclaration(html, idx)
      End If
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Sub SkipHTMLWhitespace(ByVal html As String, ByRef idx As Long, ByVal length As Long)
  With stdError.getSentry("stdHTML#SkipHTMLWhitespace", "html", html, "idx", idx, "length", length)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Do While idx <= length
        Select Case Mid$(html, idx, 1)
          Case " ", vbTab, vbCr, vbLf
            idx = idx + 1
          Case Else
            Exit Sub
        End Select
      Loop
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

Private Function ReadHTMLName(ByVal html As String, ByRef idx As Long, ByVal length As Long) As String
  With stdError.getSentry("stdHTML#ReadHTMLName", "html", html, "idx", idx, "length", length)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim startPos As Long: startPos = idx
      Do While idx <= length
        If Not IsHTMLNameChar(Mid$(html, idx, 1)) Then Exit Do
        idx = idx + 1
      Loop
      If idx > startPos Then
        ReadHTMLName = Mid$(html, startPos, idx - startPos)
      Else
        ReadHTMLName = vbNullString
      End If
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function IsHTMLNameChar(ByVal ch As String) As Boolean
  With stdError.getSentry("stdHTML#IsHTMLNameChar", "ch", ch)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If Len(ch) = 0 Then Exit Function
    
      Dim code As Long: code = AscW(ch)
      If (code >= 48 And code <= 57) Then
        IsHTMLNameChar = True
        Exit Function
      End If
      If (code >= 65 And code <= 90) Then
        IsHTMLNameChar = True
        Exit Function
      End If
      If (code >= 97 And code <= 122) Then
        IsHTMLNameChar = True
        Exit Function
      End If
    
      Select Case ch
        Case "_", "-", ":", "."
          IsHTMLNameChar = True
      End Select
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function IsHTMLVoidElement(ByVal tagName As String) As Boolean
  With stdError.getSentry("stdHTML#IsHTMLVoidElement", "tagName", tagName)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Select Case LCase$(tagName)
        Case "area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "param", "source", "track", "wbr"
          IsHTMLVoidElement = True
      End Select
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

' Strips leading/trailing HTML ASCII whitespace (TAB, LF, FF, CR, space). Middle is untouched.
Private Function TrimHTMLAsciiWhitespace(ByVal s As String) As String
  With stdError.getSentry("stdHTML#TrimHTMLAsciiWhitespace", "s", s)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim n As Long: n = Len(s)
      Dim L As Long: L = 1
      Dim r As Long: r = n
      Dim c As Long
    
      Do While L <= n
        c = AscW(Mid$(s, L, 1))
        Select Case c
          Case 9, 10, 12, 13, 32
            L = L + 1
          Case Else
            Exit Do
        End Select
      Loop
    
      Do While r >= L
        c = AscW(Mid$(s, r, 1))
        Select Case c
          Case 9, 10, 12, 13, 32
            r = r - 1
          Case Else
            Exit Do
        End Select
      Loop
    
      If L > r Then
        TrimHTMLAsciiWhitespace = vbNullString
      Else
        TrimHTMLAsciiWhitespace = Mid$(s, L, r - L + 1)
      End If
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

' True for elements where leading/trailing text-node whitespace is formatting, not meaningful content.
Private Function ShouldTrimParsedEdgeWhitespace(ByVal tagName As String) As Boolean
  With stdError.getSentry("stdHTML#ShouldTrimParsedEdgeWhitespace", "tagName", tagName)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Select Case LCase$(tagName)
        Case "pre", "textarea", "listing", "xmp", "plaintext", "svg", "math"
          ShouldTrimParsedEdgeWhitespace = False
        Case Else
          ShouldTrimParsedEdgeWhitespace = True
      End Select
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

' True if s is empty or every character is HTML ASCII whitespace (TAB, LF, FF, CR, space).
Private Function IsHTMLAsciiWhitespaceOnly(ByVal s As String) As Boolean
  With stdError.getSentry("stdHTML#IsHTMLAsciiWhitespaceOnly", "s", s)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim i As Long
      Dim n As Long
      Dim c As Long
      n = Len(s)
      For i = 1 To n
        c = AscW(Mid$(s, i, 1))
        Select Case c
          Case 9, 10, 12, 13, 32
          Case Else
            Exit Function
        End Select
      Next i
      IsHTMLAsciiWhitespaceOnly = True
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

' Parents where whitespace-only text between child elements is not meaningful (pretty-print, etc.).
' Default ignorable; only raw / RCDATA / preformatted elements keep ASCII-whitespace-only runs.
Private Function IsHTMLIgnorableInterelementWhitespaceParent(ByVal tagName As String) As Boolean
  With stdError.getSentry("stdHTML#IsHTMLIgnorableInterelementWhitespaceParent", "tagName", tagName)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Select Case LCase$(tagName)
        Case "script", "style", "pre", "textarea", "title", "listing", "xmp", "plaintext"
          IsHTMLIgnorableInterelementWhitespaceParent = False
        Case Else
          IsHTMLIgnorableInterelementWhitespaceParent = True
      End Select
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Sub AppendVariantPair(ByRef values As Variant, ByVal key As String, ByVal value As Variant)
  With stdError.getSentry("stdHTML#AppendVariantPair", "values", values, "key", key, "value", value)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If Not IsArray(values) Then
        ReDim values(0 To 1)
        values(0) = key
        values(1) = value
        Exit Sub
      End If
    
      Dim ub As Long: ub = SafeUBoundVariant(values)
      If ub < 0 Then
        ReDim values(0 To 1)
        values(0) = key
        values(1) = value
      Else
        ReDim Preserve values(0 To ub + 2)
        values(ub + 1) = key
        values(ub + 2) = value
      End If
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

Private Sub AppendParsedNode(ByVal node As stdHTML, ByVal stack As Collection, ByRef roots() As stdHTML)
  With stdError.getSentry("stdHTML#AppendParsedNode", "node", node, "stack", stack)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If node Is Nothing Then Exit Sub
    
      If stack.Count = 0 Then
        AppendHTMLNode roots, node
      Else
        stack(stack.Count).ChildAppend node
      End If
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

Private Sub AppendHTMLNode(ByRef values() As stdHTML, ByVal item As stdHTML)
  With stdError.getSentry("stdHTML#AppendHTMLNode", "item", item)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If item Is Nothing Then Exit Sub
    
      Dim ub As Long: ub = SafeUBoundHTMLArray(values)
      If ub < 0 Then
        ReDim values(0 To 0)
        Set values(0) = item
      Else
        ReDim Preserve values(0 To ub + 1)
        Set values(ub + 1) = item
      End If
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

Private Sub CloseHTMLStackToTag(ByVal stack As Collection, ByVal tagName As String)
  With stdError.getSentry("stdHTML#CloseHTMLStackToTag", "stack", stack, "tagName", tagName)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim i As Long
      For i = stack.Count To 1 Step -1
        If StrComp(stack(i).TagName, tagName, vbTextCompare) = 0 Then
          Dim j As Long
          For j = stack.Count To i Step -1
            stack.Remove j
          Next j
          Exit Sub
        End If
      Next i
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

Private Function StartsWithAt(ByVal text As String, ByVal startPos As Long, ByVal expected As String) As Boolean
  With stdError.getSentry("stdHTML#StartsWithAt", "text", text, "startPos", startPos, "expected", expected)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If startPos < 1 Then Exit Function
      If Len(expected) = 0 Then Exit Function
      StartsWithAt = (StrComp(Mid$(text, startPos, Len(expected)), expected, vbBinaryCompare) = 0)
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function SafeUBoundHTMLArray(ByRef values() As stdHTML) As Long
  With stdError.getSentry("stdHTML#SafeUBoundHTMLArray")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      On Error GoTo noValues
      SafeUBoundHTMLArray = UBound(values)
      Exit Function
    
    noValues:
      SafeUBoundHTMLArray = -1
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function SafeUBoundStringArray(ByRef values() As String) As Long
  With stdError.getSentry("stdHTML#SafeUBoundStringArray")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      On Error GoTo noValues
      SafeUBoundStringArray = UBound(values)
      Exit Function
    
    noValues:
      SafeUBoundStringArray = -1
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function SafeUBoundBooleanArray(ByRef values() As Boolean) As Long
  With stdError.getSentry("stdHTML#SafeUBoundBooleanArray")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      On Error GoTo noValues
      SafeUBoundBooleanArray = UBound(values)
      Exit Function
    
    noValues:
      SafeUBoundBooleanArray = -1
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function SafeUBoundSelectorStepArray(ByRef values() As TSelectorStep) As Long
  With stdError.getSentry("stdHTML#SafeUBoundSelectorStepArray")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      On Error GoTo noValues
      SafeUBoundSelectorStepArray = UBound(values)
      Exit Function
    
    noValues:
      SafeUBoundSelectorStepArray = -1
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function SafeUBoundSelectorGroupArray(ByRef values() As TSelectorGroup) As Long
  With stdError.getSentry("stdHTML#SafeUBoundSelectorGroupArray")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      On Error GoTo noValues
      SafeUBoundSelectorGroupArray = UBound(values)
      Exit Function
    
    noValues:
      SafeUBoundSelectorGroupArray = -1
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function NormalizeAttributes(ByVal attributes As Variant) As Variant
  With stdError.getSentry("stdHTML#NormalizeAttributes", "attributes", attributes)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
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
    
      Err_Raise vbObjectError + 2, "stdHTML", "Attributes must be key/value array pairs."
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function GetNodeAttributes(ByVal node As stdHTML) As Variant
  With stdError.getSentry("stdHTML#GetNodeAttributes", "node", node)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim state As Object
      Set state = node.protGetStateAsDictionary()
      If state.Exists("Attributes") Then
        GetNodeAttributes = state("Attributes")
      Else
        GetNodeAttributes = Array()
      End If
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function ContainsAttributeName(ByVal attributes As Variant, ByVal attrName As String) As Boolean
  With stdError.getSentry("stdHTML#ContainsAttributeName", "attributes", attributes, "attrName", attrName)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim attrUB As Long: attrUB = SafeUBoundVariant(attributes)
      If attrUB < 1 Then Exit Function
    
      Dim i As Long
      For i = 0 To attrUB Step 2
        If StrComp(CStr(attributes(i)), attrName, vbTextCompare) = 0 Then
          ContainsAttributeName = True
          Exit Function
        End If
      Next i
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function ContainsAttributeNameValue(ByVal attributes As Variant, ByVal attrName As String, ByVal expectedValue As String) As Boolean
  With stdError.getSentry("stdHTML#ContainsAttributeNameValue", "attributes", attributes, "attrName", attrName, "expectedValue", expectedValue)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim attrUB As Long: attrUB = SafeUBoundVariant(attributes)
      If attrUB < 1 Then Exit Function
    
      Dim i As Long
      For i = 0 To attrUB Step 2
        If StrComp(CStr(attributes(i)), attrName, vbTextCompare) = 0 Then
          Dim attrValue As Variant
          attrValue = attributes(i + 1)
          If Not IsEmpty(attrValue) And Not IsNull(attrValue) Then
            If StrComp(CStr(attrValue), expectedValue, vbTextCompare) = 0 Then
              ContainsAttributeNameValue = True
              Exit Function
            End If
          End If
        End If
      Next i
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function CoalesceString(ByVal value As Variant) As String
  With stdError.getSentry("stdHTML#CoalesceString", "value", value)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If IsEmpty(value) Or IsNull(value) Then Exit Function
      CoalesceString = CStr(value)
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function ContainsClass(ByVal classList As String, ByVal className As String) As Boolean
  With stdError.getSentry("stdHTML#ContainsClass", "classList", classList, "className", className)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If Len(classList) = 0 Or Len(className) = 0 Then Exit Function
    
      Dim parts() As String: parts = Split(classList, " ")
      Dim i As Long
      For i = LBound(parts) To UBound(parts)
        If StrComp(parts(i), className, vbTextCompare) = 0 Then
          ContainsClass = True
          Exit Function
        End If
      Next i
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function RemoveClass(ByVal classList As String, ByVal className As String) As String
  With stdError.getSentry("stdHTML#RemoveClass", "classList", classList, "className", className)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
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
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function NormalizeWhitespace(ByVal value As String) As String
  With stdError.getSentry("stdHTML#NormalizeWhitespace", "value", value)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      value = Replace(value, vbTab, " ")
      value = Replace(value, vbCr, " ")
      value = Replace(value, vbLf, " ")
      Do While InStr(value, "  ") > 0
        value = Replace(value, "  ", " ")
      Loop
      NormalizeWhitespace = Trim$(value)
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function EscapeHtmlText(ByVal value As String) As String
  With stdError.getSentry("stdHTML#EscapeHtmlText", "value", value)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      value = Replace(value, "&", "&amp;")
      value = Replace(value, "<", "&lt;")
      value = Replace(value, ">", "&gt;")
      EscapeHtmlText = value
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function EscapeHtmlAttribute(ByVal value As String) As String
  With stdError.getSentry("stdHTML#EscapeHtmlAttribute", "value", value)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      value = EscapeHtmlText(value)
      value = Replace(value, """", "&quot;")
      value = Replace(value, "'", "&#39;")
      EscapeHtmlAttribute = value
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function SafeUBoundVariant(ByVal values As Variant) As Long
  With stdError.getSentry("stdHTML#SafeUBoundVariant", "values", values)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      On Error GoTo noValues
      SafeUBoundVariant = UBound(values)
      Exit Function
    
    noValues:
      SafeUBoundVariant = -1
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function UBoundAttributes() As Long
  With stdError.getSentry("stdHTML#UBoundAttributes")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      On Error GoTo noAttributes
      UBoundAttributes = UBound(This.Element.Attributes)
      Exit Function
    
    noAttributes:
      UBoundAttributes = -1
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function UBoundChildren() As Long
  With stdError.getSentry("stdHTML#UBoundChildren")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      On Error GoTo noChildren
      UBoundChildren = UBound(This.Element.Children)
      Exit Function
    
    noChildren:
      UBoundChildren = -1
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

Private Function ObjFromPtr(ByVal ptr As LongPtr) As stdHTML
  With stdError.getSentry("stdHTML#ObjFromPtr", "ptr", ptr)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim tmp As stdHTML
      CopyMemory tmp, ptr, PTR_SZ
      Set ObjFromPtr = tmp
      CopyMemory tmp, NULL_PTR, PTR_SZ
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function



Private Sub Err_Raise(ByVal number as Long, Optional ByVal source as string = "", Optional ByVal description as string = "")
  Call stdError.Raise(description)
End Sub
