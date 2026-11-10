VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdPowerQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'@module
'@description Execute Power Query M and stream results through a positioned VBA record reader.
'@remark WIP: tested with 64-bit Excel runtime 2.161.8230.1; external connectors are not configured.
'@remark Contains the CLR bootstrap and native reflection calls; no helper DLL, compiler, PowerShell process, or worksheet load is used.
'@remark The Mashup API is internal; see docs/stdPowerQuery.md for tested Office versions.
'@requires Windows VBA7, .NET Framework 4, and an installed Excel Mashup runtime matching host bitness.

#If Win64 Then
Private Const PTR_SIZE As Long = 8
Private Const VT_PTR As Integer = 20
#Else
Private Const PTR_SIZE As Long = 4
Private Const VT_PTR As Integer = 3
#End If
Private Const GET_INSTANCE As Long = &H1414 'Public | Instance | GetField | GetProperty
Private Const GET_STATIC As Long = &H1418
Private Const CALL_INSTANCE As Long = &H114
Private Const CALL_STATIC As Long = &H118
Private Const RT As String = "Microsoft.Mashup.Engine1.Runtime."
Private Const IID_TYPE As String = "{BCA8B44D-AAD6-3A86-8AB7-03349F4F2DA2}"
Private Const IID_DOMAIN As String = "{05F696DC-2B29-3663-AD8B-C4389CF2A713}"

Private Declare PtrSafe Function CLRCreateInstance Lib "mscoree" (ByRef clsid As Long, ByRef iid As Long, ByRef obj As LongPtr) As Long
Private Declare PtrSafe Function CLSIDFromString Lib "ole32" (ByVal text As LongPtr, ByRef clsid As Long) As Long
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" (ByVal instance As LongPtr, ByVal offset As LongPtr, ByVal convention As Long, ByVal resultType As Integer, ByVal argc As Long, ByRef types As Integer, ByRef args As LongPtr, ByRef result As Variant) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef src As Any, ByVal size As LongPtr)
Private Declare PtrSafe Function VariantCopy Lib "oleaut32" (ByRef dest As Variant, ByRef src As Variant) As Long
Private Declare PtrSafe Function VariantCopyInd Lib "oleaut32" (ByRef dest As Variant, ByRef src As Variant) As Long
Private Declare PtrSafe Function VariantClear Lib "oleaut32" (ByRef value As Variant) As Long
Private Declare PtrSafe Function GetErrorInfo Lib "oleaut32" (ByVal reserved As Long, ByRef info As LongPtr) As Long
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal codePage As Long, ByVal flags As Long, ByVal bytes As LongPtr, ByVal byteCount As Long, ByVal text As LongPtr, ByVal charCount As Long) As Long

Private Type TThis
  Initialized As Boolean
  Assembly As Object
  Core As Object
  Types As Object
  DeclaredTypes As Object
  FieldTypes As Object
  Path As String
  MaxDepth As Long
  Root As Variant
  Iterator As Variant
  Current As Variant
  Columns As Collection
  Kind As String
  Count As Double
  IsReader As Boolean
  Closed As Boolean
  Completed As Boolean
  HasCurrent As Boolean
End Type
Private This As TThis

'Create a query positioned on its first record or EOF.
'@constructor
'@param mashup - M source expression.
'@param EnginePath - Optional engine DLL path; empty enables discovery.
'@returns - Query with its first record already available.
Public Function Create(ByVal mashup As String, Optional ByVal EnginePath As String = "") As stdPowerQuery
  If Len(Trim$(mashup)) = 0 Then Err.Raise 5, "stdPowerQuery.Create", "The M expression is empty."
  Set Create = New stdPowerQuery
  Call Create.protInit(mashup, EnginePath)
End Function

'Create a positioned query from a UTF-8 M expression file.
'@constructor
'@param mFile - UTF-8 file path; an optional UTF-8 BOM is accepted.
'@param EnginePath - Optional engine DLL path.
'@returns - Query with its first record already available.
Public Function CreateFromFile(ByVal mFile As String, Optional ByVal EnginePath As String = "") As stdPowerQuery
  Set CreateFromFile = Create(ReadUtf8(mFile), EnginePath)
End Function

'Return the path of the loaded Microsoft Mashup engine assembly.
'@returns - Absolute runtime DLL path.
Public Property Get RuntimePath() As String
  RuntimePath = This.Path
End Property

'Get or set the maximum nesting depth used when converting M values.
'@remark Allowed values are 1 through 64; the default is 32.
Public Property Get MaxDepth() As Long
  MaxDepth = This.MaxDepth
End Property

'Get or set the maximum nesting depth used when converting M values.
'@remark Allowed values are 1 through 64; the default is 32.
Public Property Let MaxDepth(ByVal value As Long)
  If value < 1 Or value > 64 Then Err.Raise 5, "stdPowerQuery.MaxDepth", "Use a depth from 1 to 64."
  This.MaxDepth = value
End Property

'Return whether there is no current record.
'@returns - True for empty, exhausted, closed, or failed queries.
Public Property Get EOF() As Boolean
  EOF = Not This.HasCurrent
End Property

'Default member: query!Column is equivalent to query.FieldValue("Column").
'Read a current field through its name or the query!Column default syntax.
'@defaultMember
'@param Name - Field name; exact matches take precedence over a unique case-insensitive match.
'@returns Variant<Dictionary|Collection|Scalar> - Converted field value.
Public Property Get FieldValue(ByVal Name As String) As Variant
Attribute FieldValue.VB_UserMemId = 0
  RequireCurrent
  Name = ResolveFieldName(Name, This.Current)
  If IsObject(This.Current(Name)) Then
    Set FieldValue = This.Current(Name)
  Else
    FieldValue = This.Current(Name)
  End If
End Property

'Declared non-Any types win; otherwise infer from the current M value.
'At EOF an untyped column reports Any. Missing fields always raise an error.
'Return the declared column type or infer the current M value type.
'@param Name - Field name.
'@returns - Type name without .Type; an untyped field reports Any at EOF.
Public Property Get FieldType(ByVal Name As String) As String
  If This.DeclaredTypes Is Nothing Then Err.Raise 5, "stdPowerQuery.FieldType", "Create a query first."
  Name = ResolveFieldName(Name, This.DeclaredTypes)
  FieldType = This.DeclaredTypes(Name)
  If FieldType = "Any" And This.HasCurrent Then FieldType = This.FieldTypes(Name)
End Property

'Resolve an exact field name or a unique case-insensitive fallback.
'@param Name - Requested field name.
'@param fields - Dictionary containing available field names.
'@returns - Actual field name; missing or ambiguous names raise an error.
Private Function ResolveFieldName(ByVal Name As String, ByVal fields As Object) As String
  Dim candidate As Variant, found As Boolean
  If fields.Exists(Name) Then
    ResolveFieldName = Name
    Exit Function
  End If
  'The VBE can recase bang identifiers. Fall back only to a unique match.
  For Each candidate In fields.Keys
    If StrComp(CStr(candidate), Name, vbTextCompare) = 0 Then
      If found Then Err.Raise 5, "stdPowerQuery", "Ambiguous field name: " & Name
      ResolveFieldName = CStr(candidate)
      found = True
    End If
  Next
  If Not found Then Err.Raise 5, "stdPowerQuery", "Unknown field: " & Name
End Function

'MoveNext pulls one item and materializes its children into VBA values.
'The previous Current remains valid if the caller keeps a reference to it.
'Advance the lazy result iterator and materialize one item.
'@returns - True when another current item is available; False at exhaustion.
'@remark Evaluation errors close the iterator and raise a VBA error.
Public Function MoveNext() As Boolean
  If Not This.IsReader Then Err.Raise 5, "stdPowerQuery.MoveNext", "Create a query first."
  If This.Closed Then Exit Function
  Dim value As Variant, reference As Variant, number As Long, description As String, phase As String
  On Error GoTo Failed
  This.HasCurrent = False
  VariantClear This.Current
  Set This.FieldTypes = Nothing
  'Opaque CLR interfaces can be VT_UNKNOWN, for which VBA IsObject is False.
  If Not IsEmpty(This.Iterator) Then
    phase = "advancing the iterator"
    If Not Reflect("System.Collections.IEnumerator", "MoveNext", CALL_INSTANCE, This.Iterator, Array(), True) Then
      This.Completed = True
      CloseStream
      Exit Function
    End If
    phase = "reading the current reference"
    VariantCopy reference, Reflect("System.Collections.IEnumerator", "Current", GET_INSTANCE, This.Iterator, Array(), True)
    phase = "reading the referenced value"
    VariantCopy value, Reflect(RT & "IValueReference", "Value", GET_INSTANCE, reference, Array())
  Else
    If This.Count > 0 Then
      This.Completed = True
      CloseStream
      Exit Function
    End If
    VariantCopy value, This.Root
  End If
  phase = "converting the current value"
  ReadRecord value
  This.HasCurrent = True
  This.Count = This.Count + 1
  MoveNext = True
  Exit Function
Failed:
  number = Err.Number: description = Err.Description
  CloseStream
  Err.Raise number, "stdPowerQuery.MoveNext", phase & ": " & description
End Function

'Return the current record as a mutable dictionary snapshot.
'@returns Variant<Dictionary> - Record whose lifetime is independent of subsequent iteration.
Public Property Get Current() As Variant
  RequireCurrent
  Set Current = This.Current
End Property

'Return the root M value kind.
'@returns - Lowercase kind name such as table, list, record, or number.
Public Property Get Kind() As String
  Kind = This.Kind
End Property

'Return a copy of the ordered column names.
'@returns Collection<String> - One-based column names, including for an empty table.
Public Property Get Columns() As Collection
  Dim result As New Collection, name As Variant
  If Not This.Columns Is Nothing Then
    For Each name In This.Columns
      result.Add name
    Next
  End If
  Set Columns = result
End Property

'Return the number of records fetched, including the prefetched current record.
'@returns - Fetched record count, not the total result size.
Public Property Get Count() As Double
  Count = This.Count
End Property

'Return whether iteration reached normal exhaustion.
'@returns - False after early close or evaluation failure.
Public Property Get Completed() As Boolean
  Completed = This.Completed
End Property

'Return whether the result iterator has been closed.
'@returns - True after exhaustion, early close, or evaluation failure.
Public Property Get IsClosed() As Boolean
  IsClosed = This.Closed
End Property

'CloseStream is idempotent. Early close does not claim successful completion.
'Release the iterator and current value.
'@remark Safe to call repeatedly. Early close does not report successful completion.
Public Sub CloseStream()
  On Error Resume Next
  If Not IsEmpty(This.Iterator) Then Reflect "System.IDisposable", "Dispose", CALL_INSTANCE, This.Iterator, Array(), True
  VariantClear This.Iterator
  VariantClear This.Root
  VariantClear This.Current
  Set This.FieldTypes = Nothing
  This.HasCurrent = False
  This.Closed = True
End Sub

'Use when binding text parameters or constructing quoted identifiers (# + MText).
'Quote VBA text as an M string literal.
'@param value - Text to quote, including quotes, control characters, and Unicode.
'@returns - An escaped M string expression.
Public Function MText(ByVal value As String) As String
  value = Replace(value, "#", "#(#)")
  value = Replace(value, """", """""")
  value = Replace(value, vbCr, "#(cr)")
  value = Replace(value, vbLf, "#(lf)")
  value = Replace(value, vbTab, "#(tab)")
  value = Replace(value, vbNullChar, "#(0000)")
  MText = """" & value & """"
End Function

'Reject field access when the query has no current record.
Private Sub RequireCurrent()
  If Not This.HasCurrent Then Err.Raise 5, "stdPowerQuery", "No current record (EOF)."
End Sub

'Initialize a positioned query. Internal constructor hook.
'@protected
'@param mashup - M source expression.
'@param EnginePath - Optional engine DLL path.
'@remark Called by Create; an instance may only be initialized once.
Public Sub protInit(ByVal mashup As String, Optional ByVal EnginePath As String = "")
  If This.Initialized Then Err.Raise 5, "stdPowerQuery.protInit", "Instance already initialized."
  If Len(Trim$(mashup)) = 0 Then Err.Raise 5, "stdPowerQuery.protInit", "The M expression is empty."
  This.Initialized = True
  InitializeRuntime EnginePath
  InitializeQuery mashup
End Sub

'Evaluate M source, read its schema, and fetch the first record.
'@param expression - M source expression.
Private Sub InitializeQuery(ByVal expression As String)
  This.IsReader = True
  Dim engine As Variant, minimal As Variant, library As Variant
  Dim number As Long, description As String
  On Error GoTo Failed
  VariantCopy engine, Reflect("Microsoft.Mashup.Engine1.Engine", "Instance", GET_STATIC, Empty, Array())
  VariantCopy minimal, Reflect("Microsoft.Mashup.Evaluator.MinimalEngineHost", "Instance", GET_STATIC, Empty, Array())
  VariantCopy library, Reflect("Microsoft.Mashup.Engine.Interface.IEngine", "GetLibrary", CALL_INSTANCE, engine, Array(minimal, Nothing))
  VariantCopy This.Root, Reflect("Microsoft.Mashup.Engine1.Language.LanguageLibrary", "Evaluate", CALL_STATIC, Empty, Array(expression, library))
  This.Kind = ValueKind(This.Root)
  Set This.Columns = New Collection
  If This.Kind = "table" Then
    Set This.Columns = ValueNames(This.Root, True)
  ElseIf This.Kind = "record" Then
    Set This.Columns = ValueNames(This.Root, False)
  Else
    This.Columns.Add "Value"
  End If
  ReadDeclaredTypes
  If This.Kind = "table" Or This.Kind = "list" Then VariantCopy This.Iterator, OpenIterator(This.Root, This.Kind)
  MoveNext
  Exit Sub
Failed:
  number = Err.Number: description = Err.Description
  CloseStream
  Err.Raise number, "stdPowerQuery.Create", description
End Sub

'Cache top-level column type names without enumerating table rows.
Private Sub ReadDeclaredTypes()
  Dim i As Long, name As String, typeValue As Variant, typeName As String
  Set This.DeclaredTypes = CreateObject("Scripting.Dictionary")
  This.DeclaredTypes.CompareMode = vbBinaryCompare
  For i = 1 To This.Columns.Count
    name = This.Columns(i)
    typeName = "Any"
    If This.Kind = "table" Then
      VariantCopy typeValue, Reflect(RT & "TableValue", "GetColumnType", CALL_INSTANCE, This.Root, Array(i - 1))
      VariantCopy typeValue, Reflect(RT & "TypeValue", "NonNullable", GET_INSTANCE, typeValue, Array())
      typeName = Reflect(RT & "TypeValue", "GetTypeName", CALL_STATIC, Empty, Array(typeValue))
      If Right$(typeName, 5) = ".Type" Then typeName = Left$(typeName, Len(typeName) - 5)
    End If
    This.DeclaredTypes.Add name, typeName
  Next
End Sub

'Materialize a record and cache its original M field kinds.
'@param value as Variant<IUnknown> - Current runtime value.
Private Sub ReadRecord(ByVal value As Variant)
  Dim name As Variant, child As Variant, row As Object
  Set This.FieldTypes = CreateObject("Scripting.Dictionary")
  This.FieldTypes.CompareMode = vbBinaryCompare
  If This.Kind = "table" Or This.Kind = "record" Then
    VariantCopy This.Current, ConvertValue(value, 0)
    For Each name In This.Columns
      VariantCopy child, Reflect(RT & "RecordValue", "Item", GET_INSTANCE, value, Array(CStr(name)))
      This.FieldTypes.Add CStr(name), InferredType(child)
    Next
  Else
    Set row = CreateObject("Scripting.Dictionary")
    row.CompareMode = vbBinaryCompare
    row.Add "Value", ConvertValue(value, 0)
    Set This.Current = row
    This.FieldTypes.Add "Value", InferredType(value)
  End If
End Sub

'Infer the record-style type name of an M value.
'@param value as Variant<IUnknown> - Runtime value.
'@returns - Integral numbers report Int64; other values retain their M kind.
Private Function InferredType(ByVal value As Variant) As String
  Dim kind As String, number As Variant
  kind = ValueKind(value)
  Select Case kind
    Case "number"
      VariantCopy number, Reflect(RT & "Value", "AsNumber", GET_INSTANCE, value, Array())
      If Reflect(RT & "NumberValue", "IsInteger64", GET_INSTANCE, number, Array()) Then
        InferredType = "Int64"
      Else
        Select Case VarType(ConvertValue(value, 0))
          Case vbDecimal: InferredType = "Decimal"
          Case Else: InferredType = "Double"
        End Select
      End If
    Case "datetime": InferredType = "DateTime"
    Case "datetimezone": InferredType = "DateTimeZone"
    Case Else: InferredType = UCase$(Left$(kind, 1)) & Mid$(kind, 2)
  End Select
End Function

'A .m file contains one UTF-8 M expression; a UTF-8 BOM is optional.
'Read and validate a UTF-8 M source file.
'@param path - Existing source file path.
'@returns - Decoded M source without an optional leading BOM.
Private Function ReadUtf8(ByVal path As String) As String
  Dim file As Integer, opened As Boolean, length As Long, chars As Long
  Dim bytes() As Byte, text As String, number As Long, description As String
  On Error GoTo Failed
  If Not CreateObject("Scripting.FileSystemObject").FileExists(path) Then Err.Raise 53, , "M file not found: " & path
  file = FreeFile
  Open path For Binary Access Read As #file
  opened = True
  length = LOF(file)
  If length > 0 Then
    ReDim bytes(0 To length - 1)
    Get #file, , bytes
  End If
  Close #file
  opened = False
  If length = 0 Then Exit Function
  chars = MultiByteToWideChar(65001, &H8, VarPtr(bytes(0)), length, 0, 0)
  If chars = 0 Then Err.Raise 5, , "M file is not valid UTF-8: " & path
  text = String$(chars, vbNullChar)
  If MultiByteToWideChar(65001, &H8, VarPtr(bytes(0)), length, StrPtr(text), chars) <> chars Then Err.Raise 5, , "UTF-8 conversion failed: " & path
  If Left$(text, 1) = ChrW$(&HFEFF) Then text = Mid$(text, 2)
  ReadUtf8 = text
  Exit Function
Failed:
  number = Err.Number: description = Err.Description
  On Error Resume Next
  If opened Then Close #file
  On Error GoTo 0
  Err.Raise number, "stdPowerQuery.CreateFromFile", description
End Function

'Resolve an opaque runtime value to its M kind name.
'@param value as Variant<IUnknown> - Mashup runtime value.
'@returns - Lowercase M kind name.
Private Function ValueKind(ByVal value As Variant) As String
  Dim kind As Variant, typ As IUnknown
  VariantCopy kind, Reflect(RT & "Value", "Kind", GET_INSTANCE, value, Array())
  Set typ = This.Assembly.GetType_2(RT & "ValueKind")
  ValueKind = LCase$(Reflect("System.Enum", "GetName", CALL_STATIC, Empty, Array(typ, kind), True))
End Function

'Read ordered table column names or record field names.
'@param value as Variant<IUnknown> - Table or record runtime value.
'@param isTable - True for table columns; False for record fields.
'@returns Collection<String> - One-based ordered names.
Private Function ValueNames(ByVal value As Variant, ByVal isTable As Boolean) As Collection
  Dim keys As Variant, result As New Collection, i As Long, length As Long
  If isTable Then
    VariantCopy keys, Reflect(RT & "TableValue", "Columns", GET_INSTANCE, value, Array())
  Else
    VariantCopy keys, Reflect(RT & "RecordValue", "Keys", GET_INSTANCE, value, Array())
  End If
  length = Reflect(RT & "Keys", "Length", GET_INSTANCE, keys, Array())
  For i = 0 To length - 1
    result.Add Reflect(RT & "Keys", "Item", GET_INSTANCE, keys, Array(i))
  Next
  Set ValueNames = result
End Function

'Open a lazy runtime iterator over list items or table records.
'@param value as Variant<IUnknown> - Table or list runtime value.
'@param kind - Root M kind.
'@returns Variant<IUnknown> - Owned runtime enumerator; the caller must dispose it.
Private Function OpenIterator(ByVal value As Variant, ByVal kind As String) As Variant
  Dim list As Variant
  If kind = "table" Then
    VariantCopy list, Reflect(RT & "TableValue", "ToRecords", CALL_INSTANCE, value, Array())
  Else
    VariantCopy list, value
  End If
  VariantCopy OpenIterator, Reflect(RT & "ListValue", "GetEnumerator", CALL_INSTANCE, list, Array())
End Function

'Convert a runtime value and its children to native VBA data.
'@param value as Variant<IUnknown> - Mashup runtime value.
'@param depth - Current nesting depth.
'@returns Variant<Dictionary|Collection|Scalar> - Materialized VBA value.
Private Function ConvertValue(ByVal value As Variant, ByVal depth As Long) As Variant
  If depth > This.MaxDepth Then Err.Raise 5, "stdPowerQuery", "Result exceeds MaxDepth."
  Dim kind As String, scalar As Variant, child As Variant, iterator As Variant, reference As Variant
  Dim record As Object, items As Collection, name As Variant
  Dim number As Long, description As String
  On Error GoTo Failed
  kind = ValueKind(value)
  Select Case kind
    Case "null": ConvertValue = Null
    Case "logical": ConvertValue = Reflect(RT & "Value", "AsBoolean", GET_INSTANCE, value, Array())
    Case "text": ConvertValue = Reflect(RT & "Value", "AsString", GET_INSTANCE, value, Array())
    Case "number"
      VariantCopy scalar, Reflect(RT & "Value", "AsNumber", GET_INSTANCE, value, Array())
      VariantCopy ConvertValue, Reflect(RT & "NumberValue", "ToObject", CALL_INSTANCE, scalar, Array())
    Case "binary"
      VariantCopy ConvertValue, Reflect(RT & "BinaryValue", "AsBytes", GET_INSTANCE, value, Array())
    Case "date"
      VariantCopy ConvertValue, Reflect(RT & "DateValue", "AsClrDateTime", GET_INSTANCE, value, Array())
    Case "datetime"
      VariantCopy ConvertValue, Reflect(RT & "DateTimeValue", "AsClrDateTime", GET_INSTANCE, value, Array())
    Case "datetimezone"
      VariantCopy scalar, Reflect("System.Globalization.CultureInfo", "InvariantCulture", GET_STATIC, Empty, Array(), True)
      ConvertValue = Reflect(RT & "DateTimeZoneValue", "ToString", CALL_INSTANCE, value, Array("o", scalar))
    Case "time"
      VariantCopy scalar, Reflect("System.Globalization.CultureInfo", "InvariantCulture", GET_STATIC, Empty, Array(), True)
      ConvertValue = Reflect(RT & "TimeValue", "ToString", CALL_INSTANCE, value, Array("HH:mm:ss.FFFFFFF", scalar))
    Case "duration"
      ConvertValue = Reflect(RT & "DurationValue", "ToString", CALL_INSTANCE, value, Array())
    Case "record"
      Set record = CreateObject("Scripting.Dictionary")
      record.CompareMode = vbBinaryCompare
      For Each name In ValueNames(value, False)
        VariantCopy child, Reflect(RT & "RecordValue", "Item", GET_INSTANCE, value, Array(CStr(name)))
        record.Add CStr(name), ConvertValue(child, depth + 1)
      Next
      Set ConvertValue = record
    Case "list", "table"
      Set items = New Collection
      VariantCopy iterator, OpenIterator(value, kind)
      Do While Reflect("System.Collections.IEnumerator", "MoveNext", CALL_INSTANCE, iterator, Array(), True)
        VariantCopy reference, Reflect("System.Collections.IEnumerator", "Current", GET_INSTANCE, iterator, Array(), True)
        VariantCopy child, Reflect(RT & "IValueReference", "Value", GET_INSTANCE, reference, Array())
        items.Add ConvertValue(child, depth + 1)
      Loop
      Reflect "System.IDisposable", "Dispose", CALL_INSTANCE, iterator, Array(), True
      VariantClear iterator
      Set ConvertValue = items
    Case Else
      Err.Raise 5, "stdPowerQuery", "Cannot export an M " & kind & " value. Convert it to data in M first."
  End Select
  Exit Function
Failed:
  number = Err.Number: description = Err.Description
  On Error Resume Next
  If Not IsEmpty(iterator) Then Reflect "System.IDisposable", "Dispose", CALL_INSTANCE, iterator, Array(), True
  On Error GoTo 0
  Err.Raise number, "stdPowerQuery", "Converting " & kind & ": " & description
End Function

'---------------- CLR bootstrap and COM ABI boundary ----------------
'The four vtable contracts below are the Windows/.NET Framework interfaces,
'not Mashup's private layout. Mashup members are resolved by name.
'Load the CLR and installed Mashup assembly for this instance.
'@param enginePath - Optional engine DLL path.
Private Sub InitializeRuntime(ByVal enginePath As String)
  Dim g1(0 To 3) As Long, g2(0 To 3) As Long
  Dim meta As LongPtr, info As LongPtr, host As LongPtr, unknownDomain As LongPtr
  Dim domain As LongPtr, core As LongPtr, unk As IUnknown, typ As IUnknown
  Dim version As String, hr As Long, number As Long, description As String
  On Error GoTo Failed
  This.Path = FindRuntime(enginePath)
  This.MaxDepth = 32
  Set This.Types = CreateObject("Scripting.Dictionary")
  Guid "{9280188D-0E8E-4867-B30C-7FA83884E8DE}", g1
  Guid "{D332DB9E-B9B3-4125-8207-A14884F53216}", g2
  CheckHR CLRCreateInstance(g1(0), g2(0), meta), "CLRCreateInstance"
  Guid "{BD39D1D2-BA2F-486A-89B0-B4B0CB466891}", g2
  version = "v4.0.30319"
  CheckHR NativeCall(meta, 3, Array(VT_PTR, VT_PTR, VT_PTR), CLngPtr(StrPtr(version)), CLngPtr(VarPtr(g2(0))), CLngPtr(VarPtr(info))), "ICLRMetaHost.GetRuntime"
  Guid "{CB2F6723-AB3A-11D2-9C40-00C04FA30A3E}", g1
  Guid "{CB2F6722-AB3A-11D2-9C40-00C04FA30A3E}", g2
  CheckHR NativeCall(info, 9, Array(VT_PTR, VT_PTR, VT_PTR), CLngPtr(VarPtr(g1(0))), CLngPtr(VarPtr(g2(0))), CLngPtr(VarPtr(host))), "ICLRRuntimeInfo.GetInterface"
  CheckHR NativeCall(host, 10, Array()), "ICorRuntimeHost.Start"
  CheckHR NativeCall(host, 13, Array(VT_PTR), CLngPtr(VarPtr(unknownDomain))), "ICorRuntimeHost.GetDefaultDomain"
  domain = QueryNativeInterface(unknownDomain, IID_DOMAIN)
  CheckHR NativeCall(domain, 44, Array(vbString, VT_PTR), "mscorlib", CLngPtr(VarPtr(core))), "_AppDomain.Load_2"
  CopyMemory unk, core, PTR_SIZE 'Transfer the returned COM reference to VBA.
  core = 0
  Set This.Core = unk
  Set typ = This.Core.GetType_2("System.Reflection.Assembly")
  Set This.Assembly = InvokeType(typ, "LoadFrom", CALL_STATIC, Empty, Array(This.Path))
  If StrComp(This.Assembly.Location, This.Path, vbTextCompare) <> 0 Then
    Err.Raise 5, "stdPowerQuery", "Another Mashup runtime is already loaded. Restart Excel to switch runtimes."
  End If
CleanUp:
  ReleasePtr core
  ReleasePtr domain
  ReleasePtr unknownDomain
  ReleasePtr host
  ReleasePtr info
  ReleasePtr meta
  If number <> 0 Then Err.Raise number, "stdPowerQuery.Create", description
  Exit Sub
Failed:
  number = Err.Number: description = Err.Description
  Resume CleanUp
End Sub

'Find the installed Mashup runtime or validate a supplied path.
'@param requested - Explicit engine path, or an empty string for discovery.
'@returns - Absolute engine DLL path.
Private Function FindRuntime(ByVal requested As String) As String
  Dim fso As Object, root As Variant, candidate As String
  Set fso = CreateObject("Scripting.FileSystemObject")
  If Len(requested) = 0 Then requested = Environ$("MASHUP_ENGINE_PATH")
  If Len(requested) > 0 Then
    If Not fso.FileExists(requested) Then Err.Raise 53, "stdPowerQuery", "Mashup runtime not found: " & requested
    FindRuntime = fso.GetAbsolutePathName(requested)
    Exit Function
  End If
  For Each root In Array(Environ$("ProgramW6432"), Environ$("ProgramFiles"), Environ$("ProgramFiles(x86)"))
    candidate = root & "\Microsoft Office\root\Office16\ADDINS\Microsoft Power Query for Excel Integrated\bin\Microsoft.MashupEngine.dll"
    If fso.FileExists(candidate) Then
      FindRuntime = candidate
      Exit Function
    End If
  Next
  Err.Raise 53, "stdPowerQuery", "Excel Mashup runtime not found. Pass its Microsoft.MashupEngine.dll path to Create."
End Function

'Look up a runtime type and invoke one of its members.
'@param typeName - Fully qualified CLR type name.
'@param member - Method, field, or property name.
'@param flags - CLR reflection binding flags.
'@param target as Variant<IUnknown|Empty> - Instance or Empty for a static member.
'@param args as Variant<Array<Variant>> - Member arguments.
'@param useCore - Resolve the type from mscorlib instead of the Mashup assembly.
'@returns Variant - Reflection result; opaque CLR values retain their COM references.
Private Function Reflect(ByVal typeName As String, ByVal member As String, ByVal flags As Long, ByVal target As Variant, ByVal args As Variant, Optional ByVal useCore As Boolean = False) As Variant
  Dim typ As IUnknown, key As String
  key = CStr(useCore) & ":" & typeName
  If This.Types.Exists(key) Then
    Set typ = This.Types(key)
  Else
    If useCore Then
      Set typ = This.Core.GetType_2(typeName)
    Else
      Set typ = This.Assembly.GetType_2(typeName)
    End If
    If typ Is Nothing Then Err.Raise 5, "stdPowerQuery", "Runtime type unavailable: " & typeName
    This.Types.Add key, typ
  End If
  VariantCopy Reflect, InvokeType(typ, member, flags, target, args)
End Function

'Invoke a member through the CLR reflection COM interface.
'@param typ - CLR type interface.
'@param member - Member name.
'@param flags - Reflection binding flags.
'@param target as Variant<IUnknown|Empty> - Member target.
'@param args as Variant<Array<Variant>> - Member arguments.
'@returns Variant - Marshaled member result.
Private Function InvokeType(ByVal typ As IUnknown, ByVal member As String, ByVal flags As Long, ByVal target As Variant, ByVal args As Variant) As Variant
  Dim ptr As LongPtr, arrayPtr As LongPtr, hr As Long, copiedArgs As Variant, copiedTarget As Variant
  Dim description As String, number As Long
  On Error GoTo Failed
  VariantCopyInd copiedArgs, args
  VariantCopyInd copiedTarget, target
  ptr = QueryNativeInterface(ObjPtr(typ), IID_TYPE)
  CopyMemory arrayPtr, ByVal VarPtr(copiedArgs) + 8, PTR_SIZE
  'On x64 a by-value VARIANT is passed by address under the Windows ABI.
#If Win64 Then
  hr = NativeCall(ptr, 57, Array(vbString, vbLong, VT_PTR, VT_PTR, VT_PTR, VT_PTR), member, flags, CLngPtr(0), CLngPtr(VarPtr(copiedTarget)), arrayPtr, CLngPtr(VarPtr(InvokeType)))
#Else
  hr = NativeCall(ptr, 57, Array(vbString, vbLong, VT_PTR, vbVariant, VT_PTR, VT_PTR), member, flags, CLngPtr(0), copiedTarget, arrayPtr, CLngPtr(VarPtr(InvokeType)))
#End If
  If hr < 0 Then description = ErrorDescription()
  ReleasePtr ptr
  If hr < 0 Then Err.Raise hr, "stdPowerQuery." & member, description
  Exit Function
Failed:
  number = Err.Number: description = Err.Description
  On Error Resume Next
  ReleasePtr ptr
  On Error GoTo 0
  Err.Raise number, "stdPowerQuery." & member, description
End Function

'Call a COM vtable slot using the Windows ABI.
'@param ptr - Borrowed interface pointer.
'@param slot - Zero-based vtable slot.
'@param types as Variant<Array<Integer>> - Native argument VARTYPEs.
'@param args as Variant<Array<Variant>> - Native arguments.
'@returns - Native HRESULT or reference count.
Private Function NativeCall(ByVal ptr As LongPtr, ByVal slot As Long, ByVal types As Variant, ParamArray args() As Variant) As Long
  Dim vt(0 To 15) As Integer, pointers(0 To 15) As LongPtr, values(0 To 15) As Variant
  Dim i As Long, hr As Long, result As Variant
  If ptr = 0 Then Err.Raise 5, "stdPowerQuery.NativeCall", "Null interface pointer."
  If UBound(args) > 15 Then Err.Raise 5, "stdPowerQuery.NativeCall", "Too many native arguments."
  For i = 0 To UBound(args)
    'ParamArray can contain VT_BYREF values. Normalize before DispCallFunc.
    VariantCopyInd values(i), args(i)
    vt(i) = types(i)
    pointers(i) = VarPtr(values(i))
  Next
  hr = DispCallFunc(ptr, slot * PTR_SIZE, 4, vbLong, UBound(args) + 1, vt(0), pointers(0), result)
  CheckHR hr, "DispCallFunc"
  NativeCall = result
End Function

'Acquire an owned COM interface pointer.
'@param ptr - Borrowed source interface pointer.
'@param iid - Requested interface GUID.
'@returns - Owned interface pointer that must be released.
Private Function QueryNativeInterface(ByVal ptr As LongPtr, ByVal iid As String) As LongPtr
  Dim id(0 To 3) As Long
  Guid iid, id
  CheckHR NativeCall(ptr, 0, Array(VT_PTR, VT_PTR), CLngPtr(VarPtr(id(0))), CLngPtr(VarPtr(QueryNativeInterface))), "QueryNativeInterface"
End Function

'Parse a GUID into four Long values.
'@param text - GUID string.
'@param result - Four-element output array.
Private Sub Guid(ByVal text As String, ByRef result() As Long)
  CheckHR CLSIDFromString(StrPtr(text), result(0)), "CLSIDFromString"
End Sub

'Release an owned COM reference and clear its pointer.
'@param ptr - Owned interface pointer, or zero.
Private Sub ReleasePtr(ByRef ptr As LongPtr)
  If ptr = 0 Then Exit Sub
  NativeCall ptr, 2, Array()
  ptr = 0
End Sub

'Raise a VBA error when a native call fails.
'@param hr - Native HRESULT.
'@param operation - Operation name for the error source.
Private Sub CheckHR(ByVal hr As Long, ByVal operation As String)
  If hr < 0 Then Err.Raise hr, "stdPowerQuery." & operation, ErrorDescription()
End Sub

'Read CLR error information and unwrap reflection exceptions.
'@returns - Underlying exception message when available.
Private Function ErrorDescription() As String
  Dim ptr As LongPtr, unk As IUnknown, exception As Object, inner As Object, i As Long
  On Error Resume Next
  GetErrorInfo 0, ptr
  If ptr = 0 Then
    ErrorDescription = "The CLR call failed."
    Exit Function
  End If
  NativeCall ptr, 5, Array(VT_PTR), CLngPtr(VarPtr(ErrorDescription))
  'CLR IErrorInfo also exposes _Exception through IDispatch. Unwrap reflection.
  CopyMemory unk, ptr, PTR_SIZE
  ptr = 0
  Set exception = unk
  For i = 1 To 16
    Set inner = Nothing
    Set inner = exception.InnerException
    If inner Is Nothing Then Exit For
    Set exception = inner
  Next
  If Not exception Is Nothing Then ErrorDescription = exception.Message
End Function

'Release any remaining iterator when VBA destroys the instance.
Private Sub Class_Terminate()
  CloseStream
  'Never Stop the CLR or unload Excel's default AppDomain.
End Sub
