# stdPowerQuery

Run Power Query M directly from VBA and iterate results as records. Import
**[src/stdPowerQuery.cls](../src/stdPowerQuery.cls)**; the class contains its CLR
bootstrap, runtime access, record reader, and value conversion. It requires no
PowerShell process, helper DLL, runtime compilation, registration, or additional
VBA reference. Results stay in VBA objects; no workbook queries, connections,
worksheets, or data-model tables are created.

**WIP:** this hosts an internal Microsoft Mashup API. The tested baseline is
64-bit Excel 16.0 with Excel's installed Mashup runtime **2.161.8230.1**. Office
updates can change that API. The 32-bit declarations are present but have not
been exercised in 32-bit Office. Mac VBA and VBA6 are unsupported.

## Requirements and source boundary

- Windows, VBA7, .NET Framework 4, and an installed Excel Mashup runtime matching
  the host process's bitness. The class itself has no Excel object-model calls;
  other Office hosts have not yet been tested.
- Import the exported class file, preserving its predeclared instance and
  default-member attributes. Pasting the body into a blank class loses these.
- Pure M expressions and explicitly supplied inline data are supported.
  `MinimalEngineHost` does **not** grant external source permissions or provide
  credentials: `File.Contents`, `Web.Contents`, SQL and other external connectors
  are not configured. Excel's saved credentials, queries and workbook context
  are not imported. Reading an M source file with `CreateFromFile` does not
  authorize a data connector.

Microsoft's installed engine parses and evaluates M; the VBA class is a host,
not an M interpreter. Only run M you trust; this is not a security sandbox.

## Create a query and process records

```vb
Sub CalculateAmounts()
  Dim query As stdPowerQuery, row As Object, total As Double
  Dim number As Long, description As String
  On Error GoTo Failed
  Set query = stdPowerQuery.Create( _
    "#table({""Qty"",""Price""},{{2,10.5},{3,4}})")
  While Not query.EOF
    Debug.Print query!Qty
    Debug.Print query.FieldValue("Price")
    Debug.Print query.FieldType("Qty") 'Int64 for these untyped integral values
    Set row = query.Current
    row("Amount") = query!Qty * query!Price
    total = total + row("Amount")
    'Pass row to your CSV writer, database command, or other application code.
    query.MoveNext
  Wend
  Debug.Print total '33
  Exit Sub
Failed:
  number = Err.Number: description = Err.Description
  If Not query Is Nothing Then query.CloseStream
  Err.Raise number, "CalculateAmounts", description
End Sub
```

`Create(mashup, [EnginePath])` evaluates once and positions on the first record.
An empty result starts at `EOF = True`. Process the current record **before**
calling `MoveNext`; the factory has already fetched it. `MoveNext` returns True
when another record is available and False after normal exhaustion.

`CreateFromFile(mFile, [EnginePath])` reads a UTF-8 M **expression**, with or
without a UTF-8 BOM. Invalid UTF-8, missing files and empty expressions raise VBA
errors. A file containing an exported `let ... in ...` expression is suitable;
automatic section/DataMashup/workbook import is not implemented.

```vb
Set query = stdPowerQuery.CreateFromFile("C:\queries\people.m")
Set query = stdPowerQuery.Create("{1,2,3}", "C:\Office\bin\Microsoft.MashupEngine.dll")
```

By default, the class searches standard Office installation folders and honors
`MASHUP_ENGINE_PATH`. `RuntimePath` reports the assembly it loaded. Restart the
host before selecting a different Mashup runtime; the class never stops the
shared CLR or unloads the default AppDomain.

## Fields, schema, and result shape

`query!Column`, `query("Column")`, and `query.FieldValue("Column")` use the same
default getter. Lookups prefer exact names, then a unique case-insensitive match
to accommodate VBA editor recasing. For M fields differing only by case, use
exact string names such as `FieldValue("a")` and `FieldValue("A")`. Unknown or
ambiguous fields raise an error without changing the current record. Reading a
field or `Current` at EOF raises an error.

`FieldType(name)` uses a declared non-Any column type when available, dropping
`.Type` and nullability: `Int64`, `Text`, `Number`, etc. Otherwise it infers the
current M value's kind. Integral values fitting Int64 report `Int64`; other
numeric values report `Double` or `Decimal`. Other kinds include `Null`,
`Logical`, `Date`, `DateTime`, `DateTimeZone`, `Time`, `Duration`, `Binary`,
`Record`, `List` and `Table`. Inferred types may vary by row and describe the
original M value, not edits to the VBA dictionary. At EOF, declared types remain
available and an untyped column reports `Any`. Empty typed tables retain types.

`Columns` returns a copy of the ordered names, including for an empty table.
`Kind` identifies the root M kind. Tables expose their fields directly; a root
record produces one record. Scalars and lists use one column called `Value`,
with one record per scalar/list item. A record inside a list remains a dictionary
inside `FieldValue("Value")`.

`Current` is a mutable dictionary snapshot. Keeping that reference preserves it
after advancing or closing; changing it does not change M's values or schema.

## Streaming, lifetime, and errors

The class pulls a lazy runtime enumerator and materializes one record and its
nested values at a time. Earlier records are not accumulated by the host.
M operations such as sorts, joins and grouping may buffer internally. Individual
nested lists/tables and binary fields are materialized within their parent row.

`Count` counts fetched records, **including the prefetched current record**;
it is not the total result size. `Completed` becomes True only after normal
exhaustion. `IsClosed` also becomes True after failure or early close.
`CloseStream` is idempotent; early close sets EOF without claiming completion.
Multiple queries have independent iterators.

A first-record error raises from the factory. A later error raises from
`MoveNext`, releases the iterator, and leaves `Completed = False`. Work already
performed by a CSV/database sink may be partial; use staging or caller-owned
transactions when all-or-nothing output is required. Calls are synchronous;
closing between pulls cannot interrupt a currently blocking M operation.
VBA `End`/VBE Reset bypass normal cleanup; close live queries before resetting.

## Native VBA values and supplied data

- Records become `Scripting.Dictionary` objects with binary key comparison.
- Lists and nested table rows become one-based `Collection` objects; nested
  table rows are dictionaries. Empty/singleton collections preserve their shape.
- Text, logical and null become String, Boolean and Null; binary becomes Byte().
- Numbers retain the runtime's numeric Variant representation, including Decimal
  values used for large integers. Converting them to Double yourself can lose
  precision. VBA operations/sinks may not support non-finite floating values.
- Dates/datetimes become VBA Date, losing their M-kind distinction and potentially
  precision at the COM DATE boundary. Datetimezones become round-trip strings
  with offsets; time and duration become strings. These M kinds remain visible
  through `FieldType`.

`MaxDepth` defaults to 32 and accepts 1 through 64. The factory fetches its first
record using the default; changing MaxDepth afterward affects subsequent rows.
Function/type/action results are rejected. M metadata, table keys and full type
annotations are not exported into the VBA containers; nested empty tables lose
their column names.

Use `MText` to quote supplied VBA strings as M source, including Unicode, quotes
and `#(...)` escapes. For JSON already obtained by the caller's VBA code:

```vb
Set query = stdPowerQuery.Create( _
  "Table.FromRecords(Json.Document(" & stdPowerQuery.MText(jsonText) & "))")
```

Dependencies/libraries can be included in the expression's `let` bindings.
`Excel.CurrentWorkbook()` requires a caller-supplied fixture/definition; there is
no automatic active workbook context.

## Validation and implementation

Run the opt-in suite with `python Tools/Testing/run.py --suite stdPowerQuery`.
The registered cases cover default access, schema, values, precision, snapshots,
UTF-8 files, empty/failed/closed states, independent queries, an infinite-list
prefix, 1,100 streamed records and no workbook output. The shared runner imports
only this class and test support into disposable Excel projects, watches owned
dialogs, and restores exact Office/VBE settings.

Factories call the documented public `protInit` hook, with state in `TThis`/`This`.
Initialization is guarded against repeated calls. Native `DispCallFunc` calls
the Windows/.NET COM interfaces; reflection resolves Mashup members by name.
No executable-memory thunk is required. The class follows stdVBA's standalone
native COM approach and executes entirely from VBA.

Platform contracts: [CLRCreateInstance](https://learn.microsoft.com/en-us/dotnet/framework/unmanaged-api/hosting/clrcreateinstance-function),
[GetDefaultDomain](https://learn.microsoft.com/en-us/dotnet/framework/unmanaged-api/hosting/icorruntimehost-getdefaultdomain-method),
and [DispCallFunc](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-dispcallfunc).
