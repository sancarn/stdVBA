Attribute VB_Name = "stdOutlookTests"
Option Explicit
#Const LIVE_TEST = False

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As LongPtr)
Private Declare PtrSafe Function MAPIAllocateBuffer Lib "mapi32.dll" (ByVal size As Long, ByRef pointer As LongPtr) As Long
Private Declare PtrSafe Function MAPIAllocateMore Lib "mapi32.dll" (ByVal size As Long, ByVal parent As LongPtr, ByRef pointer As LongPtr) As Long
Private Declare PtrSafe Function MAPIInitialize Lib "mapi32.dll" (ByVal init As LongPtr) As Long
Private Declare PtrSafe Sub MAPIUninitialize Lib "mapi32.dll" ()

#If Win64 Then
  Private Const PB As Long = 8
  Private Const SP As Long = 24
  Private Const SR As Long = 16
  Private Const RO As Long = 8
  Private Const RU As Long = 8
#Else
  Private Const PB As Long = 4
  Private Const SP As Long = 16
  Private Const SR As Long = 12
  Private Const RO As Long = 4
  Private Const RU As Long = 4
#End If

Private passed As Long
Private nextRow As Long
Private totalRows As Long
Private queryCalls As Long
Private releaseCalls As Long
Private restrictionCalls As Long
Private failQuery As Boolean
Private ansiRestriction As Boolean
Private callbackFailure As String
Private personMode As Long
Private personName As String
Private personSMTP As String
Private personLegacyDN As String
Private proxyAlias As String
Private proxyPrimary As String
Private proxyPointers(0 To 1) As LongPtr

Public Function RunCase() As String
  Dim client As stdOutlook, liveStatus As String, initialized As Boolean
  On Error GoTo Failed
  Set client = New stdOutlook
  passed = 0
  Test.Refresh
  Trace "lifecycle"
  TestLifecycle client
  Trace "properties"
  TestProperties client
  Trace "tables"
  Assert MAPIInitialize(0) >= 0, "MAPI initialized for native allocator tests"
  initialized = True
  TestTables client
  Trace "identity"
  TestIdentity client
  Trace "send and calendar"
  passed = passed + stdOutlookSendCalendarTests.RunTests()
  MAPIUninitialize
  initialized = False
  liveStatus = "not requested"
#If LIVE_TEST Then
  liveStatus = LiveSmoke()
#End If
  If Test.failures <> 0 Or Test.checks = 0 Then Err.Raise 5, "stdOutlookTests", "Missing or failed assertions: " & Test.lastCheck
  RunCase = "{""status"":""passed"",""assertions"":" & passed & ",""live"":""" & EscapeJSON(liveStatus) & """}"
  Exit Function
Failed:
  RunCase = "{""status"":""failed"",""assertions"":" & passed & ",""error"":""" & EscapeJSON(Err.Source & ": " & Err.description & " (" & Err.Number & ")") & """}"
  If initialized Then MAPIUninitialize
End Function

Private Sub Assert(ByVal condition As Boolean, ByVal label As String)
  Test.Assert label, condition
  If Not condition Then Err.Raise vbObjectError + 2001, "stdOutlookTests", label
  passed = passed + 1
  Trace "PASS " & label
End Sub

Private Sub Trace(ByVal label As String)
  Dim file As Integer, path As String
  path = Environ$("STDOUTLOOK_TEST_PROGRESS")
  If Len(path) = 0 Then Exit Sub
  file = FreeFile
  Open path For Append As #file
  Print #file, label
  Close #file
End Sub

Private Sub TestLifecycle(ByVal client As stdOutlook)
  Dim unused As Collection, number As Long
  Assert Not client.IsOpen, "New client must be closed"
  client.Disconnect
  client.Disconnect
  Assert Not client.IsOpen, "Disconnect is idempotent"
  On Error Resume Next
  Set unused = client.CurrentUser()
  number = Err.Number
  On Error GoTo 0
  Assert number = 91, "Closed client read fails cleanly"
  On Error Resume Next
  Set unused = client.SearchContacts(" ")
  number = Err.Number
  On Error GoTo 0
  Assert number = 5, "Empty search rejected"
  On Error Resume Next
  Set unused = client.SearchContacts("name", 0)
  number = Err.Number
  On Error GoTo 0
  Assert number = 5, "Zero limit rejected"
  Assert client.TestEntryID("00aB10FF") = "00AB10FF", "Binary entry ID round trip"
  On Error Resume Next
  client.TestEntryID "1G"
  number = Err.Number
  On Error GoTo 0
  Assert number = 5, "Nonhex ID rejected"
  On Error Resume Next
  client.TestEntryID "123"
  number = Err.Number
  On Error GoTo 0
  Assert number = 5, "Odd ID rejected"
End Sub

Private Sub TestProperties(ByVal client As stdOutlook)
  Dim props(0 To SP * 3 - 1) As Byte, pointer As LongPtr, text As String
  Dim ansi() As Byte, value As Variant, record As Collection, errors As Collection
  Dim binary(0 To 3) As Byte, small As Integer, fileTime As Currency
  Dim pointers(0 To 1) As LongPtr, other As String
  pointer = VarPtr(props(0))
  text = "Zo" & ChrW$(&HEB) & " " & ChrW$(&H6771) & ChrW$(&H4EAC)
  PutLong pointer, &H3001001F
  PutPtr pointer + 8, StrPtr(text)
  Assert client.TestDecode(pointer) = text, "Unicode including non-Latin characters"
  ansi = StrConv("ANSI text" & vbNullChar, vbFromUnicode)
  PutLong pointer, &H3001001E
  PutPtr pointer + 8, VarPtr(ansi(0))
  Assert client.TestDecode(pointer) = "ANSI text", "ANSI string conversion"
  PutLong pointer, &H3001001F
  PutPtr pointer + 8, 0
  Assert client.TestDecode(pointer) = "", "Null string pointer"
  PutLong pointer, &H30000003
  PutLong pointer + 8, -123456
  Assert client.TestDecode(pointer) = -123456, "Signed native Long"
  PutLong pointer, &HE1B000B
  small = 1
  CopyMemory ByVal pointer + 8, small, 2
  Assert client.TestDecode(pointer) = True, "Native boolean width"
  binary(0) = 0: binary(1) = 1: binary(2) = 128: binary(3) = 255
  PutLong pointer, &HFFF0102
  PutLong pointer + 8, 4
  PutPtr pointer + 8 + PB, VarPtr(binary(0))
  Assert client.TestDecode(pointer) = "000180FF", "SBinary native pointer alignment"
  PutLong pointer, &HE060040
  fileTime = 13348540800000@ '2024-01-01 UTC, raw FILETIME ticks / 10000
  CopyMemory ByVal pointer + 8, fileTime, 8
  Assert client.TestDecode(pointer) = DateSerial(2024, 1, 1), "FILETIME converted as UTC"
  other = "SMTP:person@example.test"
  pointers(0) = StrPtr(text): pointers(1) = StrPtr(other)
  PutLong pointer, &H800F101F
  PutLong pointer + 8, 2
  PutPtr pointer + 8 + PB, VarPtr(pointers(0))
  value = client.TestDecode(pointer)
  Assert value(0) = text And value(1) = other, "Multivalue string alignment"
  PutLong pointer + 8, 0
  value = client.TestDecode(pointer)
  Assert IsArray(value), "Empty multivalue property"
  PutLong pointer, &H3001001F
  PutPtr pointer + 8, StrPtr(text)
  PutLong pointer + SP, &H39FE000A
  PutLong pointer + SP + 8, &H8004010F
  PutLong pointer + SP * 2, &HFFE0003
  PutLong pointer + SP * 2 + 8, 6
  Set record = client.TestProperties(pointer, 3, Array(Array("Name", &H3001001F), Array("SMTP", &H39FE001F), Array("Kind", &HFFE0003)))
  Set errors = record("PropertyErrors")
  Assert record("Name") = text And record("Kind") = 6, "SPropValue stride"
  Assert IsEmpty(record("SMTP")), "Missing property is Empty"
  Assert errors("SMTP") = &H8004010F, "Partial property error preserved"
End Sub

Private Sub TestTables(ByVal client As stdOutlook)
  Dim vt(0 To 19) As LongPtr, instance As LongPtr, pointer As LongPtr
  Dim rows As Collection, capped As Boolean, row As Collection, number As Long
  vt(2) = Address(AddressOf FakeRelease)
  vt(7) = Address(AddressOf FakeColumns)
  vt(14) = Address(AddressOf FakeRestrict)
  vt(19) = Address(AddressOf FakeQueryRows)
  instance = VarPtr(vt(0)): pointer = VarPtr(instance)
  nextRow = 0: totalRows = 5: queryCalls = 0: failQuery = False: callbackFailure = ""
  Set rows = client.TestTable(pointer, 10, capped)
  Assert callbackFailure = "", "Native callback: " & callbackFailure
  Assert rows.Count = 5 And Not capped, "Pagination continues after short pages"
  Assert queryCalls = 4, "QueryRows called until zero rows"
  Set row = rows(5)
  Assert row("Number") = 4 And row("Text") = "row-4", "Rows copied before native buffers freed"
  nextRow = 0: queryCalls = 0
  Set rows = client.TestTable(pointer, 3, capped)
  Assert rows.Count = 3 And capped, "Result cap detects an extra row"
  nextRow = 0: totalRows = 3
  Set rows = client.TestTable(pointer, 3, capped)
  Assert rows.Count = 3 And Not capped, "Exactly full table is not truncated"
  nextRow = 0: totalRows = 0
  Set rows = client.TestTable(pointer, 3, capped)
  Assert rows.Count = 0 And Not capped, "Empty table"
  failQuery = True
  On Error Resume Next
  Set rows = client.TestTable(pointer, 3, capped)
  number = Err.Number
  On Error GoTo 0
  Assert number = &H80004005, "Native HRESULT propagated"
  failQuery = False
  restrictionCalls = 0: ansiRestriction = False
  client.TestRestriction pointer, "Ada"
  Assert restrictionCalls = 1 And callbackFailure = "", "Unicode ANR native structure"
  restrictionCalls = 0: ansiRestriction = True
  client.TestRestriction pointer, "Ada"
  Assert restrictionCalls = 2 And callbackFailure = "", "ANSI restriction fallback"
  releaseCalls = 0
  client.TestRelease pointer
  Assert pointer = 0 And releaseCalls = 1, "Release once, zero owned pointer"
  client.TestRelease pointer
  Assert releaseCalls = 1, "Repeated Release is a no-op"
End Sub

Private Sub TestIdentity(ByVal client As stdOutlook)
  Dim vt(0 To 5) As LongPtr, instance As LongPtr, person As Collection
  vt(5) = Address(AddressOf FakeGetProps)
  instance = VarPtr(vt(0))
  personName = "Ada Example": personSMTP = "ada@example.test"
  personLegacyDN = "/o=Example/ou=People/cn=Ada"
  proxyAlias = "smtp:alias@example.test": proxyPrimary = "SMTP:primary@example.test"
  proxyPointers(0) = StrPtr(proxyAlias): proxyPointers(1) = StrPtr(proxyPrimary)
  personMode = 0
  Set person = client.TestPerson(VarPtr(instance))
  Assert person("Email") = personSMTP, "PR_SMTP_ADDRESS preferred"
  personMode = 1
  Set person = client.TestPerson(VarPtr(instance))
  Assert person("Email") = "primary@example.test", "Uppercase primary proxy selected over alias"
  personMode = 2
  Set person = client.TestPerson(VarPtr(instance))
  Assert person("Email") = "", "Exchange legacy DN never returned as SMTP"
  Assert person("EmailAddress") = personLegacyDN, "Raw Exchange address preserved"
  Assert callbackFailure = "", "GetProps ABI and unspecified string tags"
End Sub

'Native callbacks must not unwind VBA errors across the unmanaged boundary.
Public Function FakeColumns(ByVal self As LongPtr, ByVal tags As LongPtr, ByVal flags As Long) As Long
  Trace "FakeColumns tags=" & GetLong(tags) & "; flags=" & flags
  If GetLong(tags) <> 2 Or GetLong(tags + 4) <> &H30000003 Or GetLong(tags + 8) <> &H3001001F Or flags <> 0 Then
    callbackFailure = "SetColumns argument layout"
    FakeColumns = &H80070057
  End If
End Function

Public Function FakeQueryRows(ByVal self As LongPtr, ByVal requested As Long, ByVal flags As Long, ByRef rows As LongPtr) As Long
  Dim count As Long, i As Long, props As LongPtr, row As LongPtr, text As String, buffer As LongPtr
  On Error GoTo Failed
  queryCalls = queryCalls + 1
  Trace "FakeQueryRows call=" & queryCalls & "; requested=" & requested & "; next=" & nextRow
  If queryCalls > 10 Then FakeQueryRows = &H80004005: Exit Function
  If failQuery Then FakeQueryRows = &H80004005: Exit Function
  count = totalRows - nextRow
  If count > 2 Then count = 2
  If count > requested Then count = requested
  If MAPIAllocateBuffer(RO + count * SR, rows) <> 0 Then GoTo Failed
  PutLong rows, count
  For i = 0 To count - 1
    If MAPIAllocateBuffer(SP * 2, props) <> 0 Then GoTo Failed
    row = rows + RO + i * SR
    PutLong row, 0: PutLong row + 4, 2: PutPtr row + 8, props
    PutLong props, &H30000003: PutLong props + 8, nextRow
    text = "row-" & CStr(nextRow)
    If MAPIAllocateMore((Len(text) + 1) * 2, props, buffer) <> 0 Then GoTo Failed
    CopyMemory ByVal buffer, ByVal StrPtr(text), (Len(text) + 1) * 2
    PutLong props + SP, &H3001001F: PutPtr props + SP + 8, buffer
    nextRow = nextRow + 1
  Next
  Exit Function
Failed:
  callbackFailure = "QueryRows allocation/callback failed"
  FakeQueryRows = &H80004005
End Function

Public Function FakeRestrict(ByVal self As LongPtr, ByVal restriction As LongPtr, ByVal flags As Long) As Long
  Dim prop As LongPtr, expectedTag As Long
  restrictionCalls = restrictionCalls + 1
  expectedTag = &H360C001F
  If ansiRestriction And restrictionCalls = 2 Then expectedTag = &H360C001E
  prop = GetPtr(restriction + RU + 8)
  If GetLong(restriction) <> 4 Or GetLong(restriction + RU) <> 4 Or GetLong(restriction + RU + 4) <> expectedTag Or GetLong(prop) <> expectedTag Or flags <> 0 Then
    callbackFailure = "Restriction layout"
    FakeRestrict = &H80070057
    Exit Function
  End If
  If ansiRestriction And restrictionCalls = 1 Then FakeRestrict = &H80040103
End Function

Public Function FakeRelease(ByVal self As LongPtr) As Long
  releaseCalls = releaseCalls + 1
  FakeRelease = 0
End Function

Public Function FakeGetProps(ByVal self As LongPtr, ByVal tags As LongPtr, ByVal flags As Long, ByRef count As Long, ByRef props As LongPtr) As Long
  Dim i As Long, tag As Long, prop As LongPtr, addrType As String
  On Error GoTo Failed
  count = GetLong(tags)
  If count <> 16 Or flags <> &H80000000 Then callbackFailure = "GetProps arguments"
  If MAPIAllocateBuffer(count * SP, props) <> 0 Then GoTo Failed
  For i = 0 To count - 1
    tag = GetLong(tags + 4 + i * 4)
    prop = props + i * SP
    PutLong prop, (tag And &HFFFF0000) Or 10
    PutLong prop + 8, &H8004010F
    Select Case tag And &HFFFF0000
      Case &H30010000
        If (tag And &HFFFF&) <> 0 Then callbackFailure = "Expected PT_UNSPECIFIED for strings"
        PutLong prop, &H3001001F: PutPtr prop + 8, StrPtr(personName)
      Case &H30030000
        PutLong prop, &H3003001F: PutPtr prop + 8, StrPtr(personLegacyDN)
      Case &H39FE0000
        If personMode = 0 Then PutLong prop, &H39FE001F: PutPtr prop + 8, StrPtr(personSMTP)
      Case &H800F0000
        If personMode = 1 Then
          PutLong prop, &H800F101F: PutLong prop + 8, 2: PutPtr prop + 8 + PB, VarPtr(proxyPointers(0))
        End If
    End Select
  Next
  FakeGetProps = &H40380 'MAPI_W_ERRORS_RETURNED is success with per-property errors
  Exit Function
Failed:
  callbackFailure = "GetProps callback failed"
  FakeGetProps = &H80004005
End Function

Private Function LiveSmoke() As String
  Dim client As stdOutlook, other As stdOutlook, person As Collection
  Dim books As Collection, matches As Collection, messages As Collection, message As Collection
  Dim calendar As Collection, appointment As Collection, calendarMatches As Collection
  Dim identityError As Long, id As String, query As String, matchCount As Long
  On Error GoTo Unavailable
  Set client = stdOutlook.Create()
  Set other = stdOutlook.Create()
  Assert client.IsOpen And other.IsOpen, "Two native sessions"
  other.Disconnect
  Assert Not other.IsOpen, "Native session Close actually releases state"
  Set books = client.AddressBooks()
  On Error Resume Next
  Set person = client.CurrentUser()
  identityError = Err.Number
  On Error GoTo Unavailable
  If identityError = 0 Then
    Assert Len(client.CurrentUserEmail) > 0, "Live identity supplies an SMTP address"
    Assert client.CurrentUserEmail = person("Email"), "CurrentUserEmail agrees with identity record"
    query = CStr(person("DisplayName"))
    If Len(query) > 0 Then
      Set matches = client.SearchContacts(query, 3)
      Assert client.SearchWarnings.Count = 0, "Live contact search reported provider failures"
      matchCount = matches.Count
      Assert matchCount > 0, "Live identity is found by name search"
    End If
  End If
  Set messages = client.Inbox(1)
  If messages.Count > 0 Then
    Set message = messages(1)
    Set message = client.ReadMessage(CStr(message("EntryID")))
  End If
  Set calendar = client.SearchCalendar("", 1)
  Assert calendar.Count <= 1, "Live calendar limit honored"
  If calendar.Count > 0 Then
    Set appointment = calendar(1)
    Assert appointment("RecurrenceExpanded") = False, "Calendar recurrence semantics explicit"
    If Len(CStr(appointment("Subject"))) > 0 Then
      Set calendarMatches = client.SearchCalendar(CStr(appointment("Subject")), 3)
      Assert calendarMatches.Count > 0, "Live calendar subject search finds stored appointment"
    End If
  End If
  client.Disconnect
  Assert Not client.IsOpen, "Close releases native session"
  client.OpenProfile
  Assert client.IsOpen, "Profile can reopen"
  client.Disconnect
  LiveSmoke = "passed; address books=" & books.Count & "; contact matches=" & matchCount & "; inbox sample=" & messages.Count & "; calendar sample=" & calendar.Count & "; identity HRESULT=0x" & Hex$(identityError)
  Exit Function
Unavailable:
  Dim number As Long, description As String
  number = Err.Number: description = Err.description
  If Not client Is Nothing Then client.Disconnect
  If Not other Is Nothing Then other.Disconnect
  Err.Raise number, "LiveSmoke", description
End Function

Private Function Address(ByVal pointer As LongPtr) As LongPtr
  Address = pointer
End Function

Private Sub PutLong(ByVal pointer As LongPtr, ByVal value As Long)
  CopyMemory ByVal pointer, value, 4
End Sub

Private Sub PutPtr(ByVal pointer As LongPtr, ByVal value As LongPtr)
  CopyMemory ByVal pointer, value, PB
End Sub

Private Function GetLong(ByVal pointer As LongPtr) As Long
  CopyMemory GetLong, ByVal pointer, 4
End Function

Private Function GetPtr(ByVal pointer As LongPtr) As LongPtr
  CopyMemory GetPtr, ByVal pointer, PB
End Function

Private Function EscapeJSON(ByVal value As String) As String
  EscapeJSON = Replace(Replace(Replace(Replace(value, "\", "\\"), """", "\"""), vbCr, "\r"), vbLf, "\n")
End Function
