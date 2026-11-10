Attribute VB_Name = "stdOutlookSendCalendarTests"
Option Explicit

'All submission interfaces below are synthetic VBA vtables. No live message
'store is ever attached to these tests, including when --live is requested.
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As LongPtr)
Private Declare PtrSafe Function MAPIAllocateBuffer Lib "mapi32.dll" (ByVal size As Long, ByRef pointer As LongPtr) As Long
Private Declare PtrSafe Function MAPIAllocateMore Lib "mapi32.dll" (ByVal size As Long, ByVal parent As LongPtr, ByRef pointer As LongPtr) As Long
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal pointer As LongPtr) As Long
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal cp As Long, ByVal flags As Long, ByVal source As LongPtr, ByVal length As Long, ByVal dest As LongPtr, ByVal capacity As Long) As Long

#If Win64 Then
  Private Const PB As Long = 8
  Private Const SP As Long = 24
  Private Const SR As Long = 16
  Private Const RO As Long = 8
  Private Const RU As Long = 8
  Private Const RS As Long = 32
#Else
  Private Const PB As Long = 4
  Private Const SP As Long = 16
  Private Const SR As Long = 12
  Private Const RO As Long = 4
  Private Const RU As Long = 4
  Private Const RS As Long = 16
#End If

Private sessionVT(0 To 10) As LongPtr, bookVT(0 To 18) As LongPtr
Private storeVT(0 To 17) As LongPtr, folderVT(0 To 19) As LongPtr
Private messageVT(0 To 20) As LongPtr, tableVT(0 To 14) As LongPtr
Private storesTableVT(0 To 19) As LongPtr, storesTableObj As LongPtr
Private sessionObj As LongPtr, bookObj As LongPtr, storeObj As LongPtr
Private folderObj As LongPtr, messageObj As LongPtr, tableObj As LongPtr
Private assertions As Long, failure As String, failStage As Long
Private submits As Long, creates As Long, messageReleases As Long, folderReleases As Long
Private recipientCount As Long, recipientTypes As String, recipientAddresses As String
Private capturedBody As String, capturedSubject As String, capturedHTML As Boolean
Private missingName As Boolean, expectedLid As Long, restrictionCalls As Long
Private storeRowsRead As Boolean, denyWrite As Boolean, writeOpens As Long
Private statusTableVT(0 To 19) As LongPtr, statusVT(0 To 17) As LongPtr
Private statusTableObj As LongPtr, statusObj As LongPtr, statusRowsRead As Boolean
Private syncMode As Long, syncCalls As Long, syncFlags As Long, syncEntry As Long
Private syncFail As Boolean, wrongStatusType As Boolean, statusReleases As Long
Private automation As stdOutlookSyncFake, outlookRunning As Boolean, starts As Long
Private mismatchStore As Boolean, failCompare As Boolean, failStart As Boolean

Public Function RunTests() As Long
  Dim client As stdOutlook, stage As Long, number As Long, errorSource As String, errorDescription As String
  On Error GoTo Failed
  sessionVT(2) = Address(AddressOf ReleaseFake)
  sessionVT(4) = Address(AddressOf StoresTableFake): sessionVT(5) = Address(AddressOf OpenStoreFake)
  sessionVT(8) = Address(AddressOf StatusTableFake): sessionVT(9) = Address(AddressOf OpenStatusFake)
  sessionVT(10) = Address(AddressOf CompareIDsFake)
  statusTableVT(2) = Address(AddressOf ReleaseFake): statusTableVT(7) = Address(AddressOf StatusColumnsFake)
  statusTableVT(19) = Address(AddressOf StatusRowsFake): statusTableObj = VarPtr(statusTableVT(0))
  statusVT(2) = Address(AddressOf ReleaseFake): statusVT(17) = Address(AddressOf FlushFake)
  statusObj = VarPtr(statusVT(0))
  storesTableVT(2) = Address(AddressOf ReleaseFake): storesTableVT(7) = Address(AddressOf StoreColumnsFake)
  storesTableVT(19) = Address(AddressOf StoreRowsFake): storesTableObj = VarPtr(storesTableVT(0))
  bookVT(2) = Address(AddressOf ReleaseFake): bookVT(18) = Address(AddressOf OneOffFake)
  storeVT(2) = Address(AddressOf ReleaseFake): storeVT(5) = Address(AddressOf StorePropertiesFake)
  storeVT(17) = Address(AddressOf OpenOutboxFake)
  folderVT(2) = Address(AddressOf ReleaseFake): folderVT(19) = Address(AddressOf CreateMessageFake)
  folderVT(13) = Address(AddressOf NamedIDsFake)
  messageVT(2) = Address(AddressOf ReleaseFake): messageVT(8) = Address(AddressOf SetPropertiesFake)
  messageVT(19) = Address(AddressOf RecipientsFake): messageVT(20) = Address(AddressOf SubmitFake)
  tableVT(14) = Address(AddressOf CalendarRestrictionFake)
  sessionObj = VarPtr(sessionVT(0)): bookObj = VarPtr(bookVT(0)): storeObj = VarPtr(storeVT(0))
  folderObj = VarPtr(folderVT(0)): messageObj = VarPtr(messageVT(0)): tableObj = VarPtr(tableVT(0))
  Set client = New stdOutlook
  client.TestTransport VarPtr(sessionObj), VarPtr(bookObj), VarPtr(storeObj)
  assertions = 0
  Check client.UseOutlookSendReceive, "Outlook synchronization enabled by default"
  client.UseOutlookSendReceive = False 'Retain coverage of the explicit native option.

  ResetCapture
  client.SendEmail "ada@example.test; bob@example.test", "Unicode " & ChrW$(&H6771), "Line 1" & vbCrLf & "Zo" & ChrW$(&HEB), "cc@example.test", "bcc@example.test"
  Check failure = "", "Plain send native arguments: " & failure
  Check submits = 1 And creates = 1, "Plain send submits exactly once"
  Check capturedSubject = "Unicode " & ChrW$(&H6771), "Unicode subject retained"
  Check capturedBody = "Line 1" & vbCrLf & "Zo" & ChrW$(&HEB) And Not capturedHTML, "Plain Unicode body retained"
  Check recipientCount = 4 And recipientTypes = "1123", "To, CC, and BCC recipient types"
  Check recipientAddresses = "ada@example.test;bob@example.test;cc@example.test;bcc@example.test;", "Recipient addresses copied correctly"
  Check messageReleases = 1 And folderReleases = 1, "Send releases message and Outbox"
  Check syncCalls = 1 And syncEntry = 2 And syncFlags = &H3A, "Send prefers spooler and requests outgoing async flush"
  Check client.LastSendSyncRequested And client.LastSendSyncError = "" And statusReleases = 1, "Accepted sync reported separately and status released"

  ResetCapture
  client.SendEmailHTML "ada@example.test", "HTML", "<p>Zo" & ChrW$(&HEB) & " " & ChrW$(&H6771) & "</p>"
  Check failure = "", "HTML send native arguments: " & failure
  Check capturedHTML And capturedBody = "<p>Zo" & ChrW$(&HEB) & " " & ChrW$(&H6771) & "</p>", "HTML body is UTF-8 binary without a trailing null"
  Check submits = 1, "HTML submits once"
  Check syncCalls = 1 And client.LastSendSyncRequested, "HTML also automatically flushes"
  ResetCapture
  syncFail = True
  client.SendEmail "ada@example.test", "queued", "body"
  Check submits = 1 And syncCalls = 1 And Not client.LastSendSyncRequested, "Flush failure does not fail or retry the submitted email"
  Check Len(client.LastSendSyncError) > 0 And statusReleases = 1, "Flush error available separately with cleanup"
  ResetCapture
  client.SendEmail "ada@example.test", "batch", "body", flushQueue:=False
  Check submits = 1 And syncCalls = 0 And Not client.LastSendSyncRequested And client.LastSendSyncError = "", "Opt-out resets previous sync diagnostics"
  ResetCapture
  Check client.FlushQueues(True) = 1 And syncFlags = &H3E, "Explicit flush can include incoming mail"
  ResetCapture
  syncMode = 1
  Check client.FlushQueues() = 2 And syncCalls = 2 And syncFlags = &H1A And statusReleases = 2, "Without spooler both transports flushed synchronously"
  Check failure = "", "Transport flush ABI: " & failure
  ResetCapture
  syncMode = 2
  client.SendEmailHTML "ada@example.test", "unsupported", "body"
  Check submits = 1 And syncCalls = 0 And Len(client.LastSendSyncError) > 0, "Unsupported sync does not lose submission success"
  On Error Resume Next
  client.FlushQueues
  number = Err.Number
  On Error GoTo Failed
  Check number <> 0 And syncCalls = 0, "Explicit unsupported flush raises without calling unsupported methods"
  ResetCapture
  syncMode = 3
  client.SendEmail "ada@example.test", "no status table", "body"
  Check submits = 1 And syncCalls = 0 And Len(client.LastSendSyncError) > 0, "Status-table failure retained separately"
  ResetCapture
  wrongStatusType = True
  On Error Resume Next
  client.FlushQueues
  number = Err.Number
  On Error GoTo Failed
  Check number <> 0 And syncCalls = 0 And statusReleases = 1, "Wrong native object type rejected before vtable call"
  ResetCapture
  syncFail = True
  On Error Resume Next
  client.FlushQueues
  number = Err.Number
  On Error GoTo Failed
  Check number <> 0 And syncCalls = 1 And statusReleases = 1, "Explicit failed flush raises and releases without retry"
  ResetCapture
  client.SendEmail "", vbNullString, vbNullString, "", "bcc@example.test"
  Check submits = 1 And capturedBody = "" And capturedSubject = "", "BCC-only and empty subject/body supported"
  ResetCapture
  client.SendEmailHTML "ada@example.test", "", ""
  Check submits = 1 And capturedBody = "" And capturedHTML, "Empty HTML supported"

  For stage = 1 To 4
    ResetCapture
    failStage = stage
    On Error Resume Next
    client.SendEmail "ada@example.test", "fail", "body"
    number = Err.Number
    On Error GoTo Failed
    Check number <> 0, "Native failure surfaced at stage " & stage
    If stage < 4 Then
      Check submits = 0, "Preparation failure prevents submit at stage " & stage
    Else
      Check submits = 1, "Failed submission is not retried"
    End If
    Check messageReleases = 1 And folderReleases = 1, "Failure cleanup at stage " & stage
    Check syncCalls = 0 And Not client.LastSendSyncRequested, "No flush after failed submission at stage " & stage
  Next
  ResetCapture
  On Error Resume Next
  client.SendEmail "Ada <ada@example.test>", "subject", "body"
  number = Err.Number
  On Error GoTo Failed
  Check number = 5 And creates = 0, "Display-name syntax rejected before mutation"
  On Error Resume Next
  client.SendEmail "ada@example.test;", "subject", "body"
  number = Err.Number
  On Error GoTo Failed
  Check number = 5 And creates = 0, "Empty recipient rejected before mutation"
  On Error Resume Next
  client.SendEmail "", "subject", "body"
  number = Err.Number
  On Error GoTo Failed
  Check number = 5 And creates = 0, "No recipients rejected before mutation"
  On Error Resume Next
  client.SendEmail "a" & Chr$(127) & "@example.test", "subject", "body"
  number = Err.Number
  On Error GoTo Failed
  Check number = 5 And creates = 0, "DEL control character rejected before mutation"
  On Error Resume Next
  client.SendEmailHTML "ada@example.test", "subject", "<p>" & vbNullChar
  number = Err.Number
  On Error GoTo Failed
  Check number = 5 And creates = 0, "Null body rejected before mutation"

  expectedLid = &H8208&: missingName = False
  Check client.TestAppointmentTag(VarPtr(folderObj), expectedLid, 31) = &H9000001F, "Named calendar property mapped with store-specific ID"
  Check failure = "", "MAPINAMEID native structure: " & failure
  missingName = True
  Check client.TestAppointmentTag(VarPtr(folderObj), expectedLid, 31) = 0, "Missing named property does not create a mapping"
  restrictionCalls = 0
  client.TestCalendarRestriction VarPtr(tableObj), "planning", Array(Array("Subject", &H37001F), Array("Location", &H9000001F), Array("OrganizerName", &H42001F))
  Check restrictionCalls = 1 And failure = "", "Calendar OR/substring restriction ABI: " & failure
  client.Disconnect
  client.TestTransport VarPtr(sessionObj), VarPtr(bookObj), VarPtr(storeObj), False
  ResetCapture
  writeOpens = 0: denyWrite = False
  client.SendEmail "ada@example.test", "upgrade", "body"
  Check writeOpens = 1 And submits = 1 And failure = "", "Read-only store upgraded with MDB_WRITE before sending"
  TestOutlookSync client
  client.Disconnect
  client.TestTransport VarPtr(sessionObj), VarPtr(bookObj), VarPtr(storeObj), False
  ResetCapture
  denyWrite = True
  On Error Resume Next
  client.SendEmail "ada@example.test", "denied", "body"
  number = Err.Number
  On Error GoTo Failed
  Check number <> 0 And creates = 0 And submits = 0, "Denied write access cannot create or submit a message"
  client.Disconnect
  Set client = Nothing
  RunTests = assertions
  Exit Function
Failed:
  number = Err.Number: errorSource = Err.Source: errorDescription = Err.description
  If Not client Is Nothing Then client.Disconnect
  Err.Raise number, errorSource, errorDescription
End Function

Private Sub TestOutlookSync(ByVal client As stdOutlook)
  Dim number As Long
  client.UseOutlookSendReceive = True
  ResetCapture
  outlookRunning = True: automation.WindowCount = 1
  client.SendEmail "ada@example.test", "sync", "body"
  Check starts = 0 And automation.Requests = 1 And automation.Displays = 0, "Existing Outlook reused without opening another window"
  Check submits = 1 And syncCalls = 0 And client.LastSendSyncRequested, "Automatic send uses Outlook instead of native flush"
  Check Not automation.ShowProgress And failure = "", "Automatic sync hides progress and follows native message cleanup"
  Check automation.Order = "Namespace;Initialize;Store;Sync;", "Initialization and store validation precede sync"

  ResetCapture
  client.SendEmailHTML "ada@example.test", "start", "<p>body</p>"
  Check starts = 1 And automation.Displays = 1 And automation.Requests = 1, "HTML starts Outlook and opens a window for asynchronous lifetime"
  Check automation.Order = "Namespace;Initialize;Store;Display;Sync;", "New Outlook starts before sync"
  client.SendAndReceive True
  Check starts = 1 And automation.Displays = 1 And automation.Requests = 2 And automation.ShowProgress, "Explicit retry reuses Outlook and passes progress flag"
  Check submits = 1, "Explicit sync never resubmits email"

  ResetCapture
  client.SendEmail "ada@example.test", "batch", "body", flushQueue:=False
  Check starts = 0 And automation.Requests = 0 And Not client.LastSendSyncRequested, "Batch opt-out does not start Outlook"

  ResetCapture
  mismatchStore = True
  client.SendEmail "ada@example.test", "different profile", "body"
  Check submits = 1 And automation.Requests = 0 And automation.Displays = 0, "Mismatched default store prevents sync and display"
  Check Len(client.LastSendSyncError) > 0 And Not client.LastSendSyncRequested, "Store mismatch retained as sync error after successful submission"

  ResetCapture
  failCompare = True
  On Error Resume Next
  client.SendAndReceive
  number = Err.Number
  On Error GoTo 0
  Check number <> 0 And automation.Requests = 0, "Failed native store comparison prevents sync"

  ResetCapture
  automation.FailSync = True
  client.SendEmail "ada@example.test", "queued", "body"
  Check submits = 1 And automation.Requests = 1 And Len(client.LastSendSyncError) > 0, "Automation failure preserves queued message without retry"
  On Error Resume Next
  client.SendAndReceive
  number = Err.Number
  On Error GoTo 0
  Check number <> 0 And submits = 1 And automation.Requests = 2, "Explicit sync failure raises without creating a message"

  ResetCapture
  automation.FailInitialize = True
  client.SendEmail "ada@example.test", "initialization", "body"
  Check submits = 1 And automation.Requests = 0 And Len(client.LastSendSyncError) > 0, "Initialization failure isolated from submission"
  ResetCapture
  failStart = True
  client.SendEmail "ada@example.test", "startup", "body"
  Check submits = 1 And automation.Requests = 0 And Len(client.LastSendSyncError) > 0, "Outlook startup failure isolated from submission"
  ResetCapture
  failStage = 4
  On Error Resume Next
  client.SendEmail "ada@example.test", "submit failure", "body"
  number = Err.Number
  On Error GoTo 0
  Check number <> 0 And starts = 0 And automation.Requests = 0, "Failed submission never starts Outlook"
End Sub

Public Function GetRunningOutlookFake() As Object
  If outlookRunning Then Set GetRunningOutlookFake = automation
End Function

Public Function StartOutlookFake() As Object
  starts = starts + 1
  If failStart Then Err.Raise 429, , "Synthetic Outlook startup failure"
  outlookRunning = True
  Set StartOutlookFake = automation
End Function

Public Sub CheckSyncOrder()
  If creates > 0 And (messageReleases <> 1 Or folderReleases <> 1) Then failure = "Outlook sync before submitted message release"
End Sub

Public Function CompareIDsFake(ByVal self As LongPtr, ByVal cb1 As Long, ByVal entry1 As LongPtr, ByVal cb2 As Long, ByVal entry2 As LongPtr, ByVal flags As Long, ByRef same As Long) As Long
  If cb1 <> 4 Or cb2 <> 4 Or flags <> 0 Then failure = "CompareEntryIDs arguments"
  If LongAt(entry1) <> 99 Or LongAt(entry2) <> 99 Then failure = "CompareEntryIDs store IDs"
  If failCompare Then
    CompareIDsFake = &H80004005
  ElseIf Not mismatchStore Then
    same = 1
  End If
End Function

Private Sub ResetCapture()
  failure = "": failStage = 0: submits = 0: creates = 0
  messageReleases = 0: folderReleases = 0: recipientCount = 0
  recipientTypes = "": recipientAddresses = "": capturedBody = "": capturedSubject = ""
  capturedHTML = False
  syncMode = 0: syncCalls = 0: syncFlags = 0: syncEntry = 0
  syncFail = False: wrongStatusType = False: statusReleases = 0
  Set automation = New stdOutlookSyncFake
  outlookRunning = False: starts = 0: mismatchStore = False: failCompare = False: failStart = False
End Sub

Private Sub Check(ByVal condition As Boolean, ByVal label As String)
  Test.Assert label, condition
  If Not condition Then Err.Raise vbObjectError + 2002, "stdOutlookSendCalendarTests", label
  assertions = assertions + 1
End Sub

Public Function ReleaseFake(ByVal self As LongPtr) As Long
  If self = VarPtr(messageObj) Then messageReleases = messageReleases + 1
  If self = VarPtr(folderObj) Then folderReleases = folderReleases + 1
  If self = VarPtr(statusObj) Then statusReleases = statusReleases + 1
End Function

Public Function StatusTableFake(ByVal self As LongPtr, ByVal flags As Long, ByRef table As LongPtr) As Long
  If flags <> 0 Then failure = "GetStatusTable flags"
  If syncMode = 3 Then
    StatusTableFake = &H80004005
    Exit Function
  End If
  statusRowsRead = False
  table = VarPtr(statusTableObj)
End Function

Public Function StatusColumnsFake(ByVal self As LongPtr, ByVal tags As LongPtr, ByVal flags As Long) As Long
  If LongAt(tags) <> 3 Then
    failure = "Status columns count": StatusColumnsFake = &H80004005
    Exit Function
  End If
  If LongAt(tags + 4) <> &HFFF0102 Or LongAt(tags + 8) <> &H3E030003 Or LongAt(tags + 12) <> &H3E020003 Then failure = "Status column tags"
End Function

Public Function StatusRowsFake(ByVal self As LongPtr, ByVal requested As Long, ByVal flags As Long, ByRef rows As LongPtr) As Long
  Dim props As LongPtr, entry As LongPtr, row As LongPtr, i As Long, kind As Long, methods As Long
  On Error GoTo Failed
  If MAPIAllocateBuffer(RO + SR * 2, rows) <> 0 Then GoTo Failed
  PutLong rows, 0
  If statusRowsRead Then Exit Function
  For i = 0 To 1
    If MAPIAllocateBuffer(SP * 3, props) <> 0 Then GoTo Failed
    row = rows + RO + SR * i
    PutLong row, 0: PutLong row + 4, 3: PutPtr row + 8, props
    PutLong rows, i + 1
    If MAPIAllocateMore(4, props, entry) <> 0 Then GoTo Failed
    PutLong entry, i + 1
    PutLong props, &HFFF0102: PutLong props + 8, 4: PutPtr props + 8 + PB, entry
    kind = 36: methods = 8
    If syncMode = 0 And i = 1 Then kind = 37
    If syncMode = 2 Then
      If i = 0 Then methods = 0 Else kind = 34
    End If
    PutLong props + SP, &H3E030003: PutLong props + SP + 8, kind
    PutLong props + SP * 2, &H3E020003: PutLong props + SP * 2 + 8, methods
  Next
  statusRowsRead = True
  Exit Function
Failed:
  failure = "Status rows callback": StatusRowsFake = &H80004005
End Function

Public Function OpenStatusFake(ByVal self As LongPtr, ByVal cb As Long, ByVal entry As LongPtr, ByVal iid As LongPtr, ByVal flags As Long, ByRef kind As Long, ByRef status As LongPtr) As Long
  If cb <> 4 Or iid <> 0 Or flags <> 0 Then failure = "OpenEntry status arguments"
  syncEntry = LongAt(entry)
  kind = 10
  If wrongStatusType Then kind = 3
  status = VarPtr(statusObj)
End Function

Public Function FlushFake(ByVal self As LongPtr, ByVal hwnd As LongPtr, ByVal cb As Long, ByVal entry As LongPtr, ByVal flags As Long) As Long
  syncCalls = syncCalls + 1: syncFlags = flags
  If hwnd <> 0 Or cb <> 0 Or entry <> 0 Then failure = "FlushQueues arguments"
  If creates > 0 And (messageReleases <> 1 Or folderReleases <> 1) Then failure = "Flush before releasing submitted message"
  If syncFail Then FlushFake = &H8004010B 'MAPI_E_BUSY
End Function

Public Function OneOffFake(ByVal self As LongPtr, ByVal name As LongPtr, ByVal addressType As LongPtr, ByVal address As LongPtr, ByVal flags As Long, ByRef cb As Long, ByRef entry As LongPtr) As Long
  On Error GoTo Failed
  If TextAt(addressType) <> "SMTP" Or TextAt(address) <> TextAt(name) Or flags <> (&H80000000 Or &H10000) Then failure = "CreateOneOff arguments"
  cb = 4
  OneOffFake = MAPIAllocateBuffer(cb, entry)
  If OneOffFake = 0 Then PutLong entry, &H12345678
  Exit Function
Failed:
  failure = "CreateOneOff callback": OneOffFake = &H80004005
End Function

Public Function StorePropertiesFake(ByVal self As LongPtr, ByVal tags As LongPtr, ByVal flags As Long, ByRef count As Long, ByRef props As LongPtr) As Long
  Dim entry As LongPtr, i As Long
  On Error GoTo Failed
  count = LongAt(tags)
  If count = 1 And LongAt(tags + 4) = &HFFF0102 Then
    If MAPIAllocateBuffer(SP, props) <> 0 Then GoTo Failed
    If MAPIAllocateMore(4, props, entry) <> 0 Then GoTo Failed
    PutLong entry, 99
    PutLong props, &HFFF0102: PutLong props + 8, 4: PutPtr props + 8 + PB, entry
    Exit Function
  End If
  If count <> 2 Or LongAt(tags + 4) <> &H35E20102 Or LongAt(tags + 8) <> &H35E40102 Then failure = "Outbox/Sent folder properties"
  If MAPIAllocateBuffer(SP * 2, props) <> 0 Then GoTo Failed
  For i = 0 To 1
    If MAPIAllocateMore(4, props, entry) <> 0 Then GoTo Failed
    PutLong entry, i + 1
    PutLong props + i * SP, LongAt(tags + 4 + i * 4)
    PutLong props + i * SP + 8, 4
    PutPtr props + i * SP + 8 + PB, entry
  Next
  Exit Function
Failed:
  failure = "Store properties callback": StorePropertiesFake = &H80004005
End Function

Public Function OpenOutboxFake(ByVal self As LongPtr, ByVal cb As Long, ByVal entry As LongPtr, ByVal iid As LongPtr, ByVal flags As Long, ByRef kind As Long, ByRef folder As LongPtr) As Long
  If cb <> 4 Or LongAt(entry) <> 1 Or iid <> 0 Or flags <> 1 Then failure = "Writable Outbox open"
  kind = 3: folder = VarPtr(folderObj)
End Function

Public Function CreateMessageFake(ByVal self As LongPtr, ByVal iid As LongPtr, ByVal flags As Long, ByRef message As LongPtr) As Long
  If iid <> 0 Or flags <> 0 Then failure = "CreateMessage arguments"
  creates = creates + 1: message = VarPtr(messageObj)
End Function

Public Function SetPropertiesFake(ByVal self As LongPtr, ByVal count As Long, ByVal props As LongPtr, ByRef problems As LongPtr) As Long
  Dim prop As LongPtr, length As Long, textPointer As LongPtr, chars As Long
  On Error GoTo Failed
  If count <> 6 Then failure = "Message property count"
  If LongAt(props) <> &H1A001F Or TextAt(PtrAt(props + 8)) <> "IPM.Note" Then failure = "Message class"
  capturedSubject = TextAt(PtrAt(props + SP + 8))
  prop = props + SP * 2
  capturedHTML = (LongAt(prop) = &H10130102)
  If capturedHTML Then
    length = LongAt(prop + 8): textPointer = PtrAt(prop + 8 + PB)
    If length > 0 Then
      chars = MultiByteToWideChar(65001, 8, textPointer, length, 0, 0)
      capturedBody = String$(chars, vbNullChar)
      MultiByteToWideChar 65001, 8, textPointer, length, StrPtr(capturedBody), chars
    End If
  Else
    If LongAt(prop) <> &H1000001F Then failure = "Plain body tag"
    capturedBody = TextAt(PtrAt(prop + 8))
  End If
  prop = props + SP * 3
  If LongAt(prop) <> &HE0A0102 Or LongAt(prop + 8) <> 4 Or LongAt(PtrAt(prop + 8 + PB)) <> 2 Then failure = "Sent Items entry ID"
  If LongAt(props + SP * 4 + 8) <> 8 Or LongAt(props + SP * 5 + 8) <> 65001 Then failure = "Message flags/code page"
  If failStage = 1 Then SetPropertiesFake = &H80004005
  If failStage = 2 Then
    If MAPIAllocateBuffer(16, problems) <> 0 Then GoTo Failed
    PutLong problems, 1: PutLong problems + 4, 2
    PutLong problems + 8, &H1000001F: PutLong problems + 12, &H80070005
  End If
  Exit Function
Failed:
  failure = "SetProps callback": SetPropertiesFake = &H80004005
End Function

Public Function RecipientsFake(ByVal self As LongPtr, ByVal flags As Long, ByVal list As LongPtr) As Long
  Dim row As LongPtr, props As LongPtr, i As Long
  On Error GoTo Failed
  If flags <> 2 Then failure = "ModifyRecipients flags"
  recipientCount = LongAt(list)
  For i = 0 To recipientCount - 1
    row = list + RO + i * SR: props = PtrAt(row + 8)
    If LongAt(row + 4) <> 5 Then failure = "ADRENTRY property count"
    recipientTypes = recipientTypes & LongAt(props + 8)
    recipientAddresses = recipientAddresses & TextAt(PtrAt(props + SP * 3 + 8)) & ";"
    If LongAt(props + SP * 4) <> &HFFF0102 Or LongAt(PtrAt(props + SP * 4 + 8 + PB)) <> &H12345678 Then failure = "Recipient one-off ID"
  Next
  If failStage = 3 Then RecipientsFake = &H80004005
  Exit Function
Failed:
  failure = "Recipients callback": RecipientsFake = &H80004005
End Function

Public Function SubmitFake(ByVal self As LongPtr, ByVal flags As Long) As Long
  submits = submits + 1
  If flags <> 0 Then failure = "Submit flags"
  If failStage = 4 Then SubmitFake = &H80004005
End Function

Public Function NamedIDsFake(ByVal self As LongPtr, ByVal count As Long, ByVal names As LongPtr, ByVal flags As Long, ByRef tags As LongPtr) As Long
  Dim name As LongPtr, guid As LongPtr
  On Error GoTo Failed
  name = PtrAt(names): guid = PtrAt(name)
  If count <> 1 Or flags <> 0 Or LongAt(name + PB) <> 0 Or LongAt(name + PB * 2) <> expectedLid Then failure = "MAPINAMEID count/kind/LID/flags"
  If LongAt(guid) <> &H62002 Or LongAt(guid + 4) <> 0 Or LongAt(guid + 8) <> &HC0 Or LongAt(guid + 12) <> &H46000000 Then failure = "Appointment GUID"
  If MAPIAllocateBuffer(8, tags) <> 0 Then GoTo Failed
  PutLong tags, 1
  If missingName Then
    PutLong tags + 4, 10: NamedIDsFake = &H40380
  Else
    PutLong tags + 4, &H90000000
  End If
  Exit Function
Failed:
  failure = "NamedIDs callback": NamedIDsFake = &H80004005
End Function

Public Function CalendarRestrictionFake(ByVal self As LongPtr, ByVal restriction As LongPtr, ByVal flags As Long) As Long
  Dim children As LongPtr, child As LongPtr, prop As LongPtr, i As Long, expected As Variant
  On Error GoTo Failed
  restrictionCalls = restrictionCalls + 1
  If LongAt(restriction) <> 1 Or LongAt(restriction + RU) <> 3 Or flags <> 0 Then failure = "Calendar OR restriction"
  children = PtrAt(restriction + RU + PB)
  expected = Array(&H37001F, &H9000001F, &H42001F)
  For i = 0 To 2
    child = children + RS * i: prop = PtrAt(child + RU + 8)
    If LongAt(child) <> 3 Or LongAt(child + RU) <> &H10001 Or LongAt(child + RU + 4) <> expected(i) Then failure = "Calendar content restriction"
    If LongAt(prop) <> expected(i) Or TextAt(PtrAt(prop + 8)) <> "planning" Then failure = "Calendar search value"
  Next
  Exit Function
Failed:
  failure = "Calendar restriction callback": CalendarRestrictionFake = &H80004005
End Function

Private Function Address(ByVal pointer As LongPtr) As LongPtr
  Address = pointer
End Function

Public Function StoresTableFake(ByVal self As LongPtr, ByVal flags As Long, ByRef table As LongPtr) As Long
  storeRowsRead = False
  table = VarPtr(storesTableObj)
End Function

Public Function StoreColumnsFake(ByVal self As LongPtr, ByVal tags As LongPtr, ByVal flags As Long) As Long
  If LongAt(tags) <> 2 Or LongAt(tags + 4) <> &HFFF0102 Or LongAt(tags + 8) <> &H3400000B Then failure = "Default store columns"
End Function

Public Function StoreRowsFake(ByVal self As LongPtr, ByVal requested As Long, ByVal flags As Long, ByRef rows As LongPtr) As Long
  Dim props As LongPtr, entry As LongPtr, yes As Integer
  On Error GoTo Failed
  If MAPIAllocateBuffer(RO + SR, rows) <> 0 Then GoTo Failed
  PutLong rows, 0
  If storeRowsRead Then Exit Function
  If MAPIAllocateBuffer(SP * 2, props) <> 0 Then GoTo Failed
  PutLong rows, 1: PutLong rows + RO, 0: PutLong rows + RO + 4, 2: PutPtr rows + RO + 8, props
  If MAPIAllocateMore(4, props, entry) <> 0 Then GoTo Failed
  PutLong entry, 99
  PutLong props, &HFFF0102: PutLong props + 8, 4: PutPtr props + 8 + PB, entry
  PutLong props + SP, &H3400000B
  yes = 1: CopyMemory ByVal props + SP + 8, yes, 2
  storeRowsRead = True
  Exit Function
Failed:
  failure = "Stores table callback": StoreRowsFake = &H80004005
End Function

Public Function OpenStoreFake(ByVal self As LongPtr, ByVal hwnd As LongPtr, ByVal cb As Long, ByVal entry As LongPtr, ByVal iid As LongPtr, ByVal flags As Long, ByRef store As LongPtr) As Long
  If hwnd <> 0 Or cb <> 4 Or LongAt(entry) <> 99 Or iid <> 0 Or flags <> 5 Then failure = "MDB_NO_DIALOG | MDB_WRITE required"
  writeOpens = writeOpens + 1
  If denyWrite Then
    OpenStoreFake = &H80070005
  Else
    store = VarPtr(storeObj)
  End If
End Function

Private Function LongAt(ByVal pointer As LongPtr) As Long
  CopyMemory LongAt, ByVal pointer, 4
End Function

Private Function PtrAt(ByVal pointer As LongPtr) As LongPtr
  CopyMemory PtrAt, ByVal pointer, PB
End Function

Private Sub PutLong(ByVal pointer As LongPtr, ByVal value As Long)
  CopyMemory ByVal pointer, value, 4
End Sub

Private Sub PutPtr(ByVal pointer As LongPtr, ByVal value As LongPtr)
  CopyMemory ByVal pointer, value, PB
End Sub

Private Function TextAt(ByVal pointer As LongPtr) As String
  Dim length As Long
  If pointer = 0 Then Exit Function
  length = lstrlenW(pointer)
  If length = 0 Then Exit Function
  TextAt = String$(length, vbNullChar)
  CopyMemory ByVal StrPtr(TextAt), ByVal pointer, length * 2
End Function
