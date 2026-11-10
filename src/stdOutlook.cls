VERSION 1.0 CLASS
BEGIN
  MultiUse = -1 'True
END
Attribute VB_Name = "stdOutlook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'@module
'@ai AI:X
'@remark AI-written changes: #189. Conservative AI:X disclosure; historical human scrutiny is unverified and may vary by change.
'@description Access classic Outlook profiles through native Extended MAPI from a single VBA7 class.
'@remark Experimental: 64-bit Excel tested; 32-bit runtime and live synchronization remain unverified.
'@remark Data access and submission avoid guarded Outlook automation. Send/Receive uses late-bound automation.
'@requires Windows VBA7 and classic Outlook with matching Office bitness and a configured profile.
'@example Set client = stdOutlook.Create()

'Extended MAPI client for Windows VBA7, both Office architectures.
'Import this file; no Outlook reference, stdVBA dependency, or helper DLL.
'Data/submission use MAPI; automatic sync uses unguarded Outlook automation.
'MAPI uses IUnknown vtables, NOT the guarded Outlook automation object model.
'See docs/stdOutlook.md for provider limitations and tests/ for ABI validation.

Private Declare PtrSafe Function MAPIInitialize Lib "mapi32.dll" (ByVal pInit As LongPtr) As Long
Private Declare PtrSafe Sub MAPIUninitialize Lib "mapi32.dll" ()
Private Declare PtrSafe Function MAPILogonEx Lib "mapi32.dll" (ByVal hwnd As LongPtr, ByVal profile As LongPtr, ByVal password As LongPtr, ByVal flags As Long, ByRef session As LongPtr) As Long
Private Declare PtrSafe Function MAPIFreeBuffer Lib "mapi32.dll" (ByVal buffer As LongPtr) As Long
Private Declare PtrSafe Function MAPIAllocateBuffer Lib "mapi32.dll" (ByVal size As Long, ByRef buffer As LongPtr) As Long
Private Declare PtrSafe Function MAPIAllocateMore Lib "mapi32.dll" (ByVal size As Long, ByVal parent As LongPtr, ByRef buffer As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As LongPtr)
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal value As LongPtr) As Long
Private Declare PtrSafe Function lstrlenA Lib "kernel32" (ByVal value As LongPtr) As Long
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal codePage As Long, ByVal flags As Long, ByVal source As LongPtr, ByVal sourceLength As Long, ByVal destination As LongPtr, ByVal capacity As Long) As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal codePage As Long, ByVal flags As Long, ByVal source As LongPtr, ByVal sourceLength As Long, ByVal destination As LongPtr, ByVal capacity As Long, ByVal defaultChar As LongPtr, ByVal usedDefault As LongPtr) As Long
Private Declare PtrSafe Function FileTimeToSystemTime Lib "kernel32" (ByVal fileTime As LongPtr, ByRef systemTime As SYSTEMTIME) As Long
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32.dll" (ByVal instance As LongPtr, ByVal vtableOffset As LongPtr, ByVal convention As Long, ByVal returnType As Integer, ByVal argumentCount As Long, ByVal argumentTypes As LongPtr, ByVal argumentValues As LongPtr, ByRef result As Variant) As Long

Private Type SYSTEMTIME
  year As Integer
  month As Integer
  dayOfWeek As Integer
  day As Integer
  hour As Integer
  minute As Integer
  second As Integer
  milliseconds As Integer
End Type

#If Win64 Then
  Private Const PTR_BYTES As Long = 8
  Private Const PROP_BYTES As Long = 24
  Private Const ROW_BYTES As Long = 16
  Private Const ROWSET_OFFSET As Long = 8
  Private Const RESTRICTION_BYTES As Long = 32
  Private Const RESTRICTION_UNION As Long = 8
  Private Const BINARY_POINTER_OFFSET As Long = 8
#Else
  Private Const PTR_BYTES As Long = 4
  Private Const PROP_BYTES As Long = 16
  Private Const ROW_BYTES As Long = 12
  Private Const ROWSET_OFFSET As Long = 4
  Private Const RESTRICTION_BYTES As Long = 16
  Private Const RESTRICTION_UNION As Long = 4
  Private Const BINARY_POINTER_OFFSET As Long = 4
#End If

Private Const MAPI_UNICODE As Long = &H80000000
Private Const MAPI_EXTENDED As Long = &H20&
Private Const MAPI_USE_DEFAULT As Long = &H40&
Private Const MAPI_NEW_SESSION As Long = &H2&
Private Const MAPI_MODIFY As Long = 1
Private Const MDB_WRITE As Long = 4
Private Const MAPI_E_BAD_CHARWIDTH As Long = &H80040103
Private Const MAPI_E_NOT_FOUND As Long = &H8004010F
Private Const MAPI_E_NO_SUPPORT As Long = &H80040102
Private Const MAPI_E_TOO_BIG As Long = &H80040305
Private Const AB_NO_DIALOG As Long = 1
Private Const MDB_NO_DIALOG As Long = 1
Private Const CONVENIENT_DEPTH As Long = 1
Private Const AB_RECIPIENTS As Long = 1
Private Const BATCH_SIZE As Long = 64

'Vtable slots are zero based, including IUnknown's three methods.
Private Const VT_RELEASE As Long = 2
Private Const VT_GET_PROPS As Long = 5
Private Const VT_SET_PROPS As Long = 8
Private Const VT_NAMED_IDS As Long = 13
Private Const VT_AB_ONE_OFF As Long = 18
Private Const VT_FOLDER_CREATE_MESSAGE As Long = 19
Private Const VT_MESSAGE_RECIPIENTS As Long = 19
Private Const VT_MESSAGE_SUBMIT As Long = 20
Private Const VT_SESSION_STORES As Long = 4
Private Const VT_SESSION_OPEN_STORE As Long = 5
Private Const VT_SESSION_ADDRESS_BOOK As Long = 6
Private Const VT_SESSION_STATUS_TABLE As Long = 8
Private Const VT_SESSION_OPEN_ENTRY As Long = 9
Private Const VT_SESSION_COMPARE_IDS As Long = 10
Private Const VT_STATUS_FLUSH As Long = 17
Private Const VT_SESSION_IDENTITY As Long = 16
Private Const VT_AB_OPEN_ENTRY As Long = 14
Private Const VT_CONTAINER_CONTENTS As Long = 14
Private Const VT_CONTAINER_HIERARCHY As Long = 15
Private Const VT_STORE_OPEN_ENTRY As Long = 17
Private Const VT_STORE_RECEIVE_FOLDER As Long = 19
Private Const VT_TABLE_COLUMNS As Long = 7
Private Const VT_TABLE_RESTRICT As Long = 14
Private Const VT_TABLE_SORT As Long = 17
Private Const VT_TABLE_ROWS As Long = 19

Private Type TThis
  Session As LongPtr
  AddressBook As LongPtr
  Store As LongPtr
  StoreWritable As Boolean
  Initialized As Boolean
  Warnings As Collection
  SearchTruncated As Boolean
  CalendarSearchTruncated As Boolean
  LastSendSyncError As String
  LastSendSyncRequested As Boolean
  UseOutlookSendReceive As Boolean
End Type
Private This As TThis

'Empty profile: share the current MAPI session, or use the default profile.
'Explicit profile: request a separate session so MAPI cannot ignore its name.
'Never enables logon UI. Credentials/profile setup remain Outlook's job.
'@constructor
'@param profileName - Profile name; empty shares the current session or uses the default. Explicit names request a separate MAPI session. No native logon UI.
'@returns - An initialized stdOutlook instance owning its MAPI handles.
'@example Set client = stdOutlook.Create()
Public Function Create(Optional ByVal profileName As String = "") As stdOutlook
  Dim client As stdOutlook
  Set client = New stdOutlook
  client.protInit profileName
  Set Create = client
End Function

'Release the current native handles and open another profile without logging off Outlook.
'@param profileName - Profile name; empty shares the current session or uses the default. Explicit names request a separate MAPI session. No native logon UI.
Public Sub OpenProfile(Optional ByVal profileName As String = "")
  protInit profileName
End Sub

'Initialize an instance with a native MAPI profile; prefer Create or OpenProfile.
'@protected
'@param profileName - Profile name; empty shares the current session or uses the default. Explicit names request a separate MAPI session. No native logon UI.
Public Sub protInit(Optional ByVal profileName As String = "")
  Dim flags As Long, number As Long, description As String
  If InStr(profileName, vbNullChar) > 0 Or Len(profileName) > 64 Then
    Err.Raise 5, "stdOutlook.OpenProfile", "Invalid profile name."
  End If
  Me.Disconnect
  On Error GoTo Failed
  CheckHR MAPIInitialize(0), "MAPIInitialize (requires classic Outlook and matching Office bitness)"
  This.Initialized = True
  'A mail-capable session is required for SubmitMessage, including spooler-based
  'transports. Store handles remain read-only until a send explicitly upgrades.
  flags = MAPI_EXTENDED Or MAPI_UNICODE
  If Len(profileName) = 0 Then
    flags = flags Or MAPI_USE_DEFAULT
  Else
    flags = flags Or MAPI_NEW_SESSION
  End If
  CheckHR MAPILogonEx(0, StrPtr(profileName), 0, flags, This.Session), "MAPILogonEx (open/sign in to this profile in classic Outlook first)"
  RequirePointer This.Session, "MAPILogonEx"
  Exit Sub
Failed:
  number = Err.Number: description = Err.description
  Me.Disconnect
  Err.Raise number, "stdOutlook.OpenProfile", description
End Sub

'Whether this instance owns an open MAPI session.
'@returns - True if a native session is open.
Public Property Get IsOpen() As Boolean
  IsOpen = (This.Session <> 0)
End Property

'Release owned native handles and balance MAPI initialization; safe to call repeatedly. Does not quit Outlook.
Public Sub Disconnect()
  'Release child objects before their session. Do not Logoff a shared session:
  'that can interfere with Outlook or another stdOutlook instance.
  On Error Resume Next
  ReleaseObject This.Store
  This.StoreWritable = False
  ReleaseObject This.AddressBook
  ReleaseObject This.Session
  If This.Initialized Then
    This.Initialized = False
    MAPIUninitialize
  End If
  On Error GoTo 0
End Sub

'Initialize diagnostics and the default synchronization policy.
Private Sub Class_Initialize()
  Set This.Warnings = New Collection
  This.UseOutlookSendReceive = True
End Sub

'Release this instance without logging off a shared MAPI session.
Private Sub Class_Terminate()
  Me.Disconnect
End Sub

'Primary profile identity, not every account's sending address.
'Providers without an identity (for example some PST/IMAP profiles) can fail.
'@returns Collection - Keyed identity properties including Email, DisplayName, EntryID and PropertyErrors; not all account identities.
Public Function CurrentUser() As Collection
  Dim cb As Long, entry As LongPtr, person As LongPtr, kind As Long
  Dim number As Long, description As String
  EnsureAddressBook
  On Error GoTo Failed
  CheckHR CallVTable(This.Session, VT_SESSION_IDENTITY, VarPtr(cb), VarPtr(entry)), "IMAPISession.QueryIdentity"
  If cb <= 0 Then Err.Raise MAPI_E_NOT_FOUND, "stdOutlook", "The profile has no primary identity."
  RequirePointer entry, "QueryIdentity entry ID"
  CheckHR CallVTable(This.AddressBook, VT_AB_OPEN_ENTRY, cb, entry, CLngPtr(0), 0&, VarPtr(kind), VarPtr(person)), "IAddrBook.OpenEntry(identity)"
  Set CurrentUser = ReadPerson(person)
Done:
  On Error GoTo 0
  ReleaseObject person
  FreeBuffer entry
  If number <> 0 Then Err.Raise number, "stdOutlook.CurrentUser", description
  Exit Function
Failed:
  number = Err.Number: description = Err.description
  Resume Done
End Function

'Resolve the profile identity SMTP address; raise if none is supplied by the provider.
'@returns - The resolved primary SMTP address, never an Exchange legacy DN.
Public Property Get CurrentUserEmail() As String
  Dim person As Collection
  Set person = CurrentUser()
  CurrentUserEmail = person("Email")
  If Len(CurrentUserEmail) = 0 Then
    Err.Raise MAPI_E_NOT_FOUND, "stdOutlook.CurrentUserEmail", "The identity provider did not supply an SMTP address. Inspect CurrentUser and its PropertyErrors."
  End If
End Property

'Return all address-book containers, including nested Outlook Contacts books.
'Only folders enabled as an Outlook Address Book are visible through this API.
'@returns Collection - One keyed Collection per container: EntryID, DisplayName, Flags, Depth and PropertyErrors.
Public Function AddressBooks() As Collection
  Dim root As LongPtr, table As LongPtr, kind As Long, capped As Boolean
  Dim number As Long, description As String
  EnsureAddressBook
  On Error GoTo Failed
  CheckHR CallVTable(This.AddressBook, VT_AB_OPEN_ENTRY, 0&, CLngPtr(0), CLngPtr(0), 0&, VarPtr(kind), VarPtr(root)), "IAddrBook.OpenEntry(root)"
  CheckHR CallVTable(root, VT_CONTAINER_HIERARCHY, CONVENIENT_DEPTH, VarPtr(table)), "IABContainer.GetHierarchyTable"
  Set AddressBooks = ReadTable(table, BookColumns(), 10000, capped)
  If capped Then Err.Raise MAPI_E_TOO_BIG, "stdOutlook", "Address-book hierarchy exceeded 10000 containers."
Done:
  On Error GoTo 0
  ReleaseObject table
  ReleaseObject root
  If number <> 0 Then Err.Raise number, "stdOutlook.AddressBooks", description
  Exit Function
Failed:
  number = Err.Number: description = Err.description
  Resume Done
End Function

'Provider-side ambiguous-name resolution (ANR), usually name/alias prefixes.
'Not arbitrary substring matching: email matching depends on the provider.
'containerEntryID optionally limits the search to one AddressBooks() entry.
'Warnings and SearchTruncated describe partial results; always inspect them.
'@param query - Search text; matching rules are described by this method.
'@param maxResults - Maximum returned records, from 1 to 10000.
'@param containerEntryID - Hexadecimal AddressBooks entry ID; empty searches all recipient containers.
'@returns Collection - Copied keyed contact records with Email, EntryID, AddressBook and PropertyErrors; inspect SearchWarnings and SearchTruncated.
Public Function SearchContacts(ByVal query As String, Optional ByVal maxResults As Long = 100, Optional ByVal containerEntryID As String = "") As Collection
  Dim books As Collection, book As Collection, results As New Collection
  Dim seen As New Collection, matches As Collection, person As Collection
  Dim entry As Variant, capped As Boolean
  Set This.Warnings = New Collection
  This.SearchTruncated = False
  query = Trim$(query)
  If Len(query) = 0 Or InStr(query, vbNullChar) > 0 Then Err.Raise 5, "stdOutlook.SearchContacts", "Supply a nonempty name, alias, or email search."
  ValidateLimit maxResults
  EnsureAddressBook
  If Len(containerEntryID) > 0 Then
    Set books = New Collection
    Set book = New Collection
    book.Add containerEntryID, "EntryID"
    book.Add "Selected address book", "DisplayName"
    book.Add AB_RECIPIENTS, "Flags"
    books.Add book
  Else
    Set books = AddressBooks()
  End If
  For Each book In books
    If (CLng(Val(CStr(book("Flags")))) And AB_RECIPIENTS) <> 0 Then
      Set matches = SearchBook(CStr(book("EntryID")), query, maxResults + 1, capped)
      If capped Then This.SearchTruncated = True
      For Each person In matches
        entry = person("EntryID")
        If Not HasKey(seen, CStr(entry)) Then
          If results.Count = maxResults Then
            This.SearchTruncated = True
            Exit For
          End If
          seen.Add True, CStr(entry)
          person.Add book("DisplayName"), "AddressBook"
          results.Add person
        End If
      Next
      If results.Count = maxResults Then
        'Conservative: remaining containers could contain additional matches.
        This.SearchTruncated = True
        Exit For
      End If
    End If
  Next
  Set SearchContacts = results
End Function

'Whether the latest contact search may have omitted matching records.
'@returns - True if results may be incomplete; conservative when the result limit is reached.
Public Property Get SearchTruncated() As Boolean
  SearchTruncated = This.SearchTruncated
End Property

'Return an independent copy of warnings from the latest contact search.
'@returns Collection - Warning strings; mutating this copy cannot change diagnostic state.
Public Property Get SearchWarnings() As Collection
  'Return a copy so callers cannot mutate the diagnostic state.
  Dim result As New Collection, warning As Variant
  For Each warning In This.Warnings
    result.Add warning
  Next
  Set SearchWarnings = result
End Property

'Read copied properties of an address-book entry without guarded Outlook automation.
'@param entryID - Nonempty, even-length hexadecimal MAPI entry ID in the current profile.
'@returns Collection - Copied person fields including Email and PropertyErrors. Missing values are Empty.
Public Function ReadContact(ByVal entryID As String) As Collection
  Dim bytes() As Byte, person As LongPtr, kind As Long
  Dim number As Long, description As String
  EnsureAddressBook
  bytes = DecodeEntryID(entryID)
  On Error GoTo Failed
  CheckHR CallVTable(This.AddressBook, VT_AB_OPEN_ENTRY, CLng(UBound(bytes) + 1), VarPtr(bytes(0)), CLngPtr(0), 0&, VarPtr(kind), VarPtr(person)), "IAddrBook.OpenEntry(contact)"
  Set ReadContact = ReadPerson(person)
Done:
  On Error GoTo 0
  ReleaseObject person
  If number <> 0 Then Err.Raise number, "stdOutlook.ReadContact", description
  Exit Function
Failed:
  number = Err.Number: description = Err.description
  Resume Done
End Function

'Optional mail access: newest messages in the default store's Inbox.
'Does not mark messages read. Times are UTC. No bodies in the table result.
'@param maxResults - Maximum returned records, from 1 to 10000.
'@returns Collection - Newest message metadata records, without bodies; times are UTC and entries are not marked read.
Public Function Inbox(Optional ByVal maxResults As Long = 50) As Collection
  Dim folder As LongPtr, table As LongPtr, entry As LongPtr, cb As Long, kind As Long
  Dim messageClass As String, sort(0 To 4) As Long, capped As Boolean
  Dim number As Long, description As String
  ValidateLimit maxResults
  EnsureStore
  On Error GoTo Failed
  messageClass = "IPM.Note"
  CheckHR CallVTable(This.Store, VT_STORE_RECEIVE_FOLDER, StrPtr(messageClass), MAPI_UNICODE, VarPtr(cb), VarPtr(entry), CLngPtr(0)), "IMsgStore.GetReceiveFolder"
  If cb <= 0 Then Err.Raise MAPI_E_NOT_FOUND, "stdOutlook", "No Inbox entry ID was returned."
  RequirePointer entry, "Inbox entry ID"
  CheckHR CallVTable(This.Store, VT_STORE_OPEN_ENTRY, cb, entry, CLngPtr(0), 0&, VarPtr(kind), VarPtr(folder)), "IMsgStore.OpenEntry(Inbox)"
  CheckHR CallVTable(folder, VT_CONTAINER_CONTENTS, 0&, VarPtr(table)), "IMAPIFolder.GetContentsTable"
  sort(0) = 1: sort(3) = &HE060040: sort(4) = 1 'PR_MESSAGE_DELIVERY_TIME descending
  CheckHR CallVTable(table, VT_TABLE_SORT, VarPtr(sort(0)), 0&), "IMAPITable.SortTable"
  Set Inbox = ReadTable(table, MessageColumns(False), maxResults, capped)
Done:
  On Error GoTo 0
  ReleaseObject table
  ReleaseObject folder
  FreeBuffer entry
  If number <> 0 Then Err.Raise number, "stdOutlook.Inbox", description
  Exit Function
Failed:
  number = Err.Number: description = Err.description
  Resume Done
End Function

'Entry ID must belong to the default store (e.g. from Inbox).
'Large/unavailable bodies remain Empty with the HRESULT in PropertyErrors.
'@param entryID - Nonempty, even-length hexadecimal MAPI entry ID in the current profile.
'@returns Collection - Message metadata and plain-text Body. Missing or oversized values are Empty with HRESULTs in PropertyErrors.
Public Function ReadMessage(ByVal entryID As String) As Collection
  Dim bytes() As Byte, message As LongPtr, kind As Long
  Dim number As Long, description As String
  EnsureStore
  bytes = DecodeEntryID(entryID)
  On Error GoTo Failed
  CheckHR CallVTable(This.Store, VT_STORE_OPEN_ENTRY, CLng(UBound(bytes) + 1), VarPtr(bytes(0)), CLngPtr(0), 0&, VarPtr(kind), VarPtr(message)), "IMsgStore.OpenEntry(message)"
  If kind <> 5 Then Err.Raise 5, "stdOutlook", "The entry is not a MAPI message."
  Set ReadMessage = GetProperties(message, MessageColumns(True))
Done:
  On Error GoTo 0
  ReleaseObject message
  If number <> 0 Then Err.Raise number, "stdOutlook.ReadMessage", description
  Exit Function
Failed:
  number = Err.Number: description = Err.description
  Resume Done
End Function

'Submit a plain-text message. Lists contain semicolon-separated bare SMTP
'addresses (not display names/aliases). Success means queued, not delivered.
'@param toAddresses - Semicolon-separated bare SMTP addresses. At least one To, CC, or BCC recipient is required.
'@param subject - Subject text; empty is allowed, embedded null characters are rejected.
'@param body - Plain-text Unicode body; empty is allowed, embedded null characters are rejected.
'@param ccAddresses - Optional semicolon-separated bare SMTP CC addresses.
'@param bccAddresses - Optional semicolon-separated bare SMTP BCC addresses.
'@param flushQueue - True requests synchronization after successful submission; False queues only. Defaults to True.
'@remark Success means submitted. Inspect LastSendSyncError; never resend a queued message after a sync error.
Public Sub SendEmail(ByVal toAddresses As String, ByVal subject As String, ByVal body As String, Optional ByVal ccAddresses As String = "", Optional ByVal bccAddresses As String = "", Optional ByVal flushQueue As Boolean = True)
  SendCore toAddresses, subject, body, ccAddresses, bccAddresses, False, flushQueue
End Sub

'Submit an HTML message. HTML is written as UTF-8 PR_HTML, with CPID 65001.
'No attachments, account override, or plain-text alternative is synthesized.
'@param toAddresses - Semicolon-separated bare SMTP addresses. At least one To, CC, or BCC recipient is required.
'@param subject - Subject text; empty is allowed, embedded null characters are rejected.
'@param htmlBody - HTML fragment or document stored as UTF-8 PR_HTML; no plain-text alternative is synthesized.
'@param ccAddresses - Optional semicolon-separated bare SMTP CC addresses.
'@param bccAddresses - Optional semicolon-separated bare SMTP BCC addresses.
'@param flushQueue - True requests synchronization after successful submission; False queues only. Defaults to True.
'@remark Success means submitted. Inspect LastSendSyncError; never resend a queued message after a sync error.
Public Sub SendEmailHTML(ByVal toAddresses As String, ByVal subject As String, ByVal htmlBody As String, Optional ByVal ccAddresses As String = "", Optional ByVal bccAddresses As String = "", Optional ByVal flushQueue As Boolean = True)
  SendCore toAddresses, subject, htmlBody, ccAddresses, bccAddresses, True, flushQueue
End Sub

'These describe only the latest SendEmail/SendEmailHTML call. A successful
'flush request is not a delivery receipt. False + empty error means skipped.
'@returns - True if the latest send synchronization request was accepted; not confirmation of delivery.
Public Property Get LastSendSyncRequested() As Boolean
  LastSendSyncRequested = This.LastSendSyncRequested
End Property

'Immediate synchronization error from the latest send, or an empty string; never a delivery receipt.
'@returns - Error description after a successful submission but failed synchronization; empty if skipped or no immediate error.
Public Property Get LastSendSyncError() As String
  LastSendSyncError = This.LastSendSyncError
End Property

'True by default: use Outlook's Send/Receive All after submission. False keeps
'the older native FlushQueues path, which some cached transports do not drain.
'The per-send flushQueue:=False option skips either synchronization path.
'@returns - True for Outlook automation, False for native queue flushing.
Public Property Get UseOutlookSendReceive() As Boolean
  UseOutlookSendReceive = This.UseOutlookSendReceive
End Property

'Select the automatic synchronization mechanism; True uses Outlook Send/Receive and is the default.
'@param value - True selects Outlook Send/Receive; False selects native FlushQueues. Defaults to True.
Public Property Let UseOutlookSendReceive(ByVal value As Boolean)
  This.UseOutlookSendReceive = value
End Property

'The documented equivalent of Outlook's Send/Receive All button. This is the
'only production path using Outlook automation; no guarded data or Send calls.
'Starts classic Outlook if needed, and leaves a window open so asynchronous
'transfer can continue after this class is released. Never calls Outlook.Quit.
'All accounts in that Outlook profile participate. Return is NOT delivery.
'@param showProgress - Whether to show Outlook Send/Receive progress; defaults to False.
Public Sub SendAndReceive(Optional ByVal showProgress As Boolean = False)
  Dim outlookApp As Object, session As Object, inbox As Object, explorer As Object
  EnsureSession
  EnsureStore
  Set outlookApp = RunningOutlook()
  If outlookApp Is Nothing Then Set outlookApp = CreateObject("Outlook.Application")
  Set session = outlookApp.GetNamespace("MAPI")
  'Initialize a default-profile session without Logon/profile-selection prompts.
  Set inbox = session.GetDefaultFolder(6) 'olFolderInbox
  CheckSyncStore session
  If outlookApp.Explorers.Count = 0 Then
    Set explorer = inbox.GetExplorer
    explorer.Display
  End If
  session.SendAndReceive showProgress
End Sub

'Return an existing Outlook automation instance, or Nothing when unavailable.
Private Function RunningOutlook() As Object
  On Error Resume Next
  Set RunningOutlook = GetObject(, "Outlook.Application")
  On Error GoTo 0
End Function

'Compare native and Outlook default-store entry IDs; refuse mismatched or unverifiable stores.
Private Sub CheckSyncStore(ByVal session As Object)
  Dim props As Collection, nativeID() As Byte, outlookID() As Byte, same As Long
  Set props = GetProperties(This.Store, Array(Array("EntryID", &HFFF0102)))
  nativeID = DecodeEntryID(CStr(props("EntryID")))
  outlookID = DecodeEntryID(CStr(session.DefaultStore.StoreID))
  CheckHR CallVTable(This.Session, VT_SESSION_COMPARE_IDS, CLng(UBound(nativeID) + 1), VarPtr(nativeID(0)), CLng(UBound(outlookID) + 1), VarPtr(outlookID(0)), 0&, VarPtr(same)), "IMAPISession.CompareEntryIDs(sync store)"
  If same = 0 Then Err.Raise 5, "stdOutlook.SendAndReceive", "Outlook is using a different default message store. Open Outlook with the same profile as this stdOutlook instance, then call SendAndReceive again; do not resubmit the message."
End Sub

'Request immediate transfer of pending mail in this profile, including mail
'queued by other callers. Not equivalent to synchronizing every Outlook folder.
'Returns the number of status objects that accepted the request; raises on
'unsupported providers or failure (possibly after some queues were flushed).
'Spooler requests may be asynchronous; direct transport calls can block.
'@param includeReceive - True flushes incoming and outgoing queues; False requests outgoing only. Defaults to False.
'@returns - Number of status objects accepting the request, not messages delivered; raises on unsupported providers or failure.
Public Function FlushQueues(Optional ByVal includeReceive As Boolean = False) As Long
  Dim table As LongPtr, rows As Collection, row As Collection, spooler As Collection
  Dim capped As Boolean, number As Long, description As String
  EnsureSession
  EnsureStore 'Keep the default store open while asking transports to send.
  On Error GoTo Failed
  CheckHR CallVTable(This.Session, VT_SESSION_STATUS_TABLE, 0&, VarPtr(table)), "IMAPISession.GetStatusTable"
  RequirePointer table, "GetStatusTable"
  Set rows = ReadTable(table, Array(Array("EntryID", &HFFF0102), Array("ResourceType", &H3E030003), Array("Methods", &H3E020003)), 1000, capped)
  ReleaseObject table
  If capped Then Err.Raise MAPI_E_TOO_BIG, "stdOutlook.FlushQueues", "Status table exceeds 1000 rows; no queues were flushed."
  'Prefer the spooler, which coordinates all transports. Do not flush both the
  'spooler and its individual transports for the same request.
  For Each row In rows
    If CLng(row("ResourceType")) = 37 And (CLng(row("Methods")) And 8) <> 0 Then
      Set spooler = row
      Exit For
    End If
  Next
  If Not spooler Is Nothing Then
    FlushStatus spooler, includeReceive
    FlushQueues = 1
  Else
    For Each row In rows
      If CLng(row("ResourceType")) = 36 And (CLng(row("Methods")) And 8) <> 0 Then
        FlushStatus row, includeReceive
        FlushQueues = FlushQueues + 1
      End If
    Next
  End If
  If FlushQueues = 0 Then Err.Raise MAPI_E_NO_SUPPORT, "stdOutlook.FlushQueues", "This profile exposes no transport supporting native FlushQueues; use Outlook Send/Receive."
Done:
  On Error GoTo 0
  ReleaseObject table
  If number <> 0 Then Err.Raise number, "stdOutlook.FlushQueues", description
  Exit Function
Failed:
  number = Err.Number: description = Err.description
  Resume Done
End Function

'Open and release one supported status object while requesting queue transfer.
Private Sub FlushStatus(ByVal row As Collection, ByVal includeReceive As Boolean)
  Dim bytes() As Byte, status As LongPtr, kind As Long, flags As Long
  Dim number As Long, description As String
  On Error GoTo Failed
  bytes = DecodeEntryID(CStr(row("EntryID")))
  CheckHR CallVTable(This.Session, VT_SESSION_OPEN_ENTRY, CLng(UBound(bytes) + 1), VarPtr(bytes(0)), CLngPtr(0), 0&, VarPtr(kind), VarPtr(status)), "IMAPISession.OpenEntry(status)"
  RequirePointer status, "OpenEntry(status)"
  If kind <> 10 Then Err.Raise 5, "stdOutlook.FlushQueues", "The transport entry is not a MAPI status object."
  flags = &H2 Or &H8 Or &H10 'FLUSH_UPLOAD | FLUSH_FORCE | FLUSH_NO_UI
  If includeReceive Then flags = flags Or &H4
  If CLng(row("ResourceType")) = 37 Then flags = flags Or &H20 'FLUSH_ASYNC_OK: spooler only
  CheckHR CallVTable(status, VT_STATUS_FLUSH, CLngPtr(0), 0&, CLngPtr(0), flags), "IMAPIStatus.FlushQueues (message already queued; do not resend)"
Done:
  On Error GoTo 0
  ReleaseObject status
  If number <> 0 Then Err.Raise number, "stdOutlook.FlushQueues", description
  Exit Sub
Failed:
  number = Err.Number: description = Err.description
  Resume Done
End Sub

'Record synchronization failure separately after successful submission; never resubmit.
Private Sub RequestSendSync()
  On Error GoTo Failed
  If This.UseOutlookSendReceive Then
    Me.SendAndReceive False
  Else
    Me.FlushQueues False
  End If
  This.LastSendSyncRequested = True
  Exit Sub
Failed:
  This.LastSendSyncError = Err.description
End Sub

'Validate recipients, create and submit one message, release native buffers, then request synchronization.
Private Sub SendCore(ByVal toAddresses As String, ByVal subject As String, ByVal body As String, ByVal ccAddresses As String, ByVal bccAddresses As String, ByVal isHtml As Boolean, ByVal flushQueue As Boolean)
  Dim recipients As New Collection, folders As Collection, outbox As LongPtr, message As LongPtr
  Dim bytes() As Byte, sentID As String, kind As Long, list As LongPtr
  Dim number As Long, description As String
  This.LastSendSyncError = "": This.LastSendSyncRequested = False
  'Validate everything before opening a writable store or creating a message.
  AddRecipients recipients, toAddresses, 1
  AddRecipients recipients, ccAddresses, 2
  AddRecipients recipients, bccAddresses, 3
  If recipients.Count = 0 Then Err.Raise 5, "stdOutlook.SendEmail", "At least one SMTP recipient is required."
  If InStr(subject, vbNullChar) > 0 Or InStr(body, vbNullChar) > 0 Then Err.Raise 5, "stdOutlook.SendEmail", "Subject and body cannot contain null characters."
  EnsureAddressBook
  EnsureStore True
  On Error GoTo Failed
  Set folders = GetProperties(This.Store, Array(Array("Outbox", &H35E20102), Array("Sent", &H35E40102)))
  bytes = DecodeEntryID(CStr(folders("Outbox")))
  sentID = CStr(folders("Sent"))
  If Len(sentID) = 0 Then Err.Raise MAPI_E_NOT_FOUND, "stdOutlook", "The default store has no Sent Items folder."
  list = BuildRecipientList(recipients)
  CheckHR CallVTable(This.Store, VT_STORE_OPEN_ENTRY, CLng(UBound(bytes) + 1), VarPtr(bytes(0)), CLngPtr(0), MAPI_MODIFY, VarPtr(kind), VarPtr(outbox)), "IMsgStore.OpenEntry(Outbox, write)"
  If kind <> 3 Then Err.Raise 5, "stdOutlook", "The Outbox entry is not a folder."
  CheckHR CallVTable(outbox, VT_FOLDER_CREATE_MESSAGE, CLngPtr(0), 0&, VarPtr(message)), "IMAPIFolder.CreateMessage"
  SetOutgoingProperties message, subject, body, sentID, isHtml
  CheckHR CallVTable(message, VT_MESSAGE_RECIPIENTS, 2&, list), "IMessage.ModifyRecipients(MODRECIP_ADD)"
  'SubmitMessage saves the message. Never retry automatically: a transport failure
  'can have an uncertain outcome. After success only Release is allowed.
  CheckHR CallVTable(message, VT_MESSAGE_SUBMIT, 0&), "IMessage.SubmitMessage (not retried; inspect Outbox before retrying)"
Done:
  On Error Resume Next
  ReleaseObject message
  ReleaseObject outbox
  FreeRows list 'ADRLIST and SRowSet have identical ownership/layout here.
  On Error GoTo 0
  If number <> 0 Then Err.Raise number, "stdOutlook.SendEmail", description
  'Submission has succeeded. A subsequent synchronization failure must never
  'look like a failed send: callers might otherwise create a duplicate message.
  If flushQueue Then RequestSendSync
  Exit Sub
Failed:
  number = Err.Number: description = Err.description
  Resume Done
End Sub

'Append validated SMTP recipients with their MAPI To/CC/BCC type; enforce the shared limit.
Private Sub AddRecipients(ByVal recipients As Collection, ByVal addresses As String, ByVal recipientType As Long)
  Dim parts As Variant, part As Variant, address As String, at As Long, i As Long, character As String
  If Len(Trim$(addresses)) = 0 Then Exit Sub
  parts = Split(addresses, ";")
  For Each part In parts
    address = Trim$(CStr(part))
    at = InStr(address, "@")
    If at <= 1 Or at = Len(address) Or InStr(at + 1, address, "@") > 0 Then GoTo Invalid
    For i = 1 To Len(address)
      character = Mid$(address, i, 1)
      If AscW(character) >= 0 And AscW(character) <= 32 Then GoTo Invalid
      If AscW(character) = 127 Then GoTo Invalid
      If InStr(",<>""()", character) > 0 Then GoTo Invalid
    Next
    If recipients.Count >= 1000 Then Err.Raise 5, "stdOutlook.SendEmail", "At most 1000 recipients are supported."
    recipients.Add Array(address, recipientType)
  Next
  Exit Sub
Invalid:
  Err.Raise 5, "stdOutlook.SendEmail", "Use semicolon-separated bare SMTP addresses with no empty entries, display names, or control characters."
End Sub

'Allocate an owned ADRLIST with one-off recipient IDs; caller frees it with FreeRows.
Private Function BuildRecipientList(ByVal recipients As Collection) As LongPtr
  Dim list As LongPtr, props As LongPtr, row As LongPtr, oneOff As LongPtr, cb As Long
  Dim entry As Variant, address As String, addressType As String, i As Long
  Dim number As Long, description As String, zeros() As Byte
  Dim propZeros(0 To PROP_BYTES * 5 - 1) As Byte
  On Error GoTo Failed
  ReDim zeros(0 To ROWSET_OFFSET + recipients.Count * ROW_BYTES - 1)
  CheckHR MAPIAllocateBuffer(UBound(zeros) + 1, list), "MAPIAllocateBuffer(ADRLIST)"
  CopyMemory ByVal list, zeros(0), UBound(zeros) + 1
  WriteLong list, recipients.Count
  addressType = "SMTP"
  For Each entry In recipients
    address = CStr(entry(0))
    row = PtrAdd(list, ROWSET_OFFSET + i * ROW_BYTES)
    CheckHR MAPIAllocateBuffer(5 * PROP_BYTES, props), "MAPIAllocateBuffer(recipient properties)"
    WritePointer PtrAdd(row, 8), props 'Transfer ownership before any fallible work.
    WriteLong PtrAdd(row, 4), 5
    CopyMemory ByVal props, propZeros(0), PROP_BYTES * 5
    WriteLong props, &HC150003 'PR_RECIPIENT_TYPE: To=1, CC=2, BCC=3
    WriteLong PtrAdd(props, 8), CLng(entry(1))
    AllocatedStringProperty PtrAdd(props, PROP_BYTES), &H3001001F, address, props
    AllocatedStringProperty PtrAdd(props, PROP_BYTES * 2), &H3002001F, addressType, props
    AllocatedStringProperty PtrAdd(props, PROP_BYTES * 3), &H3003001F, address, props
    CheckHR CallVTable(This.AddressBook, VT_AB_ONE_OFF, StrPtr(address), StrPtr(addressType), StrPtr(address), MAPI_UNICODE Or &H10000, VarPtr(cb), VarPtr(oneOff)), "IAddrBook.CreateOneOff(SMTP)"
    RequirePointer oneOff, "CreateOneOff"
    If cb <= 0 Then Err.Raise 5, "stdOutlook", "Empty one-off entry ID."
    Dim copy As LongPtr, binaryProp As LongPtr
    CheckHR MAPIAllocateMore(cb, props, copy), "MAPIAllocateMore(entry ID)"
    CopyMemory ByVal copy, ByVal oneOff, cb
    binaryProp = PtrAdd(props, PROP_BYTES * 4)
    WriteLong binaryProp, &HFFF0102
    WriteLong PtrAdd(binaryProp, 8), cb
    WritePointer PtrAdd(binaryProp, 8 + BINARY_POINTER_OFFSET), copy
    FreeBuffer oneOff
    i = i + 1
  Next
  BuildRecipientList = list
  list = 0
Done:
  On Error GoTo 0
  FreeBuffer oneOff
  FreeRows list
  If number <> 0 Then Err.Raise number, "stdOutlook.BuildRecipientList", description
  Exit Function
Failed:
  number = Err.Number: description = Err.description
  Resume Done
End Function

'Allocate a null-terminated Unicode string under a MAPI parent and populate its property value.
Private Sub AllocatedStringProperty(ByVal prop As LongPtr, ByVal tag As Long, ByVal value As String, ByVal parent As LongPtr)
  Dim text As LongPtr, cb As Long, terminator As Integer
  cb = LenB(value)
  CheckHR MAPIAllocateMore(cb + 2, parent, text), "MAPIAllocateMore(text)"
  If cb > 0 Then CopyMemory ByVal text, ByVal StrPtr(value), cb
  CopyMemory ByVal PtrAdd(text, cb), terminator, 2
  WriteLong prop, tag
  WritePointer PtrAdd(prop, 8), text
End Sub

'Populate message, body, recipient-independent delivery properties and reject per-property errors.
Private Sub SetOutgoingProperties(ByVal message As LongPtr, ByVal subject As String, ByVal body As String, ByVal sentID As String, ByVal isHtml As Boolean)
  Dim props(0 To PROP_BYTES * 6 - 1) As Byte, p As LongPtr, messageClass As String
  Dim sentBytes() As Byte, html() As Byte, htmlSize As Long, problems As LongPtr
  Dim number As Long, description As String
  On Error GoTo Failed
  messageClass = "IPM.Note"
  subject = subject & vbNullChar
  If Not isHtml Then body = body & vbNullChar
  sentBytes = DecodeEntryID(sentID)
  p = VarPtr(props(0))
  WriteLong p, &H1A001F: WritePointer PtrAdd(p, 8), StrPtr(messageClass)
  p = PtrAdd(p, PROP_BYTES)
  WriteLong p, &H37001F: WritePointer PtrAdd(p, 8), StrPtr(subject)
  p = PtrAdd(p, PROP_BYTES)
  If isHtml Then
    htmlSize = 0
    If Len(body) > 0 Then
      htmlSize = WideCharToMultiByte(65001, &H80&, StrPtr(body), Len(body), 0, 0, 0, 0)
      If htmlSize = 0 Then Err.Raise 5, "stdOutlook", "HTML contains invalid Unicode."
    End If
    ReDim html(0 To htmlSize)
    If htmlSize > 0 Then
      If WideCharToMultiByte(65001, &H80&, StrPtr(body), Len(body), VarPtr(html(0)), htmlSize, 0, 0) <> htmlSize Then Err.Raise 5, "stdOutlook", "Cannot encode HTML as UTF-8."
    End If
    WriteLong p, &H10130102: WriteLong PtrAdd(p, 8), htmlSize
    WritePointer PtrAdd(p, 8 + BINARY_POINTER_OFFSET), VarPtr(html(0))
  Else
    WriteLong p, &H1000001F: WritePointer PtrAdd(p, 8), StrPtr(body)
  End If
  p = PtrAdd(p, PROP_BYTES)
  WriteLong p, &HE0A0102: WriteLong PtrAdd(p, 8), UBound(sentBytes) + 1
  WritePointer PtrAdd(p, 8 + BINARY_POINTER_OFFSET), VarPtr(sentBytes(0))
  p = PtrAdd(p, PROP_BYTES)
  WriteLong p, &HE070003: WriteLong PtrAdd(p, 8), 8 'MSGFLAG_UNSENT
  p = PtrAdd(p, PROP_BYTES)
  WriteLong p, &H3FDE0003: WriteLong PtrAdd(p, 8), 65001 'PR_INTERNET_CPID
  CheckHR CallVTable(message, VT_SET_PROPS, 6&, VarPtr(props(0)), VarPtr(problems)), "IMAPIProp.SetProps(outgoing message)"
  If problems <> 0 Then
    If ReadLong(problems) > 0 Then
      'SPropProblemArray: count, then {index, tag, HRESULT} (all ULONGs).
      number = ReadLong(PtrAdd(problems, 12))
      If number >= 0 Then number = &H80004005
      description = "An outgoing message property was rejected (index " & ReadLong(PtrAdd(problems, 4)) & "). Message was not submitted."
    End If
  End If
Done:
  On Error GoTo 0
  FreeBuffer problems
  If number <> 0 Then Err.Raise number, "stdOutlook.SetOutgoingProperties", description
  Exit Sub
Failed:
  number = Err.Number: description = Err.description
  Resume Done
End Sub

'Search stored calendar items by case-insensitive subject/location/organizer
'substring. Empty query lists items. Recurring masters are NOT expanded.
'@param query - Search text; matching rules are described by this method.
'@param maxResults - Maximum returned records, from 1 to 10000.
'@returns Collection - Stored appointment records with Subject, Location, OrganizerName, StartUtc, EndUtc, IsRecurring, RecurrenceExpanded and PropertyErrors. Recurrences are not expanded.
'@remark Empty query lists appointments; nonempty text matches Subject, Location or OrganizerName. Stored masters only.
Public Function SearchCalendar(Optional ByVal query As String = "", Optional ByVal maxResults As Long = 100) As Collection
  Dim folder As LongPtr, table As LongPtr, columns As Variant, sort(0 To 4) As Long
  Dim results As Collection, record As Collection, errors As Collection, field As Variant
  Dim number As Long, description As String
  This.CalendarSearchTruncated = False
  ValidateLimit maxResults
  If InStr(query, vbNullChar) > 0 Then Err.Raise 5, "stdOutlook.SearchCalendar", "Query cannot contain null characters."
  EnsureStore
  On Error GoTo Failed
  folder = OpenCalendarFolder()
  columns = CalendarColumns(folder)
  CheckHR CallVTable(folder, VT_CONTAINER_CONTENTS, 0&, VarPtr(table)), "IMAPIFolder.GetContentsTable(calendar)"
  If Len(query) > 0 Then RestrictCalendar table, query, columns
  For Each field In columns
    If field(0) = "StartUtc" Then
      sort(0) = 1: sort(3) = CLng(field(1)): sort(4) = 1
      CheckHR CallVTable(table, VT_TABLE_SORT, VarPtr(sort(0)), 0&), "IMAPITable.SortTable(calendar)"
      Exit For
    End If
  Next
  Set results = ReadTable(table, columns, maxResults, This.CalendarSearchTruncated)
  For Each record In results
    Set errors = record("PropertyErrors")
    For Each field In Array("Location", "StartUtc", "EndUtc", "IsRecurring")
      If Not HasKey(record, CStr(field)) Then
        record.Add Empty, CStr(field)
        errors.Add MAPI_E_NOT_FOUND, CStr(field)
      End If
    Next
    record.Add False, "RecurrenceExpanded"
  Next
  Set SearchCalendar = results
Done:
  On Error GoTo 0
  ReleaseObject table
  ReleaseObject folder
  If number <> 0 Then Err.Raise number, "stdOutlook.SearchCalendar", description
  Exit Function
Failed:
  number = Err.Number: description = Err.description
  Resume Done
End Function

'Whether the latest calendar search found additional records beyond its result limit.
'@returns - True when more records exist than the requested maximum.
Public Property Get CalendarSearchTruncated() As Boolean
  CalendarSearchTruncated = This.CalendarSearchTruncated
End Property

'Resolve the default Calendar through Inbox then root properties; return an owned folder interface.
Private Function OpenCalendarFolder() As LongPtr
  Dim source As LongPtr, entry As LongPtr, cb As Long, kind As Long, record As Collection
  Dim messageClass As String, calendarID As String, bytes() As Byte, folder As LongPtr
  Dim number As Long, description As String
  On Error GoTo InboxFailed
  messageClass = "IPM.Note"
  CheckHR CallVTable(This.Store, VT_STORE_RECEIVE_FOLDER, StrPtr(messageClass), MAPI_UNICODE, VarPtr(cb), VarPtr(entry), CLngPtr(0)), "GetReceiveFolder(calendar lookup)"
  If cb <= 0 Then Err.Raise MAPI_E_NOT_FOUND
  RequirePointer entry, "Inbox entry ID"
  CheckHR CallVTable(This.Store, VT_STORE_OPEN_ENTRY, cb, entry, CLngPtr(0), 0&, VarPtr(kind), VarPtr(source)), "OpenEntry(Inbox for calendar)"
  Set record = GetProperties(source, Array(Array("Calendar", &H36D00102)))
  calendarID = CStr(record("Calendar"))
  If Len(calendarID) > 0 Then GoTo OpenCalendar
TryRoot:
  'The Calendar entry ID lives on Inbox or the store's root, not IMsgStore.
  On Error GoTo Failed
  FreeBuffer entry
  ReleaseObject source
  CheckHR CallVTable(This.Store, VT_STORE_OPEN_ENTRY, 0&, CLngPtr(0), CLngPtr(0), 0&, VarPtr(kind), VarPtr(source)), "OpenEntry(root for calendar)"
  Set record = GetProperties(source, Array(Array("Calendar", &H36D00102)))
  calendarID = CStr(record("Calendar"))
OpenCalendar:
  On Error GoTo Failed
  If Len(calendarID) = 0 Then Err.Raise MAPI_E_NOT_FOUND, "stdOutlook", "The default store has no Calendar folder."
  bytes = DecodeEntryID(calendarID)
  CheckHR CallVTable(This.Store, VT_STORE_OPEN_ENTRY, CLng(UBound(bytes) + 1), VarPtr(bytes(0)), CLngPtr(0), 0&, VarPtr(kind), VarPtr(folder)), "OpenEntry(Calendar)"
  If kind <> 3 Then Err.Raise 5, "stdOutlook", "The calendar entry is not a folder."
  OpenCalendarFolder = folder: folder = 0
Done:
  On Error GoTo 0
  FreeBuffer entry
  ReleaseObject source
  ReleaseObject folder
  If number <> 0 Then Err.Raise number, "stdOutlook.OpenCalendarFolder", description
  Exit Function
Failed:
  number = Err.Number: description = Err.description
  Resume Done
InboxFailed:
  Resume TryRoot
End Function

'Build calendar columns using store-specific appointment named-property IDs.
Private Function CalendarColumns(ByVal folder As LongPtr) As Variant
  Dim columns As Variant, fields As Variant, field As Variant, tag As Long, n As Long
  columns = Array(Array("EntryID", &HFFF0102), Array("Subject", &H37001F), _
    Array("OrganizerName", &H42001F), Array("MessageClass", &H1A001F))
  fields = Array(Array("Location", &H8208&, 31&), Array("StartUtc", &H820D&, 64&), _
    Array("EndUtc", &H820E&, 64&), Array("IsRecurring", &H8223&, 11&))
  For Each field In fields
    tag = AppointmentTag(folder, CLng(field(1)), CLng(field(2)))
    If tag <> 0 Then
      n = UBound(columns) + 1
      ReDim Preserve columns(0 To n)
      columns(n) = Array(field(0), tag)
    End If
  Next
  CalendarColumns = columns
End Function

'Resolve an appointment property without creating mappings; return zero when unavailable.
Private Function AppointmentTag(ByVal folder As LongPtr, ByVal lid As Long, ByVal propertyType As Long) As Long
  Dim guid(0 To 3) As Long, name(0 To PTR_BYTES * 3 - 1) As Byte
  Dim namePointer As LongPtr, tags As LongPtr, tag As Long, number As Long, description As String
  'PSETID_Appointment = {00062002-0000-0000-C000-000000000046}.
  guid(0) = &H62002: guid(2) = &HC0&: guid(3) = &H46000000
  WritePointer VarPtr(name(0)), VarPtr(guid(0))
  WriteLong PtrAdd(VarPtr(name(0)), PTR_BYTES), 0 'MNID_ID
  WriteLong PtrAdd(VarPtr(name(0)), PTR_BYTES * 2), lid
  namePointer = VarPtr(name(0))
  On Error GoTo Failed
  'Never MAPI_CREATE: searching must not change the named-property mapping.
  CheckHR CallVTable(folder, VT_NAMED_IDS, 1&, VarPtr(namePointer), 0&, VarPtr(tags)), "IMAPIProp.GetIDsFromNames(calendar)"
  RequirePointer tags, "GetIDsFromNames"
  If ReadLong(tags) <> 1 Then Err.Raise 5, "stdOutlook", "Unexpected named-property count."
  tag = ReadLong(PtrAdd(tags, 4))
  If (tag And &HFFFF&) <> 10 Then AppointmentTag = (tag And &HFFFF0000) Or propertyType
Done:
  On Error GoTo 0
  FreeBuffer tags
  If number <> 0 Then Err.Raise number, "stdOutlook.AppointmentTag", description
  Exit Function
Failed:
  number = Err.Number: description = Err.description
  Resume Done
End Function

'Apply a native case-insensitive substring OR restriction to available calendar text columns.
Private Sub RestrictCalendar(ByVal table As LongPtr, ByVal query As String, ByVal columns As Variant)
  Dim restriction(0 To RESTRICTION_BYTES - 1) As Byte
  Dim children(0 To RESTRICTION_BYTES * 3 - 1) As Byte, props(0 To PROP_BYTES * 3 - 1) As Byte
  Dim field As Variant, child As LongPtr, prop As LongPtr, count As Long
  For Each field In columns
    If field(0) = "Subject" Or field(0) = "Location" Or field(0) = "OrganizerName" Then
      child = PtrAdd(VarPtr(children(0)), count * RESTRICTION_BYTES)
      prop = PtrAdd(VarPtr(props(0)), count * PROP_BYTES)
      WriteLong child, 3 'RES_CONTENT
      WriteLong PtrAdd(child, RESTRICTION_UNION), &H10001 'FL_SUBSTRING | FL_IGNORECASE
      WriteLong PtrAdd(child, RESTRICTION_UNION + 4), CLng(field(1))
      WritePointer PtrAdd(child, RESTRICTION_UNION + 8), prop
      WriteLong prop, CLng(field(1))
      WritePointer PtrAdd(prop, 8), StrPtr(query)
      count = count + 1
    End If
  Next
  WriteLong VarPtr(restriction(0)), 1 'RES_OR
  WriteLong PtrAdd(VarPtr(restriction(0)), RESTRICTION_UNION), count
  WritePointer PtrAdd(VarPtr(restriction(0)), RESTRICTION_UNION + PTR_BYTES), VarPtr(children(0))
  CheckHR CallVTable(table, VT_TABLE_RESTRICT, VarPtr(restriction(0)), 0&), "IMAPITable.Restrict(calendar substring)"
End Sub

'Raise if the instance has no native session.
Private Sub EnsureSession()
  If This.Session = 0 Then Err.Raise 91, "stdOutlook", "Use stdOutlook.Create or OpenProfile before reading Outlook data."
End Sub

'Lazily open the address book with dialogs disabled.
Private Sub EnsureAddressBook()
  EnsureSession
  If This.AddressBook = 0 Then
    CheckHR CallVTable(This.Session, VT_SESSION_ADDRESS_BOOK, CLngPtr(0), CLngPtr(0), AB_NO_DIALOG, VarPtr(This.AddressBook)), "IMAPISession.OpenAddressBook"
    RequirePointer This.AddressBook, "OpenAddressBook"
  End If
End Sub

'Lazily open the default store; upgrade a read-only handle only for explicit writes.
Private Sub EnsureStore(Optional ByVal writable As Boolean = False)
  Dim table As LongPtr, rows As Collection, row As Collection, bytes() As Byte
  Dim capped As Boolean, number As Long, description As String, flags As Long
  EnsureSession
  If This.Store <> 0 Then
    If Not writable Or This.StoreWritable Then Exit Sub
    ReleaseObject This.Store
  End If
  This.StoreWritable = False
  flags = MDB_NO_DIALOG
  If writable Then flags = flags Or MDB_WRITE
  On Error GoTo Failed
  CheckHR CallVTable(This.Session, VT_SESSION_STORES, 0&, VarPtr(table)), "IMAPISession.GetMsgStoresTable"
  Set rows = ReadTable(table, Array(Array("EntryID", &HFFF0102), Array("Default", &H3400000B)), 1000, capped)
  For Each row In rows
    If row("Default") = True Then
      bytes = DecodeEntryID(CStr(row("EntryID")))
      CheckHR CallVTable(This.Session, VT_SESSION_OPEN_STORE, CLngPtr(0), CLng(UBound(bytes) + 1), VarPtr(bytes(0)), CLngPtr(0), flags, VarPtr(This.Store)), "IMAPISession.OpenMsgStore"
      RequirePointer This.Store, "OpenMsgStore"
      This.StoreWritable = writable
      Exit For
    End If
  Next
  If This.Store = 0 Then Err.Raise MAPI_E_NOT_FOUND, "stdOutlook", "The profile has no accessible default message store."
Done:
  On Error GoTo 0
  ReleaseObject table
  If number <> 0 Then Err.Raise number, "stdOutlook", description
  Exit Sub
Failed:
  number = Err.Number: description = Err.description
  ReleaseObject This.Store
  Resume Done
End Sub

'Search one address-book container, copying records and preserving partial-result diagnostics.
Private Function SearchBook(ByVal entryID As String, ByVal query As String, ByVal limit As Long, ByRef capped As Boolean) As Collection
  Dim bytes() As Byte, container As LongPtr, table As LongPtr, kind As Long
  Dim rows As Collection, row As Collection, person As Collection, result As New Collection
  Dim number As Long, description As String
  capped = False
  On Error GoTo Failed
  bytes = DecodeEntryID(entryID)
  CheckHR CallVTable(This.AddressBook, VT_AB_OPEN_ENTRY, CLng(UBound(bytes) + 1), VarPtr(bytes(0)), CLngPtr(0), 0&, VarPtr(kind), VarPtr(container)), "IAddrBook.OpenEntry(container)"
  CheckHR CallVTable(container, VT_CONTAINER_CONTENTS, 0&, VarPtr(table)), "IABContainer.GetContentsTable"
  RestrictName table, query
  Set rows = ReadTable(table, Array(Array("EntryID", &HFFF0102)), limit, capped)
  For Each row In rows
    Set person = ReadContactForSearch(CStr(row("EntryID")))
    If Not person Is Nothing Then result.Add person
  Next
Done:
  On Error GoTo 0
  ReleaseObject table
  ReleaseObject container
  If number <> 0 Then This.Warnings.Add "Address book " & entryID & ": " & description
  Set SearchBook = result
  Exit Function
Failed:
  number = Err.Number: description = Err.description
  Resume Done
End Function

'Read a search result or record its provider failure as a warning.
Private Function ReadContactForSearch(ByVal entryID As String) As Collection
  On Error GoTo Failed
  Set ReadContactForSearch = ReadContact(entryID)
  Exit Function
Failed:
  This.Warnings.Add "Contact " & entryID & ": " & Err.description
End Function

'Apply provider-side ambiguous-name resolution, retrying ANSI for unsupported Unicode.
Private Sub RestrictName(ByVal table As LongPtr, ByVal query As String)
  Dim restriction(0 To RESTRICTION_BYTES - 1) As Byte
  Dim prop(0 To PROP_BYTES - 1) As Byte
  Dim hr As Long, ansi() As Byte
  'RES_PROPERTY { RELOP_EQ, PR_ANR_W, &SPropValue }. Explicit native padding.
  WriteLong VarPtr(restriction(0)), 4
  WriteLong PtrAdd(VarPtr(restriction(0)), RESTRICTION_UNION), 4
  WriteLong PtrAdd(VarPtr(restriction(0)), RESTRICTION_UNION + 4), &H360C001F
  WritePointer PtrAdd(VarPtr(restriction(0)), RESTRICTION_UNION + 8), VarPtr(prop(0))
  WriteLong VarPtr(prop(0)), &H360C001F
  WritePointer PtrAdd(VarPtr(prop(0)), 8), StrPtr(query)
  hr = CallVTable(table, VT_TABLE_RESTRICT, VarPtr(restriction(0)), 0&)
  If hr = MAPI_E_BAD_CHARWIDTH Then
    ansi = StrConv(query & vbNullChar, vbFromUnicode)
    WriteLong PtrAdd(VarPtr(restriction(0)), RESTRICTION_UNION + 4), &H360C001E
    WriteLong VarPtr(prop(0)), &H360C001E
    WritePointer PtrAdd(VarPtr(prop(0)), 8), VarPtr(ansi(0))
    hr = CallVTable(table, VT_TABLE_RESTRICT, VarPtr(restriction(0)), 0&)
  End If
  CheckHR hr, "IMAPITable.Restrict(PR_ANR)"
End Sub

'Copy person fields and resolve primary SMTP from direct, SMTP-type or uppercase proxy values.
Private Function ReadPerson(ByVal person As LongPtr) As Collection
  Dim result As Collection, email As String, address As Variant, proxies As Variant
  Set result = GetProperties(person, PersonColumns())
  email = CStr(result("SmtpAddress"))
  If Len(email) = 0 And StrComp(CStr(result("AddressType")), "SMTP", vbTextCompare) = 0 Then email = CStr(result("EmailAddress"))
  If Len(email) = 0 Then
    proxies = result("ProxyAddresses")
    If IsArray(proxies) Then
      For Each address In proxies
        'Exchange marks the PRIMARY address with uppercase SMTP:, aliases lowercase.
        If StrComp(Left$(CStr(address), 5), "SMTP:", vbBinaryCompare) = 0 Then
          email = Mid$(CStr(address), 6)
          Exit For
        End If
      Next
    End If
  End If
  result.Add email, "Email"
  Set ReadPerson = result
End Function

'Return name/property-tag pairs for copied contact fields.
Private Function PersonColumns() As Variant
  PersonColumns = Array( _
    Array("EntryID", &HFFF0102), Array("DisplayName", &H3001001F), _
    Array("AddressType", &H3002001F), Array("EmailAddress", &H3003001F), _
    Array("SmtpAddress", &H39FE001F), Array("GivenName", &H3A06001F), _
    Array("Surname", &H3A11001F), Array("Company", &H3A16001F), _
    Array("Department", &H3A18001F), Array("JobTitle", &H3A17001F), _
    Array("BusinessPhone", &H3A08001F), Array("MobilePhone", &H3A1C001F), _
    Array("HomePhone", &H3A09001F), Array("Office", &H3A19001F), _
    Array("ObjectType", &HFFE0003), Array("ProxyAddresses", &H800F101F))
End Function

'Return name/property-tag pairs for address-book hierarchy rows.
Private Function BookColumns() As Variant
  BookColumns = Array(Array("EntryID", &HFFF0102), Array("DisplayName", &H3001001F), Array("Flags", &H36000003), Array("Depth", &H30050003))
End Function

'Return message metadata columns, optionally including the plain-text body.
Private Function MessageColumns(ByVal includeBody As Boolean) As Variant
  Dim result As Variant
  result = Array(Array("EntryID", &HFFF0102), Array("Subject", &H37001F), _
    Array("SenderName", &HC1A001F), Array("SenderAddressType", &HC1E001F), _
    Array("SenderEmailAddress", &HC1F001F), Array("SenderSmtpAddress", &H5D01001F), _
    Array("ReceivedTimeUtc", &HE060040), Array("MessageClass", &H1A001F), _
    Array("MessageFlags", &HE070003), Array("HasAttachments", &HE1B000B))
  If includeBody Then
    ReDim Preserve result(0 To 10)
    result(10) = Array("Body", &H1000001F)
  End If
  MessageColumns = result
End Function

'Copy requested native properties and always free the returned MAPI buffer.
Private Function GetProperties(ByVal instance As LongPtr, ByVal columns As Variant) As Collection
  Dim tags() As Long, count As Long, props As LongPtr
  Dim number As Long, description As String, hr As Long
  BuildTags columns, tags, True
  On Error GoTo Failed
  hr = CallVTable(instance, VT_GET_PROPS, VarPtr(tags(0)), MAPI_UNICODE, VarPtr(count), VarPtr(props))
  If hr = MAPI_E_BAD_CHARWIDTH Then
    FreeBuffer props
    hr = CallVTable(instance, VT_GET_PROPS, VarPtr(tags(0)), 0&, VarPtr(count), VarPtr(props))
  End If
  CheckHR hr, "IMAPIProp.GetProps"
  'Positive MAPI_W_ERRORS_RETURNED is partial success. Decode PT_ERROR per field.
  Set GetProperties = DecodeProperties(props, count, columns)
Done:
  On Error GoTo 0
  FreeBuffer props
  If number <> 0 Then Err.Raise number, "stdOutlook.GetProperties", description
  Exit Function
Failed:
  number = Err.Number: description = Err.description
  Resume Done
End Function

'Build an SPropTagArray from name/tag pairs, optionally allowing either string width.
Private Sub BuildTags(ByVal columns As Variant, ByRef tags() As Long, Optional ByVal unspecifiedStrings As Boolean = False)
  Dim i As Long, tag As Long
  ReDim tags(0 To UBound(columns) + 1)
  tags(0) = UBound(columns) + 1
  For i = 0 To UBound(columns)
    tag = CLng(columns(i)(1))
    If unspecifiedStrings Then
      If (tag And &HFFF&) = 31 Then tag = tag And &HFFFF0000 'PT_UNSPECIFIED; native A or W
    End If
    tags(i + 1) = tag
  Next
End Sub

'Copy paginated rows, continuing after short pages and reading one extra row to detect truncation.
Private Function ReadTable(ByVal table As LongPtr, ByVal columns As Variant, ByVal limit As Long, ByRef capped As Boolean) As Collection
  Dim tags() As Long, rows As LongPtr, count As Long, i As Long, row As LongPtr
  Dim take As Long, result As New Collection, item As Collection
  Dim number As Long, description As String, hr As Long
  capped = False
  BuildTags columns, tags
  On Error GoTo Failed
  hr = CallVTable(table, VT_TABLE_COLUMNS, VarPtr(tags(0)), 0&)
  If hr = MAPI_E_BAD_CHARWIDTH Then
    For i = 1 To UBound(tags)
      If (tags(i) And &HFFFF&) = 31 Then tags(i) = tags(i) - 1
    Next
    hr = CallVTable(table, VT_TABLE_COLUMNS, VarPtr(tags(0)), 0&)
  End If
  CheckHR hr, "IMAPITable.SetColumns"
  Do
    take = limit - result.Count + 1
    If take > BATCH_SIZE Then take = BATCH_SIZE
    CheckHR CallVTable(table, VT_TABLE_ROWS, take, 0&, VarPtr(rows)), "IMAPITable.QueryRows"
    RequirePointer rows, "QueryRows"
    count = ReadLong(rows)
    If count < 0 Or count > take Then Err.Raise 5, "stdOutlook", "Invalid native row count."
    For i = 0 To count - 1
      If result.Count = limit Then
        capped = True
        Exit For
      End If
      row = PtrAdd(rows, ROWSET_OFFSET + i * ROW_BYTES)
      Set item = DecodeProperties(ReadPointer(PtrAdd(row, 8)), ReadLong(PtrAdd(row, 4)), columns)
      result.Add item
    Next
    FreeRows rows
    If count = 0 Or capped Then Exit Do
    'A short page is NOT necessarily EOF. Keep going until QueryRows returns 0.
  Loop
  Set ReadTable = result
Done:
  On Error GoTo 0
  FreeRows rows
  If number <> 0 Then Err.Raise number, "stdOutlook.ReadTable", description
  Exit Function
Failed:
  number = Err.Number: description = Err.description
  Resume Done
End Function

'Map native properties into keyed values and a separate keyed PropertyErrors collection.
Private Function DecodeProperties(ByVal props As LongPtr, ByVal count As Long, ByVal columns As Variant) As Collection
  Dim result As New Collection, errors As New Collection
  Dim i As Long, prop As LongPtr, tag As Long, value As Variant
  If count <> UBound(columns) + 1 Then Err.Raise 5, "stdOutlook", "Provider returned an unexpected property count."
  RequirePointer props, "property array"
  For i = 0 To count - 1
    prop = PtrAdd(props, i * PROP_BYTES)
    tag = ReadLong(prop)
    value = Empty
    If (tag And &HFFFF&) = 10 Then
      errors.Add ReadLong(PtrAdd(prop, 8)), CStr(columns(i)(0))
    Else
      value = DecodeValue(prop)
    End If
    result.Add value, CStr(columns(i)(0))
  Next
  result.Add errors, "PropertyErrors"
  Set DecodeProperties = result
End Function

'Decode a supported SPropValue into a detached VBA value; timestamps are UTC.
Private Function DecodeValue(ByVal prop As LongPtr) As Variant
  Dim value As LongPtr, p As LongPtr, count As Long, i As Long
  Dim small As Integer, stamp As SYSTEMTIME, strings() As String, tag As Long
  tag = ReadLong(prop) And &HFFFF&
  value = PtrAdd(prop, 8)
  Select Case tag
    Case 0, 1, 10 'PT_UNSPECIFIED, PT_NULL, PT_ERROR
      DecodeValue = Empty
    Case 2
      CopyMemory small, ByVal value, 2
      DecodeValue = small
    Case 3
      DecodeValue = ReadLong(value)
    Case 11
      CopyMemory small, ByVal value, 2
      DecodeValue = (small <> 0)
    Case 30, 31
      DecodeValue = ReadString(ReadPointer(value), tag = 31)
    Case 64
      If FileTimeToSystemTime(value, stamp) <> 0 Then
        If stamp.year >= 100 And stamp.year <= 9999 Then
          DecodeValue = DateSerial(stamp.year, stamp.month, stamp.day) + TimeSerial(stamp.hour, stamp.minute, stamp.second)
        End If
      End If
    Case 258 'PT_BINARY, returned as a portable hex entry ID / blob
      count = ReadLong(value)
      p = ReadPointer(PtrAdd(value, BINARY_POINTER_OFFSET))
      DecodeValue = HexBytes(p, count)
    Case &H101E, &H101F 'PT_MV_STRING8 / PT_MV_UNICODE (proxy addresses)
      count = ReadLong(value)
      If count < 0 Or count > 100000 Then Err.Raise 5, "stdOutlook", "Invalid string-array size."
      If count = 0 Then
        DecodeValue = Array()
      Else
        p = ReadPointer(PtrAdd(value, BINARY_POINTER_OFFSET))
        RequirePointer p, "string array"
        ReDim strings(0 To count - 1)
        For i = 0 To count - 1
          strings(i) = ReadString(ReadPointer(PtrAdd(p, i * PTR_BYTES)), tag = &H101F)
        Next
        DecodeValue = strings
      End If
    Case Else
      Err.Raise MAPI_E_NO_SUPPORT, "stdOutlook", "Unsupported native property type 0x" & Hex$(tag)
  End Select
End Function

'DispCallFunc performs the ABI bridge. No generated machine code or RWX pages.
'ParamArray values MUST be typed: Long for ULONG/HRESULT, LongPtr for pointers.
'The array passed to DispCallFunc contains pointers to VARIANTs, not their data.
Private Function CallVTable(ByVal instance As LongPtr, ByVal slot As Long, ParamArray arguments() As Variant) As Long
  Dim values() As Variant, types() As Integer, pointers() As LongPtr
  Dim count As Long, i As Long, result As Variant, hr As Long
  RequirePointer instance, "native interface"
  values = arguments
  count = UBound(values) + 1
  ReDim types(0 To count): ReDim pointers(0 To count)
  For i = 0 To count - 1
    types(i) = VarType(values(i))
    pointers(i) = VarPtr(values(i))
  Next
  hr = DispCallFunc(instance, CLngPtr(slot * PTR_BYTES), 4, vbLong, count, VarPtr(types(0)), VarPtr(pointers(0)), result)
  CheckHR hr, "DispCallFunc"
  CallVTable = CLng(result)
End Function

'Release one owned IUnknown pointer and clear it to prevent repeated release.
Private Sub ReleaseObject(ByRef instance As LongPtr)
  Dim owned As LongPtr
  If instance = 0 Then Exit Sub
  owned = instance: instance = 0
  CallVTable owned, VT_RELEASE
End Sub

'Free one MAPI allocation and clear its pointer.
Private Sub FreeBuffer(ByRef buffer As LongPtr)
  If buffer = 0 Then Exit Sub
  MAPIFreeBuffer buffer
  buffer = 0
End Sub

'SRowSet owns separate property buffers per row. MAPIFreeBuffer(rows) alone
'leaks them. This is the documented FreeProws algorithm, without another DLL.
Private Sub FreeRows(ByRef rows As LongPtr)
  Dim i As Long, props As LongPtr
  If rows = 0 Then Exit Sub
  For i = 0 To ReadLong(rows) - 1
    props = ReadPointer(PtrAdd(rows, ROWSET_OFFSET + i * ROW_BYTES + 8))
    FreeBuffer props
  Next
  FreeBuffer rows
End Sub

'Add a byte offset while preserving unsigned 32-bit pointer arithmetic.
Private Function PtrAdd(ByVal pointer As LongPtr, ByVal offset As Long) As LongPtr
#If Win64 Then
  PtrAdd = pointer + offset
#Else
  'Handle 32-bit addresses crossing the signed Long boundary.
  Dim address As Double
  address = CDbl(pointer)
  If address < 0 Then address = address + 4294967296#
  address = address + offset
  If address >= 2147483648# Then address = address - 4294967296#
  PtrAdd = CLng(address)
#End If
End Function

'Copy a native four-byte signed integer.
Private Function ReadLong(ByVal pointer As LongPtr) As Long
  CopyMemory ReadLong, ByVal pointer, 4
End Function

'Copy a native pointer using the host pointer width.
Private Function ReadPointer(ByVal pointer As LongPtr) As LongPtr
  CopyMemory ReadPointer, ByVal pointer, PTR_BYTES
End Function

'Write a native four-byte integer.
Private Sub WriteLong(ByVal pointer As LongPtr, ByVal value As Long)
  CopyMemory ByVal pointer, value, 4
End Sub

'Write a native pointer using the host pointer width.
Private Sub WritePointer(ByVal pointer As LongPtr, ByVal value As LongPtr)
  CopyMemory ByVal pointer, value, PTR_BYTES
End Sub

'Copy a null-terminated native Unicode or system-code-page string.
Private Function ReadString(ByVal pointer As LongPtr, ByVal unicode As Boolean) As String
  Dim length As Long, characters As Long
  If pointer = 0 Then Exit Function
  If unicode Then
    length = lstrlenW(pointer)
    If length > 0 Then
      ReadString = String$(length, vbNullChar)
      CopyMemory ByVal StrPtr(ReadString), ByVal pointer, CLngPtr(length) * 2
    End If
  Else
    length = lstrlenA(pointer)
    If length > 0 Then
      characters = MultiByteToWideChar(0, 0, pointer, length, 0, 0)
      If characters = 0 Then Err.Raise 5, "stdOutlook", "Cannot decode provider ANSI text."
      ReadString = String$(characters, vbNullChar)
      If MultiByteToWideChar(0, 0, pointer, length, StrPtr(ReadString), characters) = 0 Then Err.Raise 5, "stdOutlook", "Cannot decode provider ANSI text."
    End If
  End If
End Function

'Copy native binary bytes to an uppercase hexadecimal string.
Private Function HexBytes(ByVal pointer As LongPtr, ByVal length As Long) As String
  Dim bytes() As Byte, i As Long
  If length < 0 Or length > 16777216 Then Err.Raise 5, "stdOutlook", "Invalid binary property length."
  If length = 0 Then Exit Function
  RequirePointer pointer, "binary property"
  ReDim bytes(0 To length - 1)
  CopyMemory bytes(0), ByVal pointer, length
  HexBytes = String$(length * 2, "0")
  For i = 0 To length - 1
    Mid$(HexBytes, i * 2 + 1, 2) = Right$("0" & Hex$(bytes(i)), 2)
  Next
End Function

'Validate a nonempty even-length hex ID and return a zero-based Byte array.
Private Function DecodeEntryID(ByVal value As String) As Byte()
  Dim bytes() As Byte, i As Long
  If Len(value) = 0 Or Len(value) Mod 2 <> 0 Or Len(value) > 131072 Then
    Err.Raise 5, "stdOutlook", "EntryID must be a nonempty, even-length hexadecimal string."
  End If
  ReDim bytes(0 To Len(value) \ 2 - 1)
  For i = 1 To Len(value)
    If InStr(1, "0123456789ABCDEF", Mid$(value, i, 1), vbTextCompare) = 0 Then Err.Raise 5, "stdOutlook", "EntryID contains a non-hexadecimal character."
  Next
  For i = 0 To UBound(bytes)
    bytes(i) = CByte("&H" & Mid$(value, i * 2 + 1, 2))
  Next
  DecodeEntryID = bytes
End Function

'Check membership in a VBA Collection without exposing a missing-key error.
Private Function HasKey(ByVal items As Collection, ByVal key As String) As Boolean
  Dim value As Variant
  On Error Resume Next
  value = items(key)
  HasKey = (Err.Number = 0)
  On Error GoTo 0
End Function

'Reject public result limits outside 1 through 10000.
Private Sub ValidateLimit(ByVal limit As Long)
  If limit < 1 Or limit > 10000 Then Err.Raise 5, "stdOutlook", "maxResults must be between 1 and 10000."
End Sub

'Reject a null interface or buffer returned by a native operation.
Private Sub RequirePointer(ByVal pointer As LongPtr, ByVal operation As String)
  If pointer = 0 Then Err.Raise 91, "stdOutlook", operation & " returned a null pointer."
End Sub

'Raise on negative HRESULTs while retaining the native operation context.
Private Sub CheckHR(ByVal hr As Long, ByVal operation As String)
  If hr < 0 Then Err.Raise hr, "stdOutlook", operation & " failed (HRESULT 0x" & Right$("00000000" & Hex$(hr), 8) & ")."
End Sub
