# stdOutlook

Read Outlook's current profile identity, address-book contacts, messages, and
calendar, and send plain-text or HTML email from one VBA class, without
Redemption or an Outlook reference. Data access and message submission use native
MAPI; automatic synchronization uses late-bound `Outlook.Application`.

Import **[src/stdOutlook.cls](../src/stdOutlook.cls)**. Everything needed at runtime is
inside that class; the accompanying Outlook documentation and tests are not runtime dependencies.

## How it works

The timed address-access prompts are associated with Outlook's **Object Model
Guard**, rather than a general five-minute COM lifetime. This class calls the
documented **Extended MAPI** interfaces underneath classic Outlook for reading
data and submitting messages. Synchronization uses `NameSpace.SendAndReceive`,
which is not listed among Microsoft's
[Object Model Guard protected members](https://learn.microsoft.com/en-us/office/vba/outlook/how-to/security/protected-properties-and-methods).
The class does not change security settings or use guarded address/body access
or `MailItem.Send`. Outlook startup/sign-in can still require user interaction.

`MAPIInitialize` and `MAPILogonEx` open the installed MAPI provider. VBA calls its
native interfaces through `oleaut32!DispCallFunc`. This follows the native-call
approach used by [stdVBA's stdCOM](https://github.com/sancarn/stdVBA/blob/master/src/stdCOM.cls),
without requiring stdCOM or executable-memory thunks.

MAPI uses COM-style `IUnknown` vtables internally. Data access is independent of
Outlook automation, not literally free of COM technology. It uses the existing
profile and its permissions; it does not extract credentials or parse OST/PST
files. Providers can access the server as well as cached data.

## Requirements

- Windows, **classic Outlook**, and a configured profile. New Outlook has no
  Extended MAPI support.
- A VBA7 host, such as Excel 2010 or later, with the **same bitness as Outlook**.
- Classic Outlook registered as the MAPI client used by Windows's `mapi32.dll`.
- A profile already able to sign in. Native MAPI logon and address-book dialogs
  are disabled; missing credentials, permissions, or profile setup produce VBA
  errors. Automatic synchronization also requires working Outlook automation.

64-bit Excel/classic Outlook has been exercised locally. The 32-bit branch has
layout checks but has **not yet been run in 32-bit Office**. Treat that platform
as experimental until tested there. Mac VBA, VBA6, and new Outlook are unsupported.

## Installation and first call

In the VBA editor, use **File > Import File** to import `src/stdOutlook.cls`.
Importing preserves the predeclared instance required by `stdOutlook.Create`;
pasting into a blank class does not.

```vb
Sub WhoAmI()
    Dim outlook As stdOutlook
    Set outlook = stdOutlook.Create()
    Debug.Print outlook.CurrentUserEmail
    outlook.Disconnect
End Sub
```

Keep that instance for subsequent calls. `Create()` shares the existing MAPI
session when one is available, otherwise uses the default profile. For a specific
profile, use `stdOutlook.Create("Profile name")`; this requests a new session to
ensure the profile name is honored. A profile is not an account display name.
Sessions are mail-capable (the `MAPI_NO_MAIL` flag is not used), so MAPI can start
the profile's transport/spooler. Store handles are opened read-only initially;
the first send requests a writable handle. Only the send methods submit messages.

`Disconnect` is idempotent and releases the native objects and MAPI initialization.
It does not log Outlook out of a shared session. Normal class termination also
cleans up. Avoid VBA's `End` statement or the VBE Reset button while the instance
is live: those bypass normal class cleanup. Use normal error handlers and release
the instance before editing/recompiling the project.

## Find contact details

```vb
Sub FindPeople()
    Dim outlook As stdOutlook, people As Collection, person As Collection
    Dim warning As Variant, number As Long, description As String
    On Error GoTo Failed
    Set outlook = stdOutlook.Create()
    Set people = outlook.SearchContacts("Ada", maxResults:=25)
    For Each person In people
        Debug.Print person("DisplayName"), person("Email"), _
                    person("BusinessPhone"), person("MobilePhone")
    Next
    If outlook.SearchTruncated Then Debug.Print "Result limit reached. Narrow the search."
    For Each warning In outlook.SearchWarnings
        Debug.Print warning
    Next
Done:
    On Error GoTo 0
    If Not outlook Is Nothing Then outlook.Disconnect
    If number <> 0 Then Err.Raise number, "FindPeople", description
    Exit Sub
Failed:
    number = Err.Number: description = Err.Description
    Resume Done
End Sub
```

Search uses the address-book provider's **ambiguous name resolution (ANR)**, usually
matching name or alias prefixes. Matching full SMTP addresses depends on the
provider. This is not an arbitrary substring, phone-number, or message-body search.
The provider filters results before they are fetched; the class does not download
the whole directory to search it in VBA.

`AddressBooks()` lists containers, including nested books. To target one, pass its
hexadecimal `EntryID` as the third argument to `SearchContacts`. Searching all books
skips containers without recipient contents and deduplicates identical entry IDs;
the same person with different entry IDs can still appear more than once.

Search covers the directory and Contacts folders **exposed by Outlook's address-book
providers**. A Contacts folder not enabled as an Outlook Address Book, autocomplete
recipients, and addresses occurring only in mail history are outside this API.
Distribution lists and other addressable directory objects can also match;
inspect `ObjectType` (`6` = mail user, `8` = distribution list).

Always inspect `SearchWarnings` and `SearchTruncated`. A failed provider or contact
read is reported as a warning while other books continue. A zero-match result with
warnings is not a complete negative search. Truncation is conservative when the
global result cap is reached. The maximum result limit is 10,000; default is 100.
Calls are synchronous, and a slow provider can block VBA.

## Returned records

Records are native VBA `Collection` objects with string keys. No scripting runtime
or additional COM library is needed. Data is copied out before native buffers are
released, so records remain valid after `Disconnect`.

`CurrentUser()` and `ReadContact(entryID)` return these keys:

- `EntryID`, `DisplayName`, `GivenName`, `Surname`, `ObjectType`.
- `Email`, `SmtpAddress`, `AddressType`, `EmailAddress`, `ProxyAddresses`.
- `Company`, `Department`, `JobTitle`, `Office`.
- `BusinessPhone`, `MobilePhone`, `HomePhone`.
- `PropertyErrors`, a keyed Collection of HRESULTs for unavailable properties.

Search results have the same keys plus `AddressBook`. Unsupported or unavailable
properties have the value `Empty`; `PropertyErrors("BusinessPhone")`, for example,
contains the provider's HRESULT if that field failed. `ProxyAddresses`, when
available, is a zero-based string array. Binary entry IDs are hexadecimal strings.

`Email` prefers `PR_SMTP_ADDRESS`, then an address whose type is SMTP, then the
uppercase `SMTP:` primary Exchange proxy. An Exchange `/o=...` address is preserved
in `EmailAddress` but is never presented as an SMTP address. If no primary SMTP
address is available, `Email` is an empty string and `CurrentUserEmail` raises
`MAPI_E_NOT_FOUND` instead of inventing one.

The current user is the profile's **primary MAPI identity**, not a list of all
configured accounts. Some PST/IMAP profiles do not implement `QueryIdentity`.
Those providers raise their native error; no registry or account-name guess is
used as a substitute.

`AddressBooks()` records contain `EntryID`, `DisplayName`, `Flags`, `Depth`, and
`PropertyErrors`. The hierarchy is bounded at 10,000 containers and fails explicitly
if that bound is exceeded.

## Optional message reads

```vb
Sub ReadRecentMail()
    Dim outlook As stdOutlook, headers As Collection, header As Collection
    Dim message As Collection
    Set outlook = stdOutlook.Create()
    Set headers = outlook.Inbox(10)
    For Each header In headers
        Debug.Print header("Subject"), header("ReceivedTimeUtc")
    Next
    If headers.Count > 0 Then
        Set header = headers(1)
        Set message = outlook.ReadMessage(header("EntryID"))
        Debug.Print message("Body")
    End If
    outlook.Disconnect
End Sub
```

`Inbox` reads the default message store's Inbox, newest first. Records contain
`EntryID`, `Subject`, `SenderName`, `SenderAddressType`, `SenderEmailAddress`,
`SenderSmtpAddress`, `ReceivedTimeUtc`, `MessageClass`, `MessageFlags`,
`HasAttachments`, and `PropertyErrors`. Timestamps are UTC VBA dates. Raw sender
addresses can be Exchange distinguished names; use `SenderSmtpAddress` for SMTP
when supplied. `ReadMessage` adds the plain-text `Body` and accepts an entry ID from
that same default store.

These reads do not mark messages as read. Editing existing messages, attachments,
HTML/RTF reading, arbitrary folders, secondary stores, and full-text mail search
are not implemented.
Large bodies that require `IStream` can return `Empty` with an error in
`PropertyErrors`; streaming large bodies is not implemented. There is no promise
that a provider will have every requested field cached or available offline.

## Send plain-text or HTML email

```vb
Sub SendExamples()
    Dim outlook As stdOutlook
    Set outlook = stdOutlook.Create()

    'Each call immediately submits a message; there is no compose dialog.
    outlook.SendEmail "alice@example.com;bob@example.com", _
                      "Status update", "The report is ready.", _
                      ccAddresses:="manager@example.com"
    If Len(outlook.LastSendSyncError) > 0 Then
        Debug.Print "Submitted, but synchronization failed: " & outlook.LastSendSyncError
        'Do not send the same message again. It is already queued.
    End If

    outlook.SendEmailHTML "alice@example.com", "Status update", _
                          "<h1>Report ready</h1><p>Please review the <b>summary</b>.</p>", _
                          bccAddresses:="archive@example.com"

    outlook.Disconnect
End Sub
```

The signatures are:

```vb
SendEmail(toAddresses As String, subject As String, body As String, _
          Optional ccAddresses As String = "", Optional bccAddresses As String = "", _
          Optional flushQueue As Boolean = True)
SendEmailHTML(toAddresses As String, subject As String, htmlBody As String, _
              Optional ccAddresses As String = "", Optional bccAddresses As String = "", _
              Optional flushQueue As Boolean = True)
```

Both are Subs. They submit through the default message store's transport and
request a copy in that store's Sent Items folder. They do not choose a secondary
account or override the sender. The profile must have an Outbox, Sent Items,
write access, and a working transport. If a profile was previously shared by a
client using `MAPI_NO_MAIL`, reconnect with an explicit profile name to obtain a
fresh mail-capable session.

Use semicolon-separated **bare SMTP addresses**. Display-name forms such as
`Alice <alice@example.com>`, aliases, comma-separated lists, quoted local parts,
and empty entries are rejected. At least one To, CC, or BCC recipient is required;
BCC-only messages are supported. The limit is 1,000 recipients across all lists.
Validation checks the supported syntax, not whether a mailbox exists or accepts
mail. Addresses are explicit SMTP one-off recipients; no ambiguous name is
silently resolved to a different person.

`SendEmail` stores a Unicode plain-text body. `SendEmailHTML` writes UTF-8 binary
`PR_HTML` and sets `PR_INTERNET_CPID` to 65001. It accepts an HTML fragment or full
document. It does not escape HTML, synthesize a plain-text alternative, or add
attachments/embedded images. Empty subjects and bodies are allowed. Embedded null
characters are rejected to prevent native strings from being silently truncated.

Successful return means **accepted for submission**, not confirmed delivery.
Both methods now call Outlook's **Send/Receive All** after submission and release
of the message. Earlier versions used native `FlushQueues`, which accepted requests
but did not drain the Outbox in the reported setup. Existing calls gain Outlook
synchronization automatically; use `flushQueue:=False` to retain queue-only
behavior or batch several messages before a single `SendAndReceive` call. The
existing parameter name is retained for named-argument compatibility.

Delivery depends on the transport and connectivity; classic Outlook may need to
be running for a cached profile's transport to drain the Outbox. Preparation
and per-property `SetProps` failures raise VBA errors. Preparation failures prevent
submission, and submission failures are never automatically retried. A provider
may retain an unsent item after a failure; inspect the Outbox before manually
retrying to avoid duplicates. There is no delivery receipt or cancellation API.

### Request immediate sending / receiving

```vb
'For already queued messages; does not create or resubmit an email:
outlook.SendAndReceive       'The Outlook button equivalent; hide progress
'outlook.SendAndReceive True 'Optionally show the progress dialog
```

`SendAndReceive(Optional showProgress As Boolean = False)` is a Sub using
[`NameSpace.SendAndReceive`](https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace.sendandreceive).
It attaches to running classic Outlook or starts it, initializes its default
profile by accessing the Inbox folder (without reading messages), and checks that
Outlook and this class use the same default message store. Native `CompareEntryIDs`
checks store identity rather than comparing identifier strings. A mismatch or
failed comparison prevents the Send/Receive request. For an explicitly selected
nondefault MAPI profile, open Outlook with that profile first; this method does
not switch a running Outlook session to another profile.

If Outlook has no Explorer window, the method opens the Inbox window and leaves
it open. This keeps Outlook available for asynchronous synchronization after the
macro ends or `Disconnect` releases this class. Existing windows are reused without
opening another one. The class never calls Outlook's `Quit`. Startup may prompt
for profile setup or credentials if Outlook is not already configured to sign in.

The request affects **all accounts and pending messages in that Outlook profile**,
including messages queued outside this library. It returns before synchronization
finishes; successful return does not confirm delivery. Offline/network/provider
errors can occur later in Outlook. An explicit call raises startup, store validation,
or immediate automation errors; automatic calls report them in `LastSendSyncError`.

### Native-only queue flushing

`UseOutlookSendReceive` defaults to True. Set it to False to use the previous
native-only automatic synchronization path, or call the native method directly:

```vb
outlook.UseOutlookSendReceive = False 'Future sends use native FlushQueues
Debug.Print outlook.FlushQueues()     'Outgoing mail only (default)
Debug.Print outlook.FlushQueues(True) 'Outgoing and incoming mail
```

`FlushQueues(Optional includeReceive As Boolean = False) As Long` uses native
[`IMAPIStatus::FlushQueues`](https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/imapistatus-flushqueues).
It checks the status table's advertised capabilities, prefers the MAPI spooler,
or calls supported individual transports when no capable spooler is present.
The return value counts status objects that accepted the request, not messages
delivered. An explicit call raises if no provider supports flushing or a native
call fails; earlier transports may already have processed mail before a failure.
It does not fall back to Outlook automation or retry failed requests.

**Scope:** flushing can send other pending messages in the same profile, including
ones queued outside this library. It is a mail queue operation, not a promise to
synchronize every folder or calendar like Outlook's **Send/Receive All Folders**.
Providers can reject the request or remain offline. Spooler calls allow asynchronous
processing; direct transport calls are synchronous and may block while transferring
mail. Progress UI is suppressed. In the reported setup, this native path did not
send queued mail until Outlook's Send/Receive button was used, even with Outlook
running. Treat a successful native return only as acceptance of the request.

After each `SendEmail` or `SendEmailHTML`, inspect `LastSendSyncRequested` and
`LastSendSyncError`. True means the selected synchronization request was accepted,
not that delivery completed. A nonempty error means submission succeeded but
synchronization failed or is unsupported: **do not call SendEmail again for that
message**. Call `SendAndReceive` or use Outlook's button instead. Both properties
reset at the start of each send; False with an empty error means synchronization
was skipped (or submission
failed before synchronization). Explicit `FlushQueues` and `SendAndReceive` calls
do not alter these properties. Automatic synchronization errors are deliberately
not raised as send failures, to keep callers from retrying and creating duplicate
messages.

## Search the calendar

```vb
Sub FindAppointments()
    Dim outlook As stdOutlook, appointments As Collection, item As Collection
    Set outlook = stdOutlook.Create()
    Set appointments = outlook.SearchCalendar("planning", maxResults:=50)
    For Each item In appointments
        Debug.Print item("Subject"), item("Location"), _
                    item("StartUtc"), item("EndUtc"), item("IsRecurring")
    Next
    If outlook.CalendarSearchTruncated Then Debug.Print "More calendar items match."
    outlook.Disconnect
End Sub
```

`SearchCalendar(Optional query As String = "", Optional maxResults As Long = 100)`
searches the **default store's Calendar folder**, using a provider-side,
case-insensitive substring restriction across `Subject`, `Location`, and
`OrganizerName`. An empty query lists stored calendar items. Results are ordered
by descending start time when the start property is mapped. The limit is 1–10,000.
`CalendarSearchTruncated` reports whether more matching rows were available and
is independent of contact-search diagnostics. Provider failures raise errors.

Records contain `EntryID`, `Subject`, `OrganizerName`, `MessageClass`, `Location`,
`StartUtc`, `EndUtc`, `IsRecurring`, `RecurrenceExpanded`, and `PropertyErrors`.
Organizer name comes from `PR_SENT_REPRESENTING_NAME`; a provider can leave it
unavailable. Missing properties are `Empty`, with HRESULTs in `PropertyErrors`.
Named appointment properties are resolved for this store with `GetIDsFromNames`;
the class does not assume fixed property IDs or create mappings while searching.

Results are **stored appointments and recurring-series masters**, not expanded
occurrences. `RecurrenceExpanded` is always `False`. For a recurring master,
`StartUtc`/`EndUtc` describe the first occurrence, not the next occurrence. Embedded
exceptions and their changed subjects/locations are not searched separately.
There is no date-range filtering, free/busy calculation, recurrence expansion,
meeting-body search, or traversal of shared/secondary calendars. All returned
times are UTC; convert them for display in the user's time zone.

## Development tests and VBE safeguards

The test runner follows the safeguards in
[stdVBA's AGENTS.md](https://github.com/sancarn/stdVBA/blob/master/AGENTS.md).
Install Python with `pywin32` and run from this repository root:

```powershell
python Tools/Testing/run.py --suite stdOutlook
python tests/harnesses/stdOutlook-tests/run.py --live
python tests/native/test_outlook_abi.py --download
```

The first command uses synthetic native vtables and MAPI-allocated buffers. It
requires classic Outlook's native MAPI installation but does not open a profile.
`--live` additionally opens the default profile and reads identity, contact search,
one Inbox message, and calendar search. Both send methods and queue flushing are
always tested with synthetic vtables, and Outlook synchronization uses a test-only
late-bound fake. Application acquisition is replaced only in the temporary class;
the runner rejects unexpected source changes rather than risk starting Outlook.
**No real email is sent or real queues flushed, even with `--live`.** Logs contain
assertions and counts, not email addresses,
names, subjects, or bodies. Do not run the live test against a profile you do not
intend to access.

The ABI check derives method slots from Microsoft's pinned MAPI header blobs and
checks both 32-bit and 64-bit structure layouts. To run without network access,
use `--headers path/to/MAPIStubLibrary/include` instead of `--download`.

Validation includes live SMTP identity lookup, contact search, an Inbox message
read, and a calendar subject search in 64-bit Excel. Simulated send tests cover
plain text, UTF-8 HTML, To/CC/BCC, write access, property failures, submission
failures, cleanup, and no automatic retries. Queue tests cover automatic flushing,
batch opt-out, spooler preference, direct transport fallback, incoming flags,
unsupported providers, failure isolation, and status-object cleanup. The current
synthetic suite passes 109 assertions in 64-bit Excel. Outlook tests cover existing
and newly started instances, window lifetime, matching/mismatched stores, startup
and synchronization failures, batch opt-out, and no duplicate submission. All 25
native vtable slots match Microsoft's headers, and both pointer-width layout checks
pass. Live Outlook automation startup/synchronization and delivery, every provider,
32-bit Office execution, and long-duration sessions have not been validated by
the test suite.

Before imports, the runner verifies actual CRLF bytes and stages a temporary copy
of the class with test probes. It creates a separate hidden Excel process and
disposable workbook, detects the installed Office/VBE versions, respects managed
policy, and records/restores the exact `AccessVBOM`, `BreakOnAllErrors`, and
`BreakOnServerErrors` registry values. Tests use **Break on Unhandled Errors**.
The parent watchdog dismisses only the owned instance's dialogs, treats every
such dialog as failure, enforces bounded execution, and closes only its own process.
A process lock prevents overlapping stdVBA test transactions.

The runner uses a temporary copy of stdVBA's host with an explicit compile step
and `ShowWindowAsync` so minimizing a busy Office window cannot block the watchdog.
The shared host source is not modified. Compile dialogs can still appear briefly;
the watchdog captures their text and minimizes the owned windows. Cleanup restores
the saved settings after ordinary failures and worker timeouts. As with any
process-based runner, forcibly killing the parent or powering off the machine
can prevent its cleanup.

## Native interface references

- [Outlook Object Model Guard](https://learn.microsoft.com/en-us/office/vba/outlook/how-to/security/security-behavior-of-the-outlook-object-model).
- [MAPILogonEx and profile selection](https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/mapilogonex).
- [Microsoft's MAPI interface headers](https://github.com/microsoft/MAPIStubLibrary/tree/main/include).
- [Address-book restrictions and ANR](https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/address-book-restrictions).
- [MAPI property errors](https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/handling-mapi-property-errors).
- [Submitting messages](https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/imessage-submitmessage).
- [Locating the default Calendar](https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/pidtagipmappointmententryid-canonical-property).
- [Mapping named properties](https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/imapiprop-getidsfromnames).
- [Classic versus new Outlook support](https://support.microsoft.com/en-us/outlook/getstarted/feature-comparison-between-new-outlook-and-classic-outlook).
