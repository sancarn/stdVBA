Attribute VB_Name = "stdWindowSubclassTests"
Option Explicit
Private checks As Long
#If VBA7 Then
Private Declare PtrSafe Function LoadIcon Lib "user32" Alias "LoadIconW" (ByVal instance As LongPtr, ByVal name As LongPtr) As LongPtr
Private Declare PtrSafe Function VirtualQuery Lib "kernel32" (ByVal address As LongPtr, buffer As Any, ByVal length As LongPtr) As LongPtr
Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" (ByVal address As LongPtr, ByVal length As LongPtr, ByVal allocationType As Long, ByVal protection As Long) As LongPtr
Private Declare PtrSafe Function VirtualFree Lib "kernel32" (ByVal address As LongPtr, ByVal length As LongPtr, ByVal freeType As Long) As Long
Private Declare PtrSafe Function VirtualProtect Lib "kernel32" (ByVal address As LongPtr, ByVal length As LongPtr, ByVal protection As Long, oldProtection As Long) As Long
Private Declare PtrSafe Function GetModuleHandleA Lib "kernel32" (ByVal name As String) As LongPtr
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal length As LongPtr)
#Else
Private Declare Function LoadIcon Lib "user32" Alias "LoadIconW" (ByVal instance As Long, ByVal name As Long) As Long
Private Declare Function VirtualQuery Lib "kernel32" (ByVal address As Long, buffer As Any, ByVal length As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal address As Long, ByVal length As Long, ByVal allocationType As Long, ByVal protection As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal address As Long, ByVal length As Long, ByVal freeType As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (ByVal address As Long, ByVal length As Long, ByVal protection As Long, oldProtection As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal name As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal length As Long)
#End If

#If VBA7 Then
Private Function IsReleased(ByVal address As LongPtr) As Boolean
#Else
Private Function IsReleased(ByVal address As Long) As Boolean
#End If
  Dim info(0 To 47) As Byte
  #If Win64 Then
    Const stateOffset = 32, infoSize = 48
  #Else
    Const stateOffset = 16, infoSize = 28
  #End If
  If VirtualQuery(address, info(0), infoSize) = 0 Then Err.Raise 1, , "VirtualQuery failed."
  IsReleased = (info(stateOffset) = 0 And info(stateOffset + 1) = 0 And info(stateOffset + 2) = 1 And info(stateOffset + 3) = 0)
End Function

Public Function RunResolverTests() As String
  On Error GoTo Failed
  checks = 0
  Assert "reject null module", stdWindow.protFindEbMode(0) = 0
  #If VBA7 Then
    Dim runtime As LongPtr: runtime = GetModuleHandleA("vbe7.dll")
  #Else
    Dim runtime As Long: runtime = GetModuleHandleA("vbe6.dll")
  #End If
  Assert "resolve loaded non-exported EbMode", stdWindow.protFindEbMode(runtime) <> 0
  Dim scenario As Long
  For scenario = 0 To 11
    TestResolverImage scenario
  Next
  RunResolverTests = "Passed " & checks & " EbMode resolver checks"
  Exit Function
Failed:
  RunResolverTests = "FAILED resolver after " & checks & ": " & Err.Number & " " & Err.Description
End Function

Private Sub WriteImageLong(ByRef image() As Byte, ByVal offset As Long, ByVal value As Long)
  CopyMemory image(offset), value, 4
End Sub

'Exercise the real scanner on mapped synthetic PE images and executable targets.
Private Sub TestResolverImage(ByVal scenario As Long)
  Dim address As LongPtr: address = VirtualAlloc(0, &H5000, &H3000, &H4)
  If address = 0 Then Err.Raise 7
  On Error GoTo Failed
  Dim image(0 To &H4FFF) As Byte
  WriteImageLong image, 0, &H5A4D
  WriteImageLong image, &H3C, &H80
  WriteImageLong image, &H80, &H4550
  #If Win64 Then
    WriteImageLong image, &H84, &H8664&
    Const optSize = &HF0, optMagic = &H20B, pointerOffset = 37
    Const hexPattern = "48894C240848895424104C894424184C894C242048B81111111111111111480BC0743248B80000000000000000FFD04883F802742048B81111111111111111488B4C2408488B5424104C8B4424184C8B4C2420FFE04833C0C21111"
  #Else
    WriteImageLong image, &H84, &H14C
    Const optSize = &HE0, optMagic = &H10B, pointerOffset = 10
    Const hexPattern = "A1000000000BC07413B800000000FFD083F8027407B811111111FFE033C0C21111"
  #End If
  WriteImageLong image, &H86, 1
  WriteImageLong image, &H94, optSize
  WriteImageLong image, &H98, optMagic
  WriteImageLong image, &HD0, &H5000
  Dim section As Long: section = &H98 + optSize
  WriteImageLong image, section + 8, &H3000
  WriteImageLong image, section + 12, &H1000
  WriteImageLong image, section + 36, &H60000020
  Dim pattern() As Byte: ReDim pattern(0 To Len(hexPattern) / 2 - 1)
  Dim i As Long
  For i = 0 To UBound(pattern)
    pattern(i) = CByte("&H" & Mid$(hexPattern, i * 2 + 1, 2))
  Next
  Dim target As LongPtr: target = address + &H2000
  CopyMemory pattern(pointerOffset), target, LenB(target)
  CopyMemory image(&H1000), pattern(0), UBound(pattern) + 1
  image(&H2000) = &HC3: image(&H2100) = &HC3
  Dim expected As LongPtr: expected = target
  Select Case scenario
    Case 1 'Unknown DOS header.
      image(0) = 0: expected = 0
    Case 2 'Wrong architecture.
      WriteImageLong image, &H84, 0: expected = 0
    Case 3 'Pointer outside the candidate module.
      target = 1
      CopyMemory image(&H1000 + pointerOffset), target, LenB(target)
      expected = 0
    Case 4 'More than one distinct executable target is ambiguous.
      target = address + &H2100
      CopyMemory pattern(pointerOffset), target, LenB(target)
      CopyMemory image(&H1100), pattern(0), UBound(pattern) + 1
      expected = 0
    Case 5 'Repeated templates agreeing on one pointer are valid.
      CopyMemory image(&H1100), pattern(0), UBound(pattern) + 1
    Case 6 'Partial template crossing the executable section boundary.
      For i = 0 To UBound(pattern): image(&H1000 + i) = 0: Next
      CopyMemory image(&H3FF0), pattern(0), UBound(pattern) + 1
      expected = 0
    Case 7 'Recognizable template with one incorrect fixed byte.
      image(&H1000) = image(&H1000) Xor 1: expected = 0
    Case 8 'Executable section headers cannot make an RW target callable.
      expected = 0
    Case 9 'The last complete template in a section is still found.
      For i = 0 To UBound(pattern): image(&H1000 + i) = 0: Next
      CopyMemory image(&H4000 - UBound(pattern) - 1), pattern(0), UBound(pattern) + 1
    Case 10 'Section size outside the image.
      WriteImageLong image, section + 8, &H5001: expected = 0
    Case 11 'Unbounded PE header offsets are rejected before reading them.
      WriteImageLong image, &H3C, 4092: expected = 0
  End Select
  CopyMemory ByVal address, image(0), &H5000
  Dim previous As Long
  If scenario <> 8 Then
    If VirtualProtect(address + &H1000, &H3000, &H20, previous) = 0 Then Err.Raise 1
  End If
  Dim found As LongPtr: found = stdWindow.protFindEbMode(address)
  Assert "mapped resolver scenario " & scenario, found = expected
  VirtualFree address, 0, &H8000&
  Exit Sub
Failed:
  Dim number As Long: number = Err.Number
  Dim description As String: description = Err.Description
  VirtualFree address, 0, &H8000&
  Err.Raise number, "EbMode resolver test", description
End Sub

Private Sub Assert(ByVal name As String, ByVal passed As Boolean)
  checks = checks + 1
  If Not passed Then Err.Raise vbObjectError + 86, "stdWindowSubclassTests", name
End Sub

#If VBA7 Then
Public Function ReceiveModule(ByVal hwnd As LongPtr, ByVal message As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Variant
#Else
Public Function ReceiveModule(ByVal hwnd As Long, ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long) As Variant
#End If
  If message = &H8001& Then ReceiveModule = 713 + wParam + lParam
End Function

Public Function RunTests() As String
  On Error GoTo Failed
  checks = 0
  Dim wnd As stdWindow, second As stdWindow, owner As stdWindow, child As stdWindow
  Dim probe As stdWindowSubclassProbe: Set probe = New stdWindowSubclassProbe
  Set wnd = stdWindow.Create("STATIC", 10, 10, 160, 100, "stdWindow issue 86", WS_OVERLAPPEDWINDOW)
  Set owner = stdWindow.Create("STATIC", 10, 10, 100, 100)
  Set wnd.owner = owner
  Assert "set/get owner", wnd.owner.handle = owner.handle
  Assert "owned window not on taskbar", Not wnd.isTaskbarVisible
  wnd.isTaskbarVisible = True
  Assert "force taskbar", wnd.isTaskbarVisible And wnd.isAppWindow And Not wnd.isToolWindow
  Assert "taskbar preserves owner", wnd.owner.handle = owner.handle
  Assert "hidden stays hidden", Not wnd.Visible
  wnd.Visible = True
  wnd.isTaskbarVisible = False
  Assert "hide taskbar", Not wnd.isTaskbarVisible And wnd.isToolWindow And Not wnd.isAppWindow
  Assert "visible stays visible", wnd.Visible
  wnd.Visible = False
  Set wnd.owner = Nothing
  Assert "clear owner", wnd.owner Is Nothing
  wnd.isToolWindow = False
  Assert "unowned default taskbar", wnd.isTaskbarVisible
  wnd.isMinimiseButtonVisible = False
  wnd.isMaximiseButtonVisible = False
  Assert "hide buttons", Not wnd.isMinimiseButtonVisible And Not wnd.isMaximiseButtonVisible
  wnd.isMinimiseButtonVisible = True
  wnd.isMaximiseButtonVisible = True
  Assert "show buttons", wnd.isMinimiseButtonVisible And wnd.isMaximiseButtonVisible
  Set child = stdWindow.Create("STATIC", 0, 0, 10, 10, , WS_CHILD, , wnd.handle)
  Assert "child has no owner", child.owner Is Nothing
  Assert "child not on taskbar", Not child.isTaskbarVisible
  On Error Resume Next
  child.isTaskbarVisible = True
  Dim number As Long: number = Err.Number
  Err.Clear
  On Error GoTo Failed
  Assert "reject child taskbar", number = 5
  child.ForceQuit
  Set child = Nothing

  probe.value = 5678
  wnd.Subclass probe
  Dim thunkAddress As LongPtr: thunkAddress = stdWindow.protGetSubclassThunk()
  Assert "installed", wnd.isSubclassed
  Dim received As LongPtr: received = wnd.SendMessage(&H8001&, 123, -456)
  Assert "handled message result=" & received, received = 5678
  Assert "callback arguments", probe.lastArgs(0) = wnd.handle And probe.lastArgs(1) = &H8001& And probe.lastArgs(2) = 123 And probe.lastArgs(3) = -456
  #If Win64 Then
    Dim wide As LongPtr: wide = &H123456789ABC^
    probe.value = wide
    Assert "64-bit result", wnd.SendMessage(&H8001&, wide, -wide) = wide
    Assert "64-bit params", probe.lastArgs(2) = wide And probe.lastArgs(3) = -wide
  #End If
  probe.value = 0
  Assert "zero is handled", wnd.SendMessage(&H8001&, 0, 0) = 0
  probe.mode = "resize": probe.value = 19
  Assert "callback can resize arguments", wnd.SendMessage(&H8001&, 0, 0) = 19 And probe.resized
  probe.mode = "forward"
  Assert "forward unhandled", wnd.SendMessage(&H8001&, 0, 0) = 0
  Assert "default procedure still runs", wnd.SendMessage(&HE, 0, 0) = Len(wnd.Caption)
  probe.mode = "error"
  Assert "callback error forwards", wnd.SendMessage(&H8001&, 0, 0) = 0
  Assert "callback error recorded", wnd.SubclassErrorNumber = 1234 And wnd.SubclassErrorDescription = "Expected callback error"
  probe.mode = "": probe.value = Null
  Assert "invalid callback result forwards", wnd.SendMessage(&H8001&, 0, 0) = 0
  Assert "invalid result recorded", wnd.SubclassErrorNumber = 94
  probe.mode = "nested"
  Set probe.window = wnd
  probe.value = 12
  wnd.Subclass probe
  Assert "reinstallation clears error", wnd.SubclassErrorNumber = 0
  Dim before As Long: before = probe.calls
  #If Win64 Then
    wnd.PostMessage &H8002&, wide, -wide
  #Else
    wnd.PostMessage &H8002&, 123, -456
  #End If
  Dim i As Long
  For i = 1 To 100
    DoEvents
    If probe.calls > before Then Exit For
  Next
  Assert "posted callback received", probe.calls = before + 1
  #If Win64 Then
    Assert "posted pointer widths", probe.lastArgs(2) = wide And probe.lastArgs(3) = -wide
  #End If
  before = probe.calls
  Assert "nested result", wnd.SendMessage(&H8001&, 1, 2) = 12
  Assert "nested calls", probe.calls = before + 2
  Assert "nested separate arguments", probe.lastArgs(1) = &H8002& And probe.lastArgs(2) = 33 And probe.lastArgs(3) = 44
  probe.mode = "remove"
  Assert "self-removal result", wnd.SendMessage(&H8001&, 0, 0) = 12
  Assert "self-removed", Not wnd.isSubclassed
  wnd.Unsubclass
  wnd.Unsubclass
  Assert "repeat removal", Not wnd.isSubclassed
  Assert "idle thunk released", IsReleased(thunkAddress)
  Set probe.window = Nothing

  'Release the last reference to a subclassed wrapper inside its own callback.
  probe.mode = "release": probe.value = 31
  Set probe.window = stdWindow.CreateFromHwnd(wnd.handle)
  probe.window.Subclass probe
  Assert "last wrapper release result", wnd.SendMessage(&H8001&, 0, 0) = 31
  before = probe.calls
  Assert "released wrapper detached", wnd.SendMessage(&H8001&, 0, 0) = 0 And probe.calls = before
  wnd.Unsubclass

  'Exercise stdCallback as well as an arbitrary interface implementation.
  wnd.Subclass stdCallback.CreateFromModule("stdWindowSubclassTests", "ReceiveModule")
  Assert "module stdCallback works", wnd.SendMessage(&H8001&, 7, 3) = 723
  probe.mode = "": probe.value = 89
  wnd.Subclass stdCallback.CreateFromObjectMethod(probe, "Receive")
  Assert "stdCallback works", wnd.SendMessage(&H8001&, 0, 0) = 89
  Dim other As stdWindowSubclassProbe: Set other = New stdWindowSubclassProbe
  other.mode = "forward"
  Set second = stdWindow.CreateFromHwnd(wnd.handle)
  second.Subclass other
  before = probe.calls
  Assert "chain result", wnd.SendMessage(&H8001&, 0, 0) = 89
  Assert "chain calls both", other.calls = 1 And probe.calls = before + 1
  wnd.Unsubclass
  Assert "remove older chain member", second.isSubclassed And wnd.SendMessage(&H8001&, 0, 0) = 0
  Set second = Nothing
  Assert "wrapper termination removes subclass", wnd.SendMessage(&H8001&, 0, 0) = 0

  'Reject foreign-thread windows before installing a callback.
  On Error Resume Next
  stdWindow.CreateFromDesktop().Subclass probe
  number = Err.Number
  Err.Clear
  On Error GoTo Failed
  Assert "reject foreign thread", number = 5
  On Error Resume Next
  wnd.Subclass Nothing
  number = Err.Number
  Err.Clear
  On Error GoTo Failed
  Assert "reject missing callback", number = 5

  wnd.Subclass probe
  probe.mode = "destroy": probe.value = 17
  Set probe.window = wnd
  Assert "destroy from callback result", wnd.SendMessage(&H8001&, 0, 0) = 17
  Assert "destroy detaches", Not wnd.isSubclassed And Not wnd.Exists
  Assert "destroy message delivered once", probe.destroyed = 1
  Set probe.window = Nothing
  wnd.Unsubclass
  Set wnd = Nothing
  owner.ForceQuit
  Set owner = Nothing
  RunTests = "Passed " & checks & " stdWindow checks"
  Exit Function
Failed:
  RunTests = "FAILED after " & checks & ": " & Err.Number & " " & Err.Description
  On Error Resume Next
  RunTests = RunTests & " calls=" & probe.calls & " bridgeError=" & wnd.SubclassErrorNumber & " " & wnd.SubclassErrorDescription
  Set probe.window = Nothing
  If Not second Is Nothing Then second.Unsubclass
  If Not child Is Nothing Then child.ForceQuit
  If Not wnd Is Nothing Then
    wnd.Unsubclass
    If wnd.Exists Then wnd.ForceQuit
  End If
  If Not owner Is Nothing Then owner.ForceQuit
End Function

'The isolated runner creates an empty stdWindowTestForm component before this call.
Public Function RunUserFormTests() As String
  On Error GoTo Failed
  checks = 0
  Dim form As Object: Set form = UserForms.Add("stdWindowTestForm")
  Dim wnd As stdWindow: Set wnd = stdWindow.CreateFromIUnknown(form)
  Dim probe As stdWindowSubclassProbe: Set probe = New stdWindowSubclassProbe
  Assert "UserForm handle", wnd.Exists
  wnd.isMinimiseButtonVisible = True
  wnd.isMaximiseButtonVisible = True
  Assert "UserForm buttons", wnd.isMinimiseButtonVisible And wnd.isMaximiseButtonVisible
  wnd.isTaskbarVisible = True
  Assert "UserForm taskbar", wnd.isTaskbarVisible And Not wnd.Visible
  Dim icon As LongPtr: icon = LoadIcon(0, 32512)
  Assert "shared test icon", icon <> 0
  wnd.HICON = icon
  Assert "UserForm small icon", wnd.SendMessage(&H7F, 0, 0) = icon
  Assert "UserForm big icon", wnd.SendMessage(&H7F, 1, 0) = icon
  probe.value = 654
  wnd.Subclass probe
  Assert "UserForm callback", wnd.SendMessage(&H8001&, 7, 8) = 654
  Unload form
  Assert "UserForm unload detaches", Not wnd.isSubclassed And Not wnd.Exists And probe.destroyed = 1
  wnd.Unsubclass
  RunUserFormTests = "Passed " & checks & " UserForm checks"
  Exit Function
Failed:
  RunUserFormTests = "FAILED UserForm after " & checks & ": " & Err.Number & " " & Err.Description
  On Error Resume Next
  If Not wnd Is Nothing Then wnd.Unsubclass
  If Not form Is Nothing Then Unload form
End Function
