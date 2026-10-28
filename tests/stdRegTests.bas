Attribute VB_Name = "stdRegTests"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function RegDeleteTreeA Lib "advapi32.dll" (ByVal hKey As LongPtr, ByVal lpSubKey As String) As Long
#Else
Private Declare Function RegDeleteTreeA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
#End If

Public Sub testAll()
  Test.Topic "stdReg creation"
  Dim root As stdReg, child As stdReg, item As stdReg, same As stdReg, parent As stdReg
  Dim shell As Object: Set shell = CreateObject("WScript.Shell")
  Dim testName As String: testName = "stdVBA_RegistryCreation_" & Format(Now, "yyyymmddhhnnss") & "_" & Replace(CStr(Timer), ".", "_")
  Dim testPath As String: testPath = "Software\" & testName
  Dim fullPath As String: fullPath = "HKEY_CURRENT_USER\" & testPath
  Dim failure As Long, description As String, result As Long
  On Error GoTo Failed
  Set parent = stdReg.CreateFromKey("HKCU\Software")
  Set root = parent.addKey(testName)
  Test.Assert "addKey returns a key", root.RegType = iKey
  Set child = root.addKey("MyApp")
  child.value = "Open with My&App"
  Test.Assert "Key value reads its unnamed item", child.value = "Open with My&App"
  Test.Assert "Default value is an actual registry string", shell.RegRead(fullPath & "\MyApp\") = "Open with My&App"
  With child.addItem("icon")
    Test.Assert "New item is an empty string", .ItemType = Value_String And .value = ""
    .value = "C:\Program Files\MyApp\MyIcon.ico"
    Test.Assert "Named string round trip", .value = "C:\Program Files\MyApp\MyIcon.ico"
  End With
  Test.Assert "String bytes are correctly encoded for registry", shell.RegRead(fullPath & "\MyApp\icon") = "C:\Program Files\MyApp\MyIcon.ico"
  With child.addKey("command")
    .value = """C:\Program Files\MyApp\MyApp.exe"" ""%1"""
  End With
  Test.Assert "Command-chain example", shell.RegRead(fullPath & "\MyApp\command\") = """C:\Program Files\MyApp\MyApp.exe"" ""%1"""
  Set same = root.addKey("MyApp")
  Test.Assert "Existing key is not overwritten", same.value = child.value
  Set item = child.addItem("Count", Value_DWORD)
  Test.Assert "DWORD created at zero", item.ItemType = Value_DWORD And item.value = 0
  item.value = 42
  Test.Assert "DWORD writes actual numeric data", shell.RegRead(fullPath & "\MyApp\Count") = 42
  Set item = child.addItem("Count")
  Test.Assert "Existing value data and type are retained", item.ItemType = Value_DWORD And item.value = 42
  Call item.SetValue("changed", Value_String)
  Test.Assert "SetValue refreshes cached type", item.ItemType = Value_String And item.value = "changed"
  item.value = ""
  Test.Assert "Empty string can be written", shell.RegRead(fullPath & "\MyApp\Count") = ""
  Set item = child.addItem("(Default)")
  Test.Assert "Default alias points to unnamed value", item.value = child.value
  Set item = child.addItem("Bytes", Value_Binary)
  Dim bytes() As Byte, actual As Variant
  bytes = item.value
  Test.Assert "Empty binary created", IsArray(bytes)
  item.value = bytes
  Test.Assert "Empty binary can be assigned", IsArray(item.value)
  ReDim bytes(0 To 2): bytes(0) = 1: bytes(1) = 0: bytes(2) = 255
  item.value = bytes
  actual = item.value
  Test.Assert "Binary data survives write/read", actual(LBound(actual)) = 1 And actual(LBound(actual) + 1) = 0 And actual(UBound(actual)) = 255
  actual = shell.RegRead(fullPath & "\MyApp\Bytes")
  Test.Assert "Binary is interoperable", actual(LBound(actual)) = 1 And actual(UBound(actual)) = 255
  Set item = child.addItem("Expanded", Value_String_WithEnvVars)
  item.value = "%TEMP%\sample.txt"
  Test.Assert "Expand-string preserves raw value and type", item.ItemType = Value_String_WithEnvVars And item.value = "%TEMP%\sample.txt"
  Set same = root.addKey("Nested\Leaf")
  Test.Assert "Nested key paths are created", same.name = "Leaf" And same.Parent.name = "Nested"
  Test.Assert "Children can be enumerated", child.Keys.Count = 1 And child.Items.Count >= 5

  On Error Resume Next
  Set same = item.addKey("invalid")
  failure = Err.Number: Err.Clear
  On Error GoTo Failed
  Test.Assert "Items reject subkeys", failure <> 0
  On Error Resume Next
  Set same = child.addKey("")
  failure = Err.Number: Err.Clear
  On Error GoTo Failed
  Test.Assert "Empty key names rejected", failure = 5
  On Error Resume Next
  Set same = child.addItem("invalid\name")
  failure = Err.Number: Err.Clear
  On Error GoTo Failed
  Test.Assert "Ambiguous item path rejected", failure = 5
  On Error Resume Next
  Set same = child.addItem("Unsupported", Value_String_Array)
  failure = Err.Number: Err.Clear
  On Error GoTo Failed
  Test.Assert "Unimplemented value writers rejected before creation", failure = 5

  Set same = Nothing: Set item = Nothing: Set child = Nothing: Set root = Nothing: Set parent = Nothing
  result = RegDeleteTreeA(&H80000001, testPath)
  Test.Assert "Disposable test subtree removed", result = 0
  Exit Sub
Failed:
  failure = Err.Number: description = Err.Description
  Set same = Nothing: Set item = Nothing: Set child = Nothing: Set root = Nothing: Set parent = Nothing
  Call RegDeleteTreeA(&H80000001, testPath)
  Err.Raise failure, "stdRegTests", description
End Sub

