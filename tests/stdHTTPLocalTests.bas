Attribute VB_Name = "stdHTTPLocalTests"
Option Explicit
Private baseURL As String

Public Function RunTests(ByVal fixtureURL As String, ByVal refusedURL As String) As Variant
  On Error GoTo Failed
  baseURL = fixtureURL
  Dim r As stdHTTP, path As Variant
  For Each path In Array("/empty", "/normal")
    Set r = Request(baseURL & CStr(path), 2000)
    r.Await
    Check r.isFinished, "Completion"
    Check r.ResponseStatus = 200, "Status"
    If path = "/empty" Then
      Check r.ResponseText = "", "Empty body"
    Else
      Check r.ResponseText = "hello", "Normal body"
    End If
  Next
  Set r = Request(baseURL & "/slow", 800)
  Check Not r.isFinished, "Delayed request must be pending"
  CheckFailure r, 10000, "Delayed request timeout"
  Set r = Request(baseURL & "/trickle", 400)
  CheckFailure r, 10000, "Overall deadline during continuous response"
  Set r = Request(refusedURL, 10000)
  CheckFailure r, &H80072EFD, "Connection refused"
  Set r = Request(baseURL & "/disconnect", 10000)
  CheckFailure r, 0, "Disconnected transport"
  RunTests = "Passed local HTTP regressions"
  Exit Function
Failed:
  RunTests = "FAIL: " & Err.Number & " " & Err.Description
End Function

Private Sub CheckFailure(ByVal r As stdHTTP, ByVal expected As Long, ByVal label As String)
  Dim started As Double: started = Timer
  Dim n As Long, msg As String, done As Boolean
  On Error Resume Next
  r.Await
  n = Err.Number: msg = Err.Description
  Err.Clear
  On Error GoTo 0
  Check n <> 0 And Len(msg) > 0, label & " missing error"
  If expected <> 0 Then Check n = expected, label & " unexpected error " & n
  Dim elapsed As Double: elapsed = Timer - started
  If elapsed < 0 Then elapsed = elapsed + 86400
  Check elapsed < 12, label & " exceeded wait bound"
  On Error Resume Next
  done = r.isFinished
  Dim repeated As Long: repeated = Err.Number
  Dim repeatedMsg As String: repeatedMsg = Err.Description
  Err.Clear
  On Error GoTo 0
  Check repeated = n And repeatedMsg = msg, label & " terminal polling error changed"
  On Error Resume Next
  r.Await
  repeated = Err.Number: repeatedMsg = Err.Description
  Err.Clear
  On Error GoTo 0
  Check repeated = n And repeatedMsg = msg, label & " repeated Await error changed"
End Sub

Private Sub Check(ByVal passed As Boolean, ByVal label As String)
  If Not passed Then Err.Raise vbObjectError + 121, "stdHTTPLocalTests", label
End Sub

Private Function Request(ByVal url As String, ByVal timeout As Long) As stdHTTP
  Set Request = stdHTTP.Create(url, ThreadingStyle:=HTTPAsync, options:=stdHTTP.CreateOptions(TimeoutMS:=timeout, Proxy:=stdHTTP.CreateProxy(ProxyTypeNone)))
End Function
