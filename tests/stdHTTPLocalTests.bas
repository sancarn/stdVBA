Attribute VB_Name = "stdHTTPLocalTests"
Option Explicit
Public Function RunTests() As Variant
  On Error GoTo Failed
  Dim r As stdHTTP, path As Variant, started As Double, n As Long, msg As String
  For Each path In Array("/empty", "/normal")
    Set r = Request(CStr(path), 1000)
    r.Await
    If Not r.isFinished Or r.ResponseStatus <> 200 Then Err.Raise 5, , "Completion"
    If path = "/empty" Then
      If r.ResponseText <> "" Then Err.Raise 5, , "Empty response"
    Else
      If r.ResponseText <> "hello" Then Err.Raise 5, , "Normal response"
    End If
  Next
  For Each path In Array("/slow", "/disconnect")
    started = Timer
    Set r = Request(CStr(path), 200)
    On Error Resume Next
    r.Await
    n = Err.Number: msg = Err.Description
    Err.Clear
    On Error GoTo Failed
    If n = 0 Or msg = "" Then Err.Raise 5, , "Missing transport error"
    If Timer - started > 3 Then Err.Raise 5, , "Unbounded wait"
    On Error Resume Next
    Dim done As Boolean: done = r.isFinished
    If Err.Number <> n Or Err.Description <> msg Then
      On Error GoTo Failed
      Err.Raise 5, , "Terminal error changed"
    End If
    Err.Clear
    On Error GoTo Failed
  Next
  RunTests = "Passed local HTTP regressions"
  Exit Function
Failed:
  RunTests = "FAIL: " & Err.Number & " " & Err.Description
End Function
Private Function Request(ByVal path As String, ByVal timeout As Long) As stdHTTP
  Set Request = stdHTTP.Create("http://127.0.0.1:18761" & path, ThreadingStyle:=HTTPAsync, options:=stdHTTP.CreateOptions(TimeoutMS:=timeout, Proxy:=stdHTTP.CreateProxy(ProxyTypeNone)))
End Function
