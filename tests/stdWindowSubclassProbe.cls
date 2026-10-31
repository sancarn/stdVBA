VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdWindowSubclassProbe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements stdICallable

Public calls As Long
Public destroyed As Long
Public mode As String
Public lastArgs As Variant
Public value As Variant
Public resized As Boolean
Public window As stdWindow

Public Function Receive(ParamArray args() As Variant) As Variant
  Receive = stdICallable_RunEx(args)
End Function

Private Function stdICallable_Run(ParamArray args() As Variant) As Variant
  stdICallable_Run = stdICallable_RunEx(args)
End Function

Private Function stdICallable_RunEx(ByVal params As Variant) As Variant
  If params(1) = &H82 Then
    destroyed = destroyed + 1
    stdICallable_RunEx = 999
  End If
  If params(1) < &H8001& Or params(1) > &H8002& Then Exit Function
  calls = calls + 1
  lastArgs = params
  Select Case mode
    Case "error": Err.Raise 1234, "probe", "Expected callback error"
    Case "remove": window.Unsubclass
    Case "release": Set window = Nothing
    Case "nested"
      If params(1) = &H8001& Then window.SendMessage &H8002&, 33, 44
    Case "destroy": window.ForceQuit
    Case "forward": Exit Function
    Case "resize"
      ReDim Preserve params(0 To 9)
      params(9) = "resized"
      resized = True
  End Select
  stdICallable_RunEx = value
End Function

Private Function stdICallable_Bind(ParamArray params() As Variant) As stdICallable
  Err.Raise 5
End Function

Private Function stdICallable_SendMessage(ByVal sMessage As String, ByRef success As Boolean, ByVal params As Variant) As Variant
  success = False
End Function
