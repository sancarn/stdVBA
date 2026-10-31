Attribute VB_Name = "WindowIDEProbe"
Option Explicit
Private watched As stdWindow
Private handler As stdWindowSubclassProbe
Private modeless As Object

Public Function Arm() As Variant
  Set handler = New stdWindowSubclassProbe
  handler.value = 2468
  Set watched = stdWindow.Create("STATIC", 0, 0, 10, 10, "stdWindow IDE probe")
  watched.Subclass handler
  Arm = Array(watched.handle, stdWindow.protGetSubclassThunk())
End Function

Public Function ArmModeless() As Variant
  Set modeless = UserForms.Add("stdWindowTestForm")
  modeless.StartUpPosition = 0
  modeless.Left = -10000
  modeless.Top = -10000
  modeless.Show 0
  Set handler = New stdWindowSubclassProbe
  handler.value = 2468
  Set watched = stdWindow.CreateFromIUnknown(modeless)
  watched.Subclass handler, False
  ArmModeless = Array(watched.handle, stdWindow.protGetSubclassThunk())
End Function

Public Function CheckDetached() As Boolean
  CheckDetached = Not watched.isSubclassed
End Function

Public Sub Cleanup()
  If Not watched Is Nothing Then
    watched.Unsubclass
    If Not modeless Is Nothing Then
      Unload modeless
    ElseIf watched.Exists Then
      watched.ForceQuit
    End If
  End If
  Set watched = Nothing
  Set handler = Nothing
  Set modeless = Nothing
End Sub

Public Sub EndNow()
  End
End Sub

Public Function StopAndResume() As Variant
  Stop
  StopAndResume = watched.SendMessage(&H8001&, 1, 2)
End Function
