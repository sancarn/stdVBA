VERSION 5.00
Begin VB.Form fTest 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Test-Form (Click Me!)"
   ClientHeight    =   5730
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8475
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows-Standard
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
  AutoRedraw = True: Cls
 
  Dim D As Object, i As Long, Res As String, T1!, T2!
  Set D = New cDispatcher
  
  With D.DispObj 'let's test our IDispatch-Implementation
    'a check for LateBound Properties first
    .foo = "foo" '<- test for Property-Let
    .bar = "bar" '<- test for Property-Let
    Print .foo, .bar, .foobar 'and 3 Property-Get-calls
 
    'LateBound Method-call with Params which are changed ByRef within the callee
    Dim P1 As Long: P1 = 1
    Dim P2 As Long: P2 = 2
    Print .ByRefParamTest(P1, P2); " ... changed ByRef to:"; P1; P2
  End With
  
  'now a small performance-comparison versus the original VB6-Implementation...
  Print vbLf; "Small performance-Test of the IDispatch-implementation:"
  T1 = Timer
    With D 'first, using the HostClass D in LateBound-Mode (VB6-IDispatch-Mode)
      For i = 1 To 250000
        Res = .foobarOnHostClass()
      Next
    End With
  T1 = Timer - T1
  
  T2 = Timer
    With D.DispObj 'now performing the same thing over our vbFriendly IDispatch-Implementation
      For i = 1 To 250000
        Res = .foobar()
      Next
    End With
  T2 = Timer - T2
  
  Print "  VB6-original-IDispatch: " & Format(T1, "0.00sec")
  Print "  vbFriendly-IDispatch: " & Format(T2, "0.00sec")
  Print "  "; Format((T2 - T1) / T1, "Percent") & " Difference" & IIf(App.LogMode, "", ", though native compiled this will be < 10%")
End Sub
