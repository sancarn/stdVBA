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

Private LWFactory As cLightWeightFactory

Private Sub Form_Load()
  Set LWFactory = New cLightWeightFactory
End Sub

Private Sub Form_Click()
  AutoRedraw = True: Cls
  
  Const strToReflect = "ABC"
 
  Dim oLightWeight1 As Object '<- the same LateBinding as in example #1, which still works
  Set oLightWeight1 = LWFactory.CreateLightWeightObject

  Print vbLf; "Late-Bound MethodCalls on oLightWeight1"
  Print , "Input to reflect: "; strToReflect
  Print , "Reflected result: "; oLightWeight1.StringReflection(strToReflect)
  Print
  Print , "Input to add: 1 and 2"
  Print , "Added result: "; oLightWeight1.AddLongs(1, 2)


  Dim oLightWeight2 As vbIStringsAndLongs '<- but other than in example #1, we support EarlyBinding now too
  Set oLightWeight2 = LWFactory.CreateLightWeightObject

  Print vbLf; "Early-Bound MethodCalls on oLightWeight2"
  Print , "Input to reflect: "; strToReflect
  Print , "Reflected result: "; oLightWeight2.StringReflection(strToReflect)
  Print
  Print , "Input to add: 1 and 2"
  Print , "Added result: "; oLightWeight2.AddLongs(1, 2)
End Sub
