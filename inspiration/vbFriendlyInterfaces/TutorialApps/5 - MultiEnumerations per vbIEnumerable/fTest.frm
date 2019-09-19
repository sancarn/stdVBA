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
  
  Dim V, oMulti As New cMultiEnumerations
  
  Print vbLf; "Enumeration of an internal Long-Array"
  For Each V In oMulti.EnumerateLngArr
    Print "  "; V
  Next
  
  Print vbLf; "Enumeration of an internal String-Array"
  For Each V In oMulti.EnumerateStrArr
    Print "   "; V
  Next
   
  Print vbLf; "Enumeration of an internal Enum-Type"
  For Each V In oMulti.EnumerateEnumType
    Print "   "; V
  Next
End Sub
