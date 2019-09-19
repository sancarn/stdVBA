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
  
  Dim L, E As New cEnumerable
  
  Print vbLf; "Enumeration over the (-4) VB-marked Enumeration-Method"
  For Each L In E 'enumeration on the "naked E" will work, since we marked EnumerateLngArr with a -4
    Print L
  Next
  
  Print vbLf; "Enumeration, using the EnumerateLngArr-Method explicitely"
  For Each L In E.EnumerateLngArr '<- but this will work too, since we also implement IDispatch and handle the -4 request manually
    Print L
  Next
End Sub
