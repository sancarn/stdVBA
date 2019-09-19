VERSION 5.00
Begin VB.Form fTest 
   Appearance      =   0  '2D
   BackColor       =   &H80000005&
   Caption         =   "Click Me!"
   ClientHeight    =   4185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows-Standard
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private Sub Form_Click()
  Dim MyObj As IMyClass
  Set MyObj = modMyClassFactory.CreateInstance  'creates our little Byte COM-Instance
  Caption = "Result = " & MyObj.AddTwoLongs(1, 2) 'and here we call its single Method
End Sub
 
Private Sub Form_Load()
  MsgBox "Note, that this little Demo stands on its own (is not requiring the vbFriendly- Interface Dll reference as all the other Demos)." & vbLf & vbLf & _
         "A TypeLib is required and used instead though (MyClassTypeLib.tlb, which is contained in this #0-Tutorial-Folder)." & vbLf & vbLf & _
         "This Demo is contained in the Tutorial, to show the 'bare minimum COM-implementation in C-style' (using only *.bas-Modules)"
End Sub
