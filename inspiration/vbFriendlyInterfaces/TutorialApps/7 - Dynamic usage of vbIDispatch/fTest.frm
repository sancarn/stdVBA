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
 
  Dim PS As New cPropStorage
 
  With PS.Props 'PS.Props will accept any Property-Name you throw at it, to "store things" (As Variant)
    .foo = "foo"
    .bar = "bar"
    .foobar = .foo & .bar '<- read-out-test for the dynamic props .foo and .bar
    .SomeLong = 123 'fill-in some long...
    .SomeLong = 456 '<- and now test for over-writing an existing Value under the same PropertyName
    Set .TheForm = Me 'test for storing an Object-Reference (in this case in .TheForm)
    
    'Ok, now the Test-PrintOuts for the above
    Print "foo: "; .foo
    Print "bar: "; .bar
    Print "foobar: "; .foobar
    Print "SomeLong: "; .SomeLong
    Print "TheForm.Caption: "; .TheForm.Caption
  End With
End Sub
