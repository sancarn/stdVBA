VERSION 5.00
Begin VB.Form fExtendedTypeInfo 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "  Extended-TypeInfo"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8505
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text1 
      Height          =   4935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Top             =   180
      Width           =   8235
   End
End
Attribute VB_Name = "fExtendedTypeInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowInfo(XML As String, ParentForm)
  Text1.Text = XML
  Show , ParentForm
End Sub
 
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then Cancel = 1: Me.Hide
End Sub
