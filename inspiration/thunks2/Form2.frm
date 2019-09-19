VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4995
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   4995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Show Sample Project"
      Height          =   540
      Left            =   105
      TabIndex        =   2
      Top             =   3810
      Width           =   4710
   End
   Begin VB.TextBox Text1 
      Height          =   1965
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form2.frx":0000
      Top             =   1755
      Width           =   4695
   End
   Begin VB.ListBox List1 
      Height          =   1320
      ItemData        =   "Form2.frx":0006
      Left            =   90
      List            =   "Form2.frx":0019
      TabIndex        =   0
      Top             =   270
      Width           =   4725
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Select Case List1.ListIndex
    Case 0: frmSample1.Show
    Case 1: frmSample2.Show
    Case 2: frmSample3.Show
    Case 3: frmSample4.Show
    Case 4: frmSample5.Show
    End Select
End Sub

Private Sub List1_DblClick()
    Command1.Value = True
End Sub

Private Sub Form_Load()
    List1.ListIndex = 0
End Sub

Private Sub List1_Click()
    Select Case List1.ListIndex
    Case 0
        Text1.Text = "The sample project shows how hook and subclassing thunks are combined to create " & _
            "owner-drawn comboboxes on demand. The hook is used to look for a combobox as it is " & _
            "being created so that it can have its window style changed to support owner-drawing. " & _
            "Subclassing is used to sublcass the combobox container (owner) which is where " & _
            "owner drawn messages are sent"
    Case 1
        Text1.Text = "The sample project shows a simple example of creating custom window classes " & _
            "via RegisterClassEx API. Custom window classes need a custom window procedure. The " & _
            "sample project creates that window procedure via a thunk."
    Case 2
        Text1.Text = "The sample project shows two examples of simple callbacks. One callback " & _
            "is based on a Windows API while the other is for the cDecl qSort API"
    Case 3
        Text1.Text = "The sample project shows an example of hooking a stdPicture object assigned " & _
            "to image controls. The sample could be expanded to enable loading PNGs into an " & _
            "image control. Basically, any image can be converted to a premultiplied RGB format " & _
            "that can be rendred with AlphaBlend API (as shownn in the project), or drawn " & _
            "directly to the image control via GDI+ or other libraries."
    Case 4
        Text1.Text = "The sample project shows various ways standard subclassing can crash a " & _
            "project. By attempting those actions, the thunks prevent crashing in the same " & _
            "scenarios."
    End Select
End Sub
