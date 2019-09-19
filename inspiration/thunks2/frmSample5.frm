VERSION 5.00
Begin VB.Form frmSample5 
   Caption         =   "Crash Trials"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   3015
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Execute End Statement"
      Height          =   510
      Left            =   195
      TabIndex        =   2
      Top             =   225
      Width           =   4365
   End
   Begin VB.CommandButton Command2 
      Caption         =   """With"" block End execution"
      Height          =   510
      Left            =   210
      TabIndex        =   1
      Top             =   1530
      Width           =   4365
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delayed IDE Termination"
      Height          =   510
      Left            =   210
      TabIndex        =   0
      Top             =   900
      Width           =   4365
   End
   Begin VB.Label Label1 
      Caption         =   "FYI: The form is being subclassed, along with all 3 buttons while you are playing around"
      Height          =   720
      Left            =   270
      TabIndex        =   3
      Top             =   2190
      Width           =   4350
   End
End
Attribute VB_Name = "frmSample5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cThunks As clsThunks
Dim oTasker As Object
Dim oTimer As Object

Dim bBreakMe As Boolean

Private Sub Form_Load()
    Set cThunks = New clsThunks
    Set oTasker = cThunks.CreateTasker_Subclass(Me, 1, Me.hWnd, True)
    oTasker.AddWindow Array(Command1.hWnd, Command2.hWnd, Command3.hWnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oTasker = Nothing
    Set cThunks = Nothing
End Sub

Private Sub Command1_Click()
    Set oTimer = cThunks.CreateTasker_Timer(Me, 2, 3000)
    MsgBox "In a few seconds, hover mouse over the form. The caption will change. At that point, an End was executed " & _
        "within the form's subclass procedure. No Crash." & vbCrLf & vbCrLf & _
        "But when you click the Ok button to close this msgbox, the project will terminate.", vbInformation + vbOKOnly
End Sub

Private Sub Command2_Click()

    MsgBox "One case where other thunks may not be able to prevent a crash is when the IDE is " & _
        "inside a With statement/block and termination occurs. When you click the ok button, " & _
        "a debug msgbox will appear, click 'Debug'. After you return to the IDE, click the " & _
        "toolbar's End button -- no crash", vbInformation + vbOKOnly

    With Me
        With .Command2
            With .Parent
                With .Font
                    .Charset = .Size / 0
                End With
            End With
        End With
    End With
End Sub

Private Sub Command3_Click()
    MsgBox "When this msgbox closes, an End statement will be executed", vbInformation + vbOKOnly
    End
End Sub


' //////////////// DO NOT ADD ANY NEW CODE FROM HERE TO THE END OF THE CODE PAGE \\\\\\\\\\\\\\\\\\\\
Private Function myTimerProc(ByVal TickCount As Long, ByVal pTasker As Object) As Long
    
    bBreakMe = True
    Set oTimer = Nothing        ' example of releasing the timer tasker within the callback

End Function    ' ordinal #2

Private Function myWindowProc(ByVal hWnd As Long, ByVal uMsg As Long, _
                        ByVal wParam As Long, ByVal lParam As Long, _
                        ByVal hWndTag As Long, ByVal pTasker As Object, _
                        ByRef EatMessage As Boolean) As Long

    If bBreakMe Then
        bBreakMe = False
        Me.Caption = "End Executed"
        End
    End If

End Function    ' ordinal #1
' //////////////////////////////// DO NOT ADD ANY NEW CODE BELOW THIS BANNER \\\\\\\\\\\\\\\\\\\\\\\\\\\\\


