VERSION 5.00
Begin VB.Form frmSample2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custom Window Classes"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5130
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Unload Sample Windows"
      Enabled         =   0   'False
      Height          =   690
      Left            =   255
      TabIndex        =   5
      Top             =   1695
      Width           =   1890
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create Sample Windows"
      Height          =   690
      Left            =   255
      TabIndex        =   1
      Top             =   900
      Width           =   1890
   End
   Begin VB.Label lblHwnd 
      Caption         =   "hWnd:"
      Height          =   315
      Index           =   4
      Left            =   2490
      TabIndex        =   8
      Top             =   2145
      Width           =   2280
   End
   Begin VB.Label lblHwnd 
      Caption         =   "hWnd:"
      Height          =   315
      Index           =   3
      Left            =   2490
      TabIndex        =   7
      Top             =   1830
      Width           =   2280
   End
   Begin VB.Label Label2 
      Caption         =   "When hWnds are closed, their label's background changes to red"
      Height          =   495
      Left            =   315
      TabIndex        =   6
      Top             =   2535
      Width           =   4635
   End
   Begin VB.Label lblHwnd 
      Caption         =   "hWnd:"
      Height          =   315
      Index           =   2
      Left            =   2490
      TabIndex        =   4
      Top             =   1515
      Width           =   2280
   End
   Begin VB.Label lblHwnd 
      Caption         =   "hWnd:"
      Height          =   315
      Index           =   1
      Left            =   2490
      TabIndex        =   3
      Top             =   1200
      Width           =   2280
   End
   Begin VB.Label lblHwnd 
      Caption         =   "hWnd:"
      Height          =   315
      Index           =   0
      Left            =   2490
      TabIndex        =   2
      Top             =   885
      Width           =   2280
   End
   Begin VB.Label Label1 
      Caption         =   "This sample uses a thunk to create a class window procedure for a custom window class."
      Height          =   495
      Left            =   195
      TabIndex        =   0
      Top             =   225
      Width           =   4215
   End
End
Attribute VB_Name = "frmSample2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cThunks As clsThunks
Private oTasker As Object
Private Const WM_DESTROY = 2


'///////////// following used for creating custom window classes, none of it for the thunks

Private Type WNDCLASSEX
    cbSize As Long
    Style As Long
    lpfnWndProc As Long
    cbClsExtra As Long
    cbWndExtra As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
    hIconSm As Long
End Type
Private Const CS_OWNDC = &H20
Private Const CS_HREDRAW = &H2
Private Const CS_VREDRAW = &H1
Private Const COLOR_APPWORKSPACE = 12
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_SYSMENU As Long = &H80000
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const CW_USEDEFAULT As Long = &H80000000
Private Const WS_EX_CLIENTEDGE As Long = &H200&
Private Const WS_OVERLAPPED As Long = &H0&
Private Const WS_OVERLAPPEDWINDOW As Long = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function UpdateWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function RegisterClassEx Lib "user32.dll" Alias "RegisterClassExA" (lpwcx As WNDCLASSEX) As Long

Private Sub Command1_Click()

    Dim classinfo As WNDCLASSEX  ' holds info about the class
    Dim classatom As Long        ' receives an atom to the class just registered
    Dim lHwnds() As Long, n As Long

    If oTasker Is Nothing Then
        Set oTasker = cThunks.CreateTasker_CustomClass(Me, 1, , True)
        oTasker.AddFilterMessage 0, WM_DESTROY
    
        ' Load a decription of the new class into the strucure.
        ' very basic example...
        With classinfo
            .cbSize = Len(classinfo)
            ' Class style: give each window its own DC; redraw when resized.
            .Style = CS_OWNDC Or CS_HREDRAW Or CS_VREDRAW
            .hInstance = App.hInstance
            .hbrBackground = COLOR_APPWORKSPACE
            .lpszClassName = "CustomClass"
            .lpfnWndProc = oTasker.AddressOf ' use thunk as window procedure
            Debug.Assert .lpfnWndProc <> 0
        End With
    
        classatom = RegisterClassEx(classinfo)
        oTasker.Atom = classatom            ' thunk will unregister class for us when it unloads
    End If
    
    For n = 1 To 5
        pvCreateCustWindow                  ' create a few sample windows
    Next
    
    ReDim lHwnds(0 To 4)                    ' sample usage of GetWindows method
    oTasker.GetWindows 5, VarPtr(lHwnds(0))
    For n = 0 To 4
        lblHwnd(n).Caption = "hWnd: " & lHwnds(n)
        lblHwnd(n).BackColor = vbGreen
        oTasker.TagHwnd(lHwnds(n)) = n      ' we'll cache the label index w/the hWnd
    Next                                    ' queried in the class procedure
    Command2.Enabled = True
    Command1.Enabled = False

End Sub

Private Sub Command2_Click()
    oTasker.RemoveWindow -1                 ' close all custom windows
    Command2.Enabled = False                ' command1 re-enables when last window is destroyed
End Sub

Private Sub Form_Load()
    Set cThunks = New clsThunks
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set oTasker = Nothing
    Set cThunks = Nothing
End Sub

Private Function pvCreateCustWindow() As Long

   Dim lWnd As Long
   lWnd = CreateWindowEx(WS_EX_CLIENTEDGE, "CustomClass", "Sample Custom Window", _
        WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, 320, 120, 0, 0, App.hInstance, ByVal 0)

    ShowWindow lWnd, 1
    UpdateWindow lWnd
    pvCreateCustWindow = lWnd

End Function

' //////////////// DO NOT ADD ANY NEW CODE FROM HERE TO THE END OF THE CODE PAGE \\\\\\\\\\\\\\\\\\\\
Private Function myClassProc(ByVal hWnd As Long, ByVal uMsg As Long, _
                        ByVal wParam As Long, ByVal lParam As Long, _
                        ByVal pTasker As Object, ByRef EatMessage As Boolean) As Long

    If uMsg = WM_DESTROY Then
        lblHwnd(pTasker.TagHwnd(hWnd)).BackColor = vbRed
        If pTasker.State = 1 Then ' this is the last custom hWnd
            Command1.Enabled = True
            Command2.Enabled = False
        End If
    End If

End Function    ' ordinal #1
' //////////////////////////////// DO NOT ADD ANY NEW CODE BELOW THIS BANNER \\\\\\\\\\\\\\\\\\\\\\\\\\\\\

