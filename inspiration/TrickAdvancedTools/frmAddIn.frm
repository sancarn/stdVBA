VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Trick Advanced Tools"
   ClientHeight    =   10020
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10020
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  'None
      Height          =   3135
      Index           =   4
      Left            =   5460
      ScaleHeight     =   3135
      ScaleWidth      =   5175
      TabIndex        =   26
      Top             =   6240
      Visible         =   0   'False
      Width           =   5175
      Begin VB.Frame Frame3 
         Caption         =   "Events"
         Height          =   2835
         Left            =   120
         TabIndex        =   27
         Top             =   120
         Width           =   4995
         Begin VB.TextBox txtCompBefore 
            Height          =   975
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   29
            Top             =   480
            Width           =   4755
         End
         Begin VB.TextBox txtCompAfter 
            Height          =   975
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   28
            Top             =   1740
            Width           =   4755
         End
         Begin VB.Label Label10 
            Caption         =   "Before:"
            Height          =   315
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   4395
         End
         Begin VB.Label Label9 
            Caption         =   "After:"
            Height          =   315
            Left            =   120
            TabIndex        =   30
            Top             =   1500
            Width           =   4395
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2820
      TabIndex        =   1
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   4680
      Width           =   1215
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  'None
      Height          =   4575
      Index           =   3
      Left            =   120
      ScaleHeight     =   4575
      ScaleWidth      =   5175
      TabIndex        =   18
      Top             =   5280
      Visible         =   0   'False
      Width           =   5175
      Begin VB.TextBox txtLinkerOptions 
         Height          =   1335
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   24
         Top             =   300
         Width           =   4935
      End
      Begin VB.Frame Frame1 
         Caption         =   "Events"
         Height          =   2835
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   4995
         Begin VB.TextBox txtLinkBefore 
            Height          =   975
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   21
            Top             =   480
            Width           =   4755
         End
         Begin VB.TextBox txtLinkAfter 
            Height          =   975
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   20
            Top             =   1740
            Width           =   4755
         End
         Begin VB.Label Label5 
            Caption         =   "Before:"
            Height          =   315
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   4395
         End
         Begin VB.Label Label2 
            Caption         =   "After:"
            Height          =   315
            Left            =   120
            TabIndex        =   22
            Top             =   1500
            Width           =   4395
         End
      End
      Begin VB.Label Label6 
         Caption         =   "Linker options:"
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   60
         Width           =   4395
      End
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  'None
      Height          =   4575
      Index           =   2
      Left            =   5460
      ScaleHeight     =   4575
      ScaleWidth      =   5175
      TabIndex        =   10
      Top             =   1620
      Visible         =   0   'False
      Width           =   5175
      Begin VB.Frame fraEvents 
         Caption         =   "Events"
         Height          =   2835
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   4995
         Begin VB.TextBox txtBuildAfter 
            Height          =   975
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   17
            Top             =   1740
            Width           =   4755
         End
         Begin VB.TextBox txtBuildBefore 
            Height          =   975
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   15
            Top             =   480
            Width           =   4755
         End
         Begin VB.Label Label4 
            Caption         =   "After:"
            Height          =   315
            Left            =   120
            TabIndex        =   16
            Top             =   1500
            Width           =   4395
         End
         Begin VB.Label Label3 
            Caption         =   "Before:"
            Height          =   315
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   4395
         End
      End
      Begin VB.TextBox txtCompilerOptions 
         Height          =   1335
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   12
         Top             =   300
         Width           =   4935
      End
      Begin VB.Label Label1 
         Caption         =   "Compiler options:"
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   60
         Width           =   4395
      End
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  'None
      Height          =   1095
      Index           =   1
      Left            =   5460
      ScaleHeight     =   1095
      ScaleWidth      =   5175
      TabIndex        =   7
      Top             =   420
      Visible         =   0   'False
      Width           =   5175
      Begin VB.TextBox txtArgsComp 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   4995
      End
      Begin VB.Label lblArgIDE 
         Caption         =   "Compiled form:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   3855
      End
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  'None
      Height          =   4575
      Index           =   0
      Left            =   180
      ScaleHeight     =   4575
      ScaleWidth      =   5175
      TabIndex        =   3
      Top             =   420
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CheckBox chkArrBnd 
         Caption         =   "Remove array bounds check in IDE"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   4995
      End
      Begin VB.CheckBox chkFloat 
         Caption         =   "Remove floating point error check in IDE"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   420
         Width           =   4995
      End
      Begin VB.CheckBox chkIntOvflw 
         Caption         =   "Remove integer overflow check in IDE."
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   4815
      End
   End
   Begin ComctlLib.TabStrip tabPages 
      Height          =   5115
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   9022
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   5
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Global checking"
            Key             =   "global_checking"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Conditional compilation"
            Key             =   "cond_compilation"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Build"
            Key             =   "build"
            Object.Tag             =   ""
            Object.ToolTipText     =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Linking"
            Key             =   "linking"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Compile events"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' // frmAddIn.frm - Main window of Trick Advanced Tools addin
' // © Krivous Anatoly Anatolevich (The trick), 2016

Option Explicit

Private Const TCM_FIRST     As Long = &H1300
Private Const TCM_HITTEST   As Long = (TCM_FIRST + 13)

Private Type POINTAPI
    x       As Long
    y       As Long
End Type
Private Type TCHITTESTINFO
    pt      As POINTAPI
    flags   As Long
End Type

Private Declare Function SendMessage Lib "user32" _
                         Alias "SendMessageA" ( _
                         ByVal hWnd As Long, _
                         ByVal wMsg As Long, _
                         ByVal wParam As Long, _
                         ByRef lParam As Any) As Long
                         
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lFlags  As Long
    
    On Error Resume Next
    
    ' // Each bit in the value represents an flag
    ' // 0 - Remove integer overflow checking
    ' // 1 - Remove float operation checking
    ' // 2 - Remove array bounds checking
    If chkIntOvflw.Value = vbChecked Then lFlags = lFlags Or 1
    If chkFloat.Value = vbChecked Then lFlags = lFlags Or 2
    If chkArrBnd.Value = vbChecked Then lFlags = lFlags Or 4
    
    ' // Store string values
    pVBInstance.ActiveVBProject.WriteProperty "TAT", "GlobalChecking", lFlags
    pVBInstance.ActiveVBProject.WriteProperty "TAT", "CondCOMP", txtArgsComp.Text
    pVBInstance.ActiveVBProject.WriteProperty "VBCompiler", "C2Switches", txtCompilerOptions.Text
    pVBInstance.ActiveVBProject.WriteProperty "VBCompiler", "LinkSwitches", txtLinkerOptions.Text
    pVBInstance.ActiveVBProject.WriteProperty "TAT", "LinkBefore", txtLinkBefore.Text
    pVBInstance.ActiveVBProject.WriteProperty "TAT", "LinkAfter", txtLinkAfter.Text
    pVBInstance.ActiveVBProject.WriteProperty "TAT", "CompBefore", txtCompBefore.Text
    pVBInstance.ActiveVBProject.WriteProperty "TAT", "CompAfter", txtCompAfter.Text
    pVBInstance.ActiveVBProject.WriteProperty "TAT", "BuildBefore", txtBuildBefore.Text
    pVBInstance.ActiveVBProject.WriteProperty "TAT", "BuildAfter", txtBuildAfter.Text

    Unload Me
    
End Sub

Private Sub Form_Load()
    Dim lFlags  As Long
    
    On Error Resume Next
    
    ' // Set icon of window
    SetIconToForm Me.hWnd, "ICON"

    lFlags = pVBInstance.ActiveVBProject.ReadProperty("TAT", "GlobalChecking")
    
    If lFlags And 1 Then chkIntOvflw.Value = vbChecked
    If lFlags And 2 Then chkFloat.Value = vbChecked
    If lFlags And 4 Then chkArrBnd.Value = vbChecked
    
    txtArgsComp.Text = pVBInstance.ActiveVBProject.ReadProperty("TAT", "CondCOMP")
    txtCompilerOptions.Text = pVBInstance.ActiveVBProject.ReadProperty("VBCompiler", "C2Switches")
    txtLinkerOptions.Text = pVBInstance.ActiveVBProject.ReadProperty("VBCompiler", "LinkSwitches")
    txtBuildBefore.Text = pVBInstance.ActiveVBProject.ReadProperty("TAT", "BuildBefore")
    txtBuildAfter.Text = pVBInstance.ActiveVBProject.ReadProperty("TAT", "BuildAfter")
    txtLinkBefore.Text = pVBInstance.ActiveVBProject.ReadProperty("TAT", "LinkBefore")
    txtLinkAfter.Text = pVBInstance.ActiveVBProject.ReadProperty("TAT", "LinkAfter")
    txtCompBefore.Text = pVBInstance.ActiveVBProject.ReadProperty("TAT", "CompBefore")
    txtCompAfter.Text = pVBInstance.ActiveVBProject.ReadProperty("TAT", "CompAfter")
    
    ' // Place controls
    AdjustControls
    
End Sub

' // Place all the controls. Set form size
Private Sub AdjustControls()
    Dim ncSizeCX    As Long
    Dim ncSizeCY    As Long
    
    ' // Calculate non-client area size
    ncSizeCX = Me.width - Me.ScaleX(Me.ScaleWidth, Me.ScaleMode, vbTwips)
    ncSizeCY = Me.height - Me.ScaleY(Me.ScaleHeight, Me.ScaleMode, vbTwips)
    
    Me.width = (tabPages.width + 10 * Screen.TwipsPerPixelX) + ncSizeCX
    Me.height = (tabPages.height + 15 * Screen.TwipsPerPixelY) + ncSizeCY + cmdOK.height
    
    cmdOK.Move (Me.ScaleWidth - (cmdOK.width + cmdCancel.width + 5 * Screen.TwipsPerPixelX)) / 2, tabPages.height + 10 * Screen.TwipsPerPixelY
    cmdCancel.Move cmdOK.Left + cmdOK.width + 5 * Screen.TwipsPerPixelX, cmdOK.Top
    
    UpdateTab
    
End Sub

' // Show pictureboz container that belongs the selected tab
Private Sub UpdateTab()
    Dim lIndex  As Long
    Dim picCont As PictureBox
    
    lIndex = tabPages.SelectedItem.index - 1

    For Each picCont In picPage
        
        If picCont.index <> lIndex Then
            picCont.Visible = False
        Else
        
            picCont.Move tabPages.ClientLeft, tabPages.ClientTop, tabPages.ClientWidth, tabPages.ClientHeight
            picCont.Visible = True
            
        End If
        
    Next
    
End Sub

Private Sub tabPages_MouseDown( _
            ByRef Button As Integer, _
            ByRef Shift As Integer, _
            ByRef x As Single, _
            ByRef y As Single)
    Dim lIndex  As Long
    Dim hitInf  As TCHITTESTINFO
    
    hitInf.pt.x = x / Screen.TwipsPerPixelX
    hitInf.pt.y = y / Screen.TwipsPerPixelY
    
    ' // Get selected item
    lIndex = SendMessage(tabPages.hWnd, TCM_HITTEST, 0, hitInf)
    
    If lIndex = tabPages.SelectedItem.index - 1 Or lIndex = -1 Then Exit Sub
    
    picPage(tabPages.SelectedItem.index - 1).Visible = False
    picPage(lIndex).Move tabPages.ClientLeft, tabPages.ClientTop, tabPages.ClientWidth, tabPages.ClientHeight
    picPage(lIndex).Visible = True

End Sub
