VERSION 5.00
Begin VB.Form fTest 
   Caption         =   "Test-Form"
   ClientHeight    =   3975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7035
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame frLightWeight 
      Caption         =   "LightWeight-Approach"
      Height          =   3195
      Left            =   3780
      TabIndex        =   4
      Top             =   300
      Width           =   2655
      Begin VB.CommandButton cmdRemoveAll 
         Caption         =   "Remove all Inst."
         Height          =   495
         Index           =   1
         Left            =   420
         TabIndex        =   7
         Top             =   1740
         Width           =   1695
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add 100000 Inst."
         Height          =   495
         Index           =   1
         Left            =   420
         TabIndex        =   6
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton cmdFindBirthDay 
         Caption         =   "Find BirthDay"
         Height          =   495
         Index           =   1
         Left            =   420
         TabIndex        =   5
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label lblCurCount 
         Caption         =   "Col-Count: 0"
         Height          =   255
         Index           =   1
         Left            =   420
         TabIndex        =   9
         Top             =   540
         Width           =   1635
      End
   End
   Begin VB.Frame frClassic 
      Caption         =   "Classic-Approach"
      Height          =   3195
      Left            =   300
      TabIndex        =   0
      Top             =   300
      Width           =   2655
      Begin VB.CommandButton cmdFindBirthDay 
         Caption         =   "Find BirthDay"
         Height          =   495
         Index           =   0
         Left            =   420
         TabIndex        =   3
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add 100000 Inst."
         Height          =   495
         Index           =   0
         Left            =   420
         TabIndex        =   2
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton cmdRemoveAll 
         Caption         =   "Remove all Inst."
         Height          =   495
         Index           =   0
         Left            =   420
         TabIndex        =   1
         Top             =   1740
         Width           =   1695
      End
      Begin VB.Label lblCurCount 
         Caption         =   "Col-Count: 0"
         Height          =   255
         Index           =   0
         Left            =   420
         TabIndex        =   8
         Top             =   540
         Width           =   1635
      End
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const AddrCount As Long = 100000

Private AddrFactoryClassic As cAddressClassicFactory
Private AddrFactoryLWeight As cAddressLWeightFactory
 
Private Sub Form_Load()
  Set AddrFactoryClassic = New cAddressClassicFactory
  Set AddrFactoryLWeight = New cAddressLWeightFactory
End Sub
 
Private Sub cmdAdd_Click(Index As Integer)
  cmdRemoveAll_Click Index
  
  Dim i As Long, D As Date
  D = Now - 12345
  Tag = Timer
    If Index = 0 Then 'classic-mode
      For i = 1 To AddrCount
        AddrFactoryClassic.Add i, "Name" & i, "LastName" & i, D + i
      Next
      lblCurCount(Index) = "Col-Count: " & AddrFactoryClassic.Count
    Else 'lightweight-mode
      For i = 1 To AddrCount
        AddrFactoryLWeight.Add i, "Name" & i, "LastName" & i, D + i
      Next
      lblCurCount(Index) = "Col-Count: " & AddrFactoryLWeight.Count
    End If
  Caption = Format((Timer - Tag) * 1000, "0msec")
End Sub

Private Sub cmdRemoveAll_Click(Index As Integer)
  Tag = Timer
    If Index = 0 Then 'classic-mode
      AddrFactoryClassic.RemoveAll
      lblCurCount(Index) = "Col-Count: " & AddrFactoryClassic.Count
    Else 'lightweight-mode
      AddrFactoryLWeight.RemoveAll
      lblCurCount(Index) = "Col-Count: " & AddrFactoryLWeight.Count
    End If
  Caption = Format((Timer - Tag) * 1000, "0msec")
End Sub
 
Private Sub cmdFindBirthDay_Click(Index As Integer)
  Dim AClassic As cAddressClassic, ALWeight As vbIAddress, CC As Long
  Tag = Timer
    If Index = 0 Then 'classic-mode
      For Each AClassic In AddrFactoryClassic
        If AClassic.BirthDayToday Then CC = CC + 1
      Next
    Else 'lightweight-mode
      For Each ALWeight In AddrFactoryLWeight
        If ALWeight.BirthDayToday Then CC = CC + 1
      Next
    End If
  Caption = Format((Timer - Tag) * 1000, "0msec") & " to find: " & CC & " persons have birthday today"
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
  AddrFactoryClassic.RemoveAll
  AddrFactoryLWeight.RemoveAll
End Sub
