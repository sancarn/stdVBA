VERSION 5.00
Begin VB.Form fTest 
   BackColor       =   &H00FFFFFF&
   Caption         =   "vbFriendly IPicture-Implementation"
   ClientHeight    =   5730
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8295
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
   ScaleHeight     =   382
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   553
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CheckBox chkGifAnimate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gif-Animation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   6600
      Style           =   1  'Grafisch
      TabIndex        =   5
      Top             =   120
      Width           =   1515
   End
   Begin VB.Timer tmrGifAnimate 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7680
      Top             =   1080
   End
   Begin VB.CheckBox chkPattern 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Checker-BGnd"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   120
      Style           =   1  'Grafisch
      TabIndex        =   1
      Top             =   120
      Width           =   1515
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Test IDispatch"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   3360
      Style           =   1  'Grafisch
      TabIndex        =   3
      Top             =   120
      Width           =   1515
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close Form"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   4980
      Style           =   1  'Grafisch
      TabIndex        =   4
      Top             =   120
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Start new Form"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   1740
      Style           =   1  'Grafisch
      TabIndex        =   2
      Top             =   120
      Width           =   1515
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   1515
      Left            =   300
      ScaleHeight     =   1515
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   1920
      Width           =   1875
   End
   Begin VB.Image Image2 
      Appearance      =   0  '2D
      Enabled         =   0   'False
      Height          =   1515
      Left            =   5820
      Stretch         =   -1  'True
      Top             =   1860
      Width           =   2055
   End
   Begin VB.Image Image1 
      Appearance      =   0  '2D
      Enabled         =   0   'False
      Height          =   1515
      Left            =   2940
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   2115
   End
   Begin VB.Image Image3 
      Appearance      =   0  '2D
      Enabled         =   0   'False
      Height          =   1575
      Left            =   2280
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   3495
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private Sub Form_Load()
  Set Icon = LoadPictureEx("Png3_16x16", vbPicTypeIcon) 'just to show, that a Form-Icon can easily be derived from a Png-Resource
  Set MouseIcon = LoadPictureEx("HandCursor,11,5", vbPicTypeIcon): MousePointer = 99 'load an Alpha-MouseCursor from a PNG (note the HotSpot-Offsets, given behind the ImgKey)
  
  Set Command1.Picture = LoadPictureEx("Ico3") 'CommandButton-Picture, derived from an Icon-Resource
  Set Command2.Picture = LoadPictureEx("Png2_31x31") 'CommandButton-Picture from a PNG (in an "unusual" size)
  Set Command3.Picture = LoadPictureEx("Png3_29x29") 'CommandButton-Picture from a PNG (in an "unusual" size)
  Set chkGifAnimate.Picture = LoadPictureEx("Gif1_55x55") 'an animated Gif-Resource
  
  Set Picture1.Picture = LoadPictureEx("Png3_128x128", , PICTURE_TRANSPARENT Or PICTURE_SCALABLE, True) 'the scalable-Attribute ensures AutoStretching in a PicBox
  Set Image1.Picture = LoadPictureEx("Ico1", , , True) 'an Icon is rendered in a VB-Image-Ctl (Image1 has Stretch=True)
  Set Image2.Picture = LoadPictureEx("Ico2", , , True) 'an Icon is rendered in a VB-Image-Ctl (Image2 has Stretch=True)

  Set Image3.Picture = LoadPictureEx("Png1") 'a PNG is rendered in a normal VB-Image-Ctl (Image3 has Stretch=True)

  Set chkPattern.Picture = LoadPictureEx("ChkSmall") 'finally also a PNG-resource for our CheckBox-Button (to make the Alpha-Renderings more obvious over a Checker-Pattern)
End Sub

Private Sub chkPattern_Click()
  Set Picture = Nothing
  If chkPattern Then Set Picture = LoadPicture(App.Path & "\Res\CheckerBG.jpg") 'VBs LoadPicture is sufficient to supply the (jpg) Form-BG-Image
  chkPattern.Refresh
End Sub
Private Sub Command1_Click()
Dim F As New fTest
    F.Show 'start a new Form, to easier check for Memory- or Handle-Leaks (by opening - and
           'closing new Form-Instances, watching the appropriate Columns in the TaskManager-ListView)
    'Note, that cGDIPlusCache (in conjunction with cPictureEx) ensures, that these
    'additional Forms will not consume more image-memory whilst showing and rendering
    'all these StdPicture/IPicureDisp/IPicture resources, since those are now truly "shared ones"
End Sub
Private Sub Command2_Click()
Dim PDisp As StdPicture, MyPicture As New cPictureEx
Set PDisp = MyPicture.Picture("Png1")
    MsgBox PDisp.Handle & ", " & PDisp.Width & ", " & PDisp.Height 'just to show, that IPictureDisp (IDispatch) works too
End Sub
Private Sub Command3_Click()
  Unload Me
End Sub
Private Sub chkGifAnimate_Click()
  tmrGifAnimate.Enabled = chkGifAnimate.Value
End Sub

Private Sub Form_Resize() 'to demonstrate the high-quality scaling-behaviour of the Alpha-Resources
  Picture1.Move 0.05 * ScaleWidth, 0.42 * ScaleHeight, 0.15 * ScaleWidth, 0.22 * ScaleHeight
  Image1.Move 0.4 * ScaleWidth, 0.42 * ScaleHeight, 0.15 * ScaleWidth, 0.22 * ScaleHeight
  Image2.Move 0.75 * ScaleWidth, 0.4 * ScaleHeight, 0.18 * ScaleWidth, 0.24 * ScaleHeight
  Image3.Move 0.25 * ScaleWidth, 0.68 * ScaleHeight, 0.5 * ScaleWidth, 0.3 * ScaleHeight
End Sub
 
Private Sub tmrGifAnimate_Timer()
Static FrIdx As Long
  If Not GDIPlusCache.Exists("Gif1_55x55" & "|" & FrIdx) Then FrIdx = 0
  Set chkGifAnimate.Picture = LoadPictureEx("Gif1_55x55" & "|" & FrIdx)
      chkGifAnimate.Refresh
  FrIdx = FrIdx + 1
End Sub
