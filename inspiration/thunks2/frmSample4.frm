VERSION 5.00
Begin VB.Form frmSample4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hooking COM Sample"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Toggle Hook On/Off"
      Height          =   960
      Left            =   4140
      TabIndex        =   0
      Top             =   120
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "Left image (GIF) will be centered/scaled proportionally when hooked. The right image (bitmap) will show transparency"
      Height          =   855
      Left            =   375
      TabIndex        =   1
      Top             =   165
      Width           =   3690
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   2250
      Index           =   1
      Left            =   3000
      Picture         =   "frmSample4.frx":0000
      Top             =   1200
      Width           =   2430
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2265
      Index           =   0
      Left            =   330
      Picture         =   "frmSample4.frx":17BF2
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   2670
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H80000003&
      FillStyle       =   4  'Upward Diagonal
      Height          =   2265
      Left            =   330
      Top             =   1200
      Width           =   5115
   End
End
Attribute VB_Name = "frmSample4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oTasker(0 To 1) As Object
Dim m_hDC As Long

Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal xDest As Long, ByVal yDest As Long, ByVal WidthDest As Long, ByVal HeightDest As Long, ByVal hdcSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long, ByVal Blendfunc As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long

Private Sub Command1_Click()
    
    If oTasker(0).HookedObjPtr = 0 Then ' not hooked, do hook
        oTasker(0).HookedObjPtr = ObjPtr(Image1(0).Picture)
        oTasker(1).HookedObjPtr = ObjPtr(Image1(1).Picture)
    Else
        oTasker(0).HookedObjPtr = 0     ' unhook
        oTasker(1).HookedObjPtr = 0
    End If
    Me.Refresh  ' for changes to visibly take effect, image container needs refereshing
        
End Sub

Private Sub Form_Load()

    Dim cThunks As clsThunks, i As Long
    Dim tpic As StdPicture
    
    Set cThunks = New clsThunks
    Set tpic = New StdPicture
    
    ' here we are creating two thunks (1 for each image)
    ' when done, tPic-hook is removed and ready for user
    ' to toggle hooks as desired
    For i = 0 To 1
        ' called first: identify host (Me) and method count of tPic's primary interface (18)
        cThunks.CreateTasker_HookCOM 0, 0, 18, Me
        ' hooking 2 methods, so need 2 calls, one for each method
        cThunks.CreateTasker_HookCOM 1, 9, 10, tpic ' hook IPicture::Render
        cThunks.CreateTasker_HookCOM 2, 17, 1, tpic ' hook IPicture::get_Attributes
        ' called last to set/return the Tasker object. passing -1 = want IDE-safety
        Set oTasker(i) = cThunks.CreateTasker_HookCOM(-1, 0, 0, Nothing)
        oTasker(i).HookedObjPtr = 0 ' remove hook
    Next
    oTasker(1).Tag = 1  ' flag = use AlphaBlend for rendering
    m_hDC = CreateCompatibleDC(Me.hDC)  ' used for AlphaBlend calls
    Set tpic = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase oTasker()
    DeleteDC m_hDC: m_hDC = 0
End Sub

' //////////////// DO NOT ADD ANY NEW CODE FROM HERE TO THE END OF THE CODE PAGE \\\\\\\\\\\\\\\\\\\\
Private Function IUnknownRelease(ByVal pTasker As Object, ByRef EatMessage As Boolean, _
                                ByVal pInterface As IPicture) As Long

    ' not used, just an example of including this if one wanted to know when
    ' the hooked object was completely released. The method hook call is:
    '   cThunks.CreateTasker_HookCOM [host ordinal], 3, 0, [object] ' hook IUnknown::Release
    
    If pTasker.HookedObjPtr = 0 Then MsgBox "Hooked object has been released"

End Function    ' ordinal #3

Private Function IPictureGetAttrs(ByVal pTasker As Object, ByRef EatMessage As Boolean, _
                                ByVal pInterface As IPicture, ByRef pAttributes As Long) As Long
    
    If pTasker.Tag = 1 Then
        ' VB needs to be told that the bitmap has transparency (set bit 2)
        pAttributes = pInterface.Attributes Or 2
        EatMessage = True   ' prevent picture from getting this message
    End If

End Function    ' ordinal #2

Private Function IPictureRender(ByVal pTasker As Object, ByRef EatMessage As Boolean, _
                                ByVal pInterface As IPicture, _
                                ByVal hDC As Long, ByVal pX As Long, ByVal pY As Long, _
                                ByVal pCx As Long, ByVal pCy As Long, _
                                ByVal pXSrc As Long, ByVal pYSrc As Long, _
                                ByVal pCxSrc As Long, ByVal pCySrc As Long, ByVal prcWBounds As Long) As Long

    Dim sngX As Single, sngY As Single, hObj As Long
    Dim cx As Long, cy As Long, dx As Long, dy As Long
    
    pCxSrc = pInterface.Width: pCySrc = pInterface.Height
    ' note: ScaleX,ScaleY should not be used this way for DPI-aware applications
    cx = ScaleX(pCxSrc, vbHimetric, vbPixels)
    cy = ScaleY(pCySrc, vbHimetric, vbPixels)
    
    ' if actual size and AlphaBlend usage is n/a, then no action needed
    If Not (cx = pCx And cy = pCy And pTasker.Tag = 0) Then
        sngX = pCx / cx: sngY = pCy / cy
        If sngY < sngX Then sngX = sngY
        dx = cx * sngX: dy = cy * sngX
        pX = pX + ((pCx - dx) \ 2)
        pY = pY + ((pCy - dy) \ 2)
        If pTasker.Tag = 1 Then ' use AlphaBlend
            pInterface.SelectPicture m_hDC, hObj, 0&    ' similar to SelectObject() API
            AlphaBlend hDC, pX, pY, dx, dy, pInterface.CurDC, 0, 0, cx, cy, &H1FF0000
            pInterface.SelectPicture hObj, 0&, 0&
        Else
            pInterface.Render hDC, pX, pY, dx, dy, 0&, pCySrc, pCxSrc, -pCySrc, prcWBounds
        End If
        EatMessage = True   ' prevent picture from getting this message
    End If

End Function    ' ordinal #1
' //////////////////////////////// DO NOT ADD ANY NEW CODE BELOW THIS BANNER \\\\\\\\\\\\\\\\\\\\\\\\\\\\\


