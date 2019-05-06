VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "QuadTree"
   ClientHeight    =   8175
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   ScaleHeight     =   545
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   752
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "LOOP"
      Height          =   1575
      Left            =   9120
      TabIndex        =   3
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Query"
      Height          =   1095
      Left            =   9120
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "draw"
      Height          =   1215
      Left            =   9120
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7695
      Left            =   240
      ScaleHeight     =   513
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   569
      TabIndex        =   0
      Top             =   120
      Width           =   8535
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim Q           As cQuadTree

Dim curridx     As Long


Private Sub Command1_Click()


    vbDrawCC.SetSourceRGB 0.5, 0.5, 0.1
    vbDrawCC.Paint


    Q.DRAW -1
    PIC.Refresh

    vbDRAW.Srf.DrawToDC PicHDC

End Sub





Private Sub Command2_Click()
    Dim rX() As Double, rY() As Double, rIDX() As Long
    ReDim rX(0)
    ReDim rY(0)
    ReDim rIDX(0)
    Dim I       As Long


    Q.Query 100, 100, 200, 200, rX(), rY(), rIDX(), True


    For I = 1 To UBound(rIDX)
        Debug.Print rIDX(I)
    Next


    MsgBox UBound(rIDX)




    vbDrawCC.SetSourceColor vbRed


    For I = 1 To UBound(rIDX)

        '        PIC.Circle (rX(I) * 1, rY(I) * 1), 2, vbRed

        vbDrawCC.ARC rX(I), rY(I), 5
        vbDrawCC.Fill

    Next
    vbDRAW.Srf.DrawToDC PicHDC


End Sub

Private Sub Command3_Click()

    Const tNP   As Long = 5000
    Dim x()     As Double
    Dim y()     As Double
    Dim I       As Long
    Dim J       As Long
    Dim dx      As Double
    Dim dy      As Double
    Dim R       As Double
    Dim T       As Double
    Dim diam2   As Double
    Dim TT      As Double

    Dim PairsChecked As Long
    Dim CollisionsFound As Long



    R = 2

    diam2 = (R * 2) ^ 2

    ReDim x(tNP)
    ReDim y(tNP)

    For I = 1 To tNP
        x(I) = Rnd * PIC.Width
        y(I) = Rnd * PIC.Height
        Q.InsertSinglePoint x(I), y(I), I
    Next


    Dim rX()    As Double
    Dim rY()    As Double
    Dim rIDX()  As Long


    Do

        vbDrawCC.SetSourceRGB 0.2, 0.3, 0.04
        vbDrawCC.Paint

        Q.DRAW 0


        vbDrawCC.SetSourceColor vbRed, 0.5

        ' Debug.Print
        PairsChecked = 0
        CollisionsFound = 0

        For I = 1 To tNP - 1
            Q.Query x(I) - R * 2, y(I) - R * 2, _
                    x(I) + R * 2, y(I) + R * 2, rX(), rY(), rIDX(), True


            For J = 1 To UBound(rX)
                If rIDX(J) > I Then
                    PairsChecked = PairsChecked + 1
                    dx = x(I) - rX(J)
                    dy = y(I) - rY(J)
                    If dx * dx + dy * dy < diam2 Then
                        CollisionsFound = CollisionsFound + 1
                        'Debug.Print I, rIDX(J)
                        With vbDrawCC
                            .ARC x(I), y(I), R
                            .ARC rX(J), rY(J), R
                            .Fill
                        End With
                    End If
                End If

            Next
            '  Next
        Next

        vbDRAW.Srf.DrawToDC PicHDC
        DoEvents

        Q.Setup 0, 0, MaxW * 1, maxH * 1, 30
        For I = 1 To tNP
            x(I) = x(I) + Rnd * 2 - 1
            y(I) = y(I) + Rnd * 2 - 1
            '    Q.InsertSinglePoint x(I), y(I), I
        Next
        Q.InsertPoints x, y


        TT = Timer
        T = TT - T
        T = 1 / T
        T = Int(T * 1000) * 0.001
        Me.Caption = "FPS: " & T & "    Points:" & tNP & "   PairsChecked:" & PairsChecked & "  Collisions:" & CollisionsFound
        T = TT


    Loop While True

End Sub

Private Sub Form_Load()

    Dim W       As Double
    Dim H       As Double

    W = PIC.Width * 0.5
    H = PIC.Height * 0.5

    Set Q = New cQuadTree

    '    Q.Setup W, H, W, H, 20
    Q.Setup 0, 0, W * 2, H * 2, 40



    InitRC



End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    UnloadRC

End Sub

Private Sub Form_Unload(Cancel As Integer)
    End

End Sub

Private Sub PIC_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button Then
        vbDrawCC.SetSourceRGB 0.5, 0.5, 0.1
        vbDrawCC.Paint

        curridx = curridx + 1

        Q.InsertSinglePoint x * 1, y * 1, curridx

        Q.DRAW -1
        vbDRAW.Srf.DrawToDC PicHDC


    End If

End Sub
