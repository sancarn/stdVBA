VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cQuadTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type tRect
    x1          As Double
    y1          As Double
    x2          As Double
    y2          As Double
End Type

Private Boundary As tRect

Private ChildNW As cQuadTree
Private ChildNE As cQuadTree
Private ChildSW As cQuadTree
Private ChildSE As cQuadTree

Private Capacity As Long

Private NP      As Long

Private pX()    As Double
Private pY()    As Double
Private pIDX()  As Long


Public Divided  As Boolean


Private Sub SubDivide()
    Dim cnX     As Double
    Dim cnY     As Double

    If ChildNW Is Nothing Then Set ChildNW = New cQuadTree
    If ChildNE Is Nothing Then Set ChildNE = New cQuadTree
    If ChildSW Is Nothing Then Set ChildSW = New cQuadTree
    If ChildSE Is Nothing Then Set ChildSE = New cQuadTree


    With Boundary
        cnX = (.x2 + .x1) * 0.5
        cnY = (.y2 + .y1) * 0.5

        ChildNW.Setup .x1, .y1, cnX, cnY, Capacity
        ChildNE.Setup cnX, .y1, .x2, cnY, Capacity
        ChildSW.Setup .x1, cnY, cnX, .y2, Capacity
        ChildSE.Setup cnX, cnY, .x2, .y2, Capacity

    End With

    Divided = True



End Sub


Friend Sub Setup(x1 As Double, y1 As Double, x2 As Double, y2 As Double, Capac As Long)

    If Capac Then
        With Boundary
            .x1 = x1
            .y1 = y1
            .x2 = x2
            .y2 = y2
        End With
        Capacity = Capac


        ReDim pX(Capacity)
        ReDim pY(Capacity)
        ReDim pIDX(Capacity)
    End If


    ''    If Divided Then
    ''        ChildNE.Setup 0, 0, 0, 0, 0
    ''        ChildNW.Setup 0, 0, 0, 0, 0
    ''        ChildSE.Setup 0, 0, 0, 0, 0
    ''        ChildSW.Setup 0, 0, 0, 0, 0
    ''        Set ChildNE = Nothing
    ''        Set ChildNW = Nothing
    ''        Set ChildSE = Nothing
    ''        Set ChildSW = Nothing
    ''    End If

    Divided = False
    NP = 0

End Sub



Friend Function InsertSinglePoint(pointX As Double, pointY As Double, pointIDX As Long) As Boolean

    If Not (BoundaryContainPoint(pointX, pointY)) Then Exit Function

    If NP < Capacity Then
        NP = NP + 1
        pX(NP) = pointX
        pY(NP) = pointY
        pIDX(NP) = pointIDX
        InsertSinglePoint = True

    Else
        If Not (Divided) Then SubDivide

        If ChildNW.InsertSinglePoint(pointX, pointY, pointIDX) Then InsertSinglePoint = True: Exit Function
        If ChildNE.InsertSinglePoint(pointX, pointY, pointIDX) Then InsertSinglePoint = True: Exit Function
        If ChildSW.InsertSinglePoint(pointX, pointY, pointIDX) Then InsertSinglePoint = True: Exit Function
        If ChildSE.InsertSinglePoint(pointX, pointY, pointIDX) Then InsertSinglePoint = True: Exit Function

    End If

End Function

Friend Sub InsertPoints(pointX() As Double, pointY() As Double)
    Dim I       As Long
    Dim N       As Long
    Dim pntX    As Double
    Dim pntY    As Double

    N = UBound(pointX)
    For I = 1 To N

        pntX = pointX(I)
        pntY = pointY(I)

        If (BoundaryContainPoint(pntX, pntY)) Then
            If NP < Capacity Then
                NP = NP + 1
                pX(NP) = pntX
                pY(NP) = pntY
                pIDX(NP) = I

            Else
                If Not (Divided) Then SubDivide

                If Not (ChildNW.InsertSinglePoint(pntX, pntY, I)) Then _
                   If Not (ChildNE.InsertSinglePoint(pntX, pntY, I)) Then _
                   If Not (ChildSW.InsertSinglePoint(pntX, pntY, I)) Then _
                   ChildSE.InsertSinglePoint pntX, pntY, I

            End If
        End If
    Next

End Sub

Private Function BoundaryContainPoint(x As Double, y As Double) As Boolean

    With Boundary
        If x > .x2 Then Exit Function
        If x < .x1 Then Exit Function
        If y > .y2 Then Exit Function
        If y < .y1 Then Exit Function
    End With
    BoundaryContainPoint = True

End Function


Friend Sub DRAW(ShowQuads As Long)
    Dim I       As Long
    Dim x       As Double
    Dim y       As Double

    If NP = 0 Then Exit Sub

    If ShowQuads Then
        '    'DrawCross
        With Boundary
            x = (.x2 + .x1) * 0.5
            y = (.y2 + .y1) * 0.5
            vbDrawCC.DrawLine x, .y1, x, .y2, , 1, vbWhite, 0.2
            vbDrawCC.DrawLine .x1, y, .x2, y, , 1, vbWhite, 0.2
        End With
    End If

    'DrawPoints
    vbDrawCC.SetSourceColor vbYellow, 0.8
    For I = 1 To NP
        vbDrawCC.ARC pX(I), pY(I), 2
        vbDrawCC.Fill
    Next


    If Divided Then
        ChildNW.DRAW ShowQuads
        ChildNE.DRAW ShowQuads
        ChildSW.DRAW ShowQuads
        ChildSE.DRAW ShowQuads
    End If


End Sub





Private Function BBOverlap(BB1minX As Double, BB1minY As Double, _
                           BB1maxX As Double, BB1maxY As Double, _
                           BB2minX As Double, BB2minY As Double, _
                           BB2maxX As Double, BB2maxY As Double) As Boolean
    If BB1maxX < BB2minX Then Exit Function
    If BB1maxY < BB2minY Then Exit Function
    If BB1minX > BB2maxX Then Exit Function
    If BB1minY > BB2maxY Then Exit Function
    BBOverlap = True

End Function



Friend Sub Query(Rx1 As Double, _
                 Ry1 As Double, _
                 Rx2 As Double, _
                 Ry2 As Double, rpX() As Double, rpY() As Double, rpIDX() As Long, Optional isRoot As Boolean = False)

'Returns all points in rectamgle Rx1.Ry1    Rx2.Ry2

    Dim I       As Long

    Dim rnp     As Long
    Dim MaxArr  As Long


    If Not (BBOverlap(Rx1, Ry1, _
                      Rx2, Ry2, _
                      Boundary.x1, Boundary.y1, _
                      Boundary.x2, Boundary.y2)) Then Exit Sub


    If isRoot Then
        ReDim rpX(0)
        ReDim rpY(0)
        ReDim rpIDX(0)
    End If

    rnp = UBound(rpX)
    MaxArr = rnp

    For I = 1 To NP    'Point in Range ?
        If pX(I) >= Rx1 Then
            If pX(I) <= Rx2 Then
                If pY(I) >= Ry1 Then
                    If pY(I) <= Ry2 Then
                        rnp = rnp + 1
                        If rnp > MaxArr Then
                            MaxArr = MaxArr + 64
                            ReDim Preserve rpX(MaxArr)
                            ReDim Preserve rpY(MaxArr)
                            ReDim Preserve rpIDX(MaxArr)
                        End If

                        rpX(rnp) = pX(I)
                        rpY(rnp) = pY(I)
                        rpIDX(rnp) = pIDX(I)
                    End If
                End If
            End If
        End If
    Next

    If rnp < MaxArr Then
        ReDim Preserve rpX(rnp)
        ReDim Preserve rpY(rnp)
        ReDim Preserve rpIDX(rnp)
    End If

    If Divided Then
        ChildNW.Query Rx1, Ry1, Rx2, Ry2, rpX(), rpY(), rpIDX()
        ChildNE.Query Rx1, Ry1, Rx2, Ry2, rpX(), rpY(), rpIDX()
        ChildSW.Query Rx1, Ry1, Rx2, Ry2, rpX(), rpY(), rpIDX()
        ChildSE.Query Rx1, Ry1, Rx2, Ry2, rpX(), rpY(), rpIDX()
    End If



End Sub






'********************************************************************
'Query for 2D collision:

''Q.Setup 0, 0, MaxW * 1, maxH * 1, 30
'''For I = 1 To NP
'''    Q.InsertSinglePoint x(I), y(I), I
'''Next
''Q.inserpoints x, y
''
''
''
''For I = 1 To NP
''    Q.Query x(I) - R * 2, y(I) - R * 2, _
 ''            x(I) + R * 2, y(I) + R * 2, rX(), rY(), rIDX(), True
''
''    For J = 1 To UBound(rX)
''        If I < rIDX(J) Then
''            dx = x(I) - rX(J)
''            dy = y(I) - rY(J)
''            If dx * dx + dy * dy < diam2 Then  'Diam*Diam   (2*R)^2
''                'COLLISION HAPPENED
''            End If
''        End If
''
''    Next
''Next


