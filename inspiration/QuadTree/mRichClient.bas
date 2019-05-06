Attribute VB_Name = "mRichClient"


Option Explicit


' After Draw ---> REFRESH:
'vbDRAW.Srf.DrawToDC PicHDC
'DoEvents




Public vbDRAW   As cVBDraw
Attribute vbDRAW.VB_VarUserMemId = 1073741826
Public vbDrawCC As cCairoContext
Attribute vbDrawCC.VB_VarUserMemId = 1073741827

Public CONS     As cConstructor
Attribute CONS.VB_VarUserMemId = 1073741828

Public PicHDC   As Long
Attribute PicHDC.VB_VarUserMemId = 1073741829
Public MaxW     As Long
Attribute MaxW.VB_VarUserMemId = 1073741830
Public maxH     As Long
Attribute maxH.VB_VarUserMemId = 1073741831

Public CenX     As Double
Attribute CenX.VB_VarUserMemId = 1073741832
Public CenY     As Double
Attribute CenY.VB_VarUserMemId = 1073741833

Public wMinX    As Double
Attribute wMinX.VB_VarUserMemId = 1073741834
Public wMinY    As Double
Attribute wMinY.VB_VarUserMemId = 1073741835
Public wMaxX    As Double
Attribute wMaxX.VB_VarUserMemId = 1073741836
Public wMaxY    As Double
Attribute wMaxY.VB_VarUserMemId = 1073741837

Public New_c    As cConstructor
Attribute New_c.VB_VarUserMemId = 1073741826
Public Cairo    As cCairo    '<- global defs of the two Main-"EntryPoints" into the RC5
Attribute Cairo.VB_VarUserMemId = 1073741827



Public Sub InitRC()
' Set Srf = Cairo.CreateSurface(400, 400)    'size of our rendering-area in Pixels
' Set CC = Srf.CreateContext    'create a Drawing-Context from the PixelSurface above

    MaxW = fMain.PIC.Width    '- PointSize * 0.5
    maxH = fMain.PIC.Height    '- PointSize * 0.5

    CenX = MaxW * 0.5
    CenY = maxH * 0.5

    '    wMinX = CenX - MaxW * 2.2    'Must be<0
    '    wMinY = CenY - maxH * 2.2
    '    wMaxX = CenX + MaxW * 2.2
    '    wMaxY = CenY + maxH + 2.2


    Set New_c = New cConstructor
    Set Cairo = New_c.Cairo


    Set vbDRAW = Cairo.CreateVBDrawingObject
    '    Set vbDRAW.Srf = Cairo.CreateSurface(400, 400)    'size of our rendering-area in Pixels
    Set vbDRAW.Srf = Cairo.CreateSurface(fMain.PIC.Width, fMain.PIC.Height, ImageSurface)       'size of our rendering-area in Pixels

    Set vbDrawCC = vbDRAW.Srf.CreateContext    'create a Drawing-Context from the PixelSurface above

    'vbDRAW.BindTo frmMain.PIC

    With vbDrawCC
        .AntiAlias = CAIRO_ANTIALIAS_GRAY
        '.CC.SetSourceSurface Srf
        .SetLineCap CAIRO_LINE_CAP_ROUND
        .SetLineJoin CAIRO_LINE_JOIN_ROUND
        .SetLineWidth 1, True
        .SelectFont "Courier New", 9, vbWhite
    End With

    PicHDC = fMain.PIC.hDC
End Sub

Public Sub UnloadRC()
    Set vbDRAW = Nothing

    Set CONS = New cConstructor
    CONS.CleanupRichClientDll
End Sub


