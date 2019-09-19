Attribute VB_Name = "modMain"
Option Explicit

Public vbI As New vbInterfaces.cInterfaces  'this Object doesn't hold any "state" - so we make it globally available
Public GDIPlusCache As New cGDIPlusCache    'we store all the Resources once (under a StringKey) in this global Container

Private Sub Main()
  GDIPlusCache.AddImage "ChkSmall", App.Path & "\Res\ChkSmall.png"
  GDIPlusCache.AddImage "Gif1_55x55", App.Path & "\Res\Earth.gif", 32, 32, True
  
  GDIPlusCache.AddImage "Png1", App.Path & "\Res\Png1.png"
  GDIPlusCache.AddImage "Png2_31x31", App.Path & "\Res\Png2.png", 31, 31 'a 31x31 quality-scale-down from a 128x128 Png-File for a CommandButton-Picture
  GDIPlusCache.AddImage "Png3_16x16", App.Path & "\Res\Png3.png", 16, 16 'same here, a downscaled, third PNG-resource will be used for the Form-Icon later
  GDIPlusCache.AddImage "Png3_29x29", App.Path & "\Res\Png3.png", 29, 29 '29x29 from the same Png-File for another CommandButton-Picture
  GDIPlusCache.AddImage "Png3_128x128", App.Path & "\Res\Png3.png" 'without specifying a desired size, the original Size (128x128) of the PNG is used
  
  GDIPlusCache.AddIcon "Ico1", App.Path & "\Res\XP-alpha.ico", 48, 48 '<- that size is the maximum available size in that Icon-File
  GDIPlusCache.AddIcon "Ico2", App.Path & "\Res\Vista-Alpha+Png.ico", 256, 256 '<- that larger desired size will retrieve the Icons PNG-content
  GDIPlusCache.AddIcon "Ico3", App.Path & "\Res\Vista-Alpha+Png.ico", 32, 32 'cache only the smaller 32x32 Alpha-Icon from the same *.ico File
  
  GDIPlusCache.AddImage "HandCursor", App.Path & "\Res\BlueHandCursor.png", 31, 31 'this PNG-resource will be used as a Hand-MouseCursor on the Form
  
  fTest.Show
End Sub

'the globally available Function below is using the cPictureEx-Class, to provide a better alternative to VBs LoadPicture-Function
Public Function LoadPictureEx(ImgKey As String, Optional ByVal PictureType As PictureTypeConstants = vbPicTypeBitmap, _
                                                Optional ByVal Attributes As RenderAttributes = PICTURE_TRANSPARENT, _
                                                Optional ByVal HighStretchQuality As Boolean, Optional ByVal Alpha As Double = 1) As StdPicture
  Dim Pic As New cPictureEx
  Set LoadPictureEx = Pic.Picture(ImgKey, PictureType, Attributes, HighStretchQuality, Alpha)
End Function
