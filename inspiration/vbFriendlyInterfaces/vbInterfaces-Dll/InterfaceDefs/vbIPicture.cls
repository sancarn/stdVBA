VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbIPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GetHandle(UserData As Long) As Long
End Function
Public Function GetHPal(UserData As Long) As Long
End Function
Public Function GetPictureType(UserData As Long) As PictureTypeConstants
End Function
Public Function GetWidth(UserData As Long) As Long
End Function
Public Function GetHeight(UserData As Long) As Long
End Function
Public Sub Render(UserData As Long, ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, _
                  ByVal xSrc As Long, ByVal ySrc As Long, ByVal cxSrc As Long, ByVal cySrc As Long, ByVal pRcBounds As Long)
End Sub
Public Sub SetHPal(UserData As Long, ByVal NewHPal As Long)
End Sub
Public Function GetCurDC(UserData As Long) As Long
End Function
Public Sub SelectPicture(UserData As Long, ByVal hDCToSelectInto As Long, hDCPrevious As Long, hBmp As Long)
End Sub
Public Sub PictureChanged(UserData As Long)
End Sub
Public Sub SaveAsFile(UserData As Long, ByVal pStm As Long, ByVal fSaveMemCopy As Boolean, SavedBytes As Long)
End Sub
Public Function GetAttributes(UserData As Long) As Long
End Function

