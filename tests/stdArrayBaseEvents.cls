VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "stdArrayBaseEvents"
Public WithEvents Values As stdArray
Public BeforeIndex As Long
Public AfterIndex As Long
Private Sub Values_BeforeAdd(ByRef arr As stdArray, ByVal iIndex As Long, ByRef item As Variant, ByRef cancel As Boolean)
  BeforeIndex = iIndex
End Sub
Private Sub Values_AfterAdd(ByRef arr As stdArray, ByVal iIndex As Long, ByRef item As Variant)
  AfterIndex = iIndex
End Sub
Private Sub Values_BeforeRemove(ByRef arr As stdArray, ByVal iIndex As Long, ByRef item As Variant, ByRef cancel As Boolean)
  BeforeIndex = iIndex
End Sub
Private Sub Values_AfterRemove(ByRef arr As stdArray, ByVal iIndex As Long)
  AfterIndex = iIndex
End Sub

