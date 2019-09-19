VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMultiEnumerations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbIEnumerable 'a Class-Interface from vbInterfaces.dll, which encapsulates IEnumVariant quite similar to
                       'what was shown in Tutorial-Folder #4 - but compared to a full IEnumVariant-Implementation
                       'requiring much less "noise" in the implementing Class

'So, this class makes use of that more convenient Helper, to offer Enumerations for three internal "things":
'- the LngArr(1 to 3)
'- the StrArr(0 to 2)
'- the Members of the eMultiEnumeration Enum-Type
'all three directly defined below

Private Enum eMultiEnumeration
  EnmLngArr
  EnmStrArr
  EnmMultiEnumeration
End Enum

Private LngArr(1 To 3) As Long, StrArr(0 To 2) As String
 
Private Sub Class_Initialize() 'fill in some Demo-Data into our (later enumerated) Arrays
  LngArr(1) = 1: LngArr(2) = 2: LngArr(3) = 3
  StrArr(0) = "S0": StrArr(1) = "S1": StrArr(2) = "S2"
End Sub
 
Public Function EnumerateLngArr() As Object
  With New vbIEnumerable
    Set EnumerateLngArr = .EnumerateOn(Me, 1, eMultiEnumeration.EnmLngArr)
  End With
End Function
 
Public Function EnumerateStrArr() As Object
  With New vbIEnumerable
    Set EnumerateStrArr = .EnumerateOn(Me, 0, eMultiEnumeration.EnmStrArr)
  End With
End Function
 
Public Function EnumerateEnumType() As Object
  With New vbIEnumerable
    Set EnumerateEnumType = .EnumerateOn(Me, 0, eMultiEnumeration.EnmMultiEnumeration)
  End With
End Function

'**** implementation of the IEnumerable-Helper callback-routine
Private Function vbIEnumerable_ElementForIndex(ByVal Index As Long, UserData As Variant, CancelEnumeration As Boolean) As Variant
  Select Case UserData
    Case eMultiEnumeration.EnmLngArr
      If Index > UBound(LngArr) Then CancelEnumeration = True Else vbIEnumerable_ElementForIndex = LngArr(Index)
    
    Case eMultiEnumeration.EnmStrArr
      If Index > UBound(StrArr) Then CancelEnumeration = True Else vbIEnumerable_ElementForIndex = StrArr(Index)
    
    Case eMultiEnumeration.EnmMultiEnumeration
      If Index > eMultiEnumeration.EnmMultiEnumeration Then CancelEnumeration = True: Exit Function
      Select Case Index
        Case eMultiEnumeration.EnmLngArr:           vbIEnumerable_ElementForIndex = "EnmLngArr"
        Case eMultiEnumeration.EnmStrArr:           vbIEnumerable_ElementForIndex = "EnmStrArr"
        Case eMultiEnumeration.EnmMultiEnumeration: vbIEnumerable_ElementForIndex = "EnmMultiEnumeration"
      End Select
  End Select
End Function
'never needed (just there, to cover it all in one single Helper-Class: IEnumerable)
Private Function vbIEnumerable_EnumerateOn(oCallback As vbIEnumerable, ByVal StartIndex As Long, UserData As Variant) As Object
End Function
