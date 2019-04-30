VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "STD_Applications_Excel_Worksheet"  'stdXLWorksheet
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Initialised as Boolean
Private pWorksheet as Worksheet

Public Function Create(sht as worksheet) as STD_Applications_Excel_Worksheet
  set pWorksheet = sht
End FUnction

Public Property Get Value() as Worksheet
Attribute Value.VB_UserMemID = 0
  set Value = pWorksheet
End Property
Public Property Set Value(obj as Worksheet) 
  set pWorksheet = Value
End Property

Public Property Get ValueRange() as STD_Applications_Excel_
    'Get used range
  Dim ur As Range
  Set ur = sht.UsedRange

  'If used range is 1x1 then result is 1x1
  If ur.Rows.Count = 1 And ur.Columns.Count = 1 Then
    Set ValueRange = ur
    Exit Function
  End If

  'Find via array
  'Get array of all values:
  Dim v As Variant
  v = ur.Value

  'Offsets if they exist
  Dim offR, offC As Long
  With ur
    offR = .row - 1
    offC = .Column - 1
  End With

  'Define required values
  Dim colMin, colMax, rowMin, rowMax, row, col As Long

  'Find min row:
  Dim ubndR, ubndC, lbndR, lbndC As Long
  lbndR = LBound(v, 1)
  lbndC = LBound(v, 2)
  ubndR = UBound(v, 1)
  ubndC = UBound(v, 2)

  For row = lbndR To ubndR
    For col = lbndC To ubndC
      If Not IsEmpty(v(row, col)) Then
        rowMin = row
        GoTo NextNum
      End If
    Next
  Next
NextNum:
  'Find max row
  For row = ubndR To lbndR Step -1
    For col = lbndC To ubndC
      If Not IsEmpty(v(row, col)) Then
        rowMax = row
        GoTo NextNum2
      End If
    Next
  Next
NextNum2:
  'Find min col:
  For col = lbndC To ubndC
    For row = lbndR To ubndR
      If Not IsEmpty(v(row, col)) Then
        colMin = col
        GoTo NextNum3
      End If
    Next

  Next
NextNum3:
  'Find max col
  For col = ubndC To lbndC Step -1
    For row = lbndR To ubndR
      If Not IsEmpty(v(row, col)) Then
        colMax = col
        GoTo NextNum4
      End If
    Next
  Next
NextNum4:
  Set ValueRange = Range(sht.Cells(offR + rowMin, offC + colMin), sht.Cells(offR + rowMax, offC + colMax))
End Property