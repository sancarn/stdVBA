Attribute VB_Name = "vbCOMHelpers"
Option Explicit

Public Type tIID
  IID(0 To 15) As Byte
End Type

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function EbMode Lib "vba6" () As Long
Private Declare Function EbIsResetting Lib "vba6" () As Long

Declare Function lstrcpynW& Lib "kernel32" (ByVal lpDst&, ByVal lpSrc&, ByVal MaxLength&)
Declare Sub Assign Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, Optional ByVal CB& = 4)
Declare Sub AssignZero Lib "kernel32" Alias "RtlZeroMemory" (pDst As Any, Optional ByVal CB& = 4)
Declare Sub BindArray Lib "kernel32" Alias "RtlMoveMemory" (PArr() As Any, pSrc&, Optional ByVal CB& = 4)
Declare Sub ReleaseArray Lib "kernel32" Alias "RtlMoveMemory" (PArr() As Any, Optional pSrc& = 0, Optional ByVal CB& = 4)
Declare Function VariantCopyToPtrAPI& Lib "oleaut32" Alias "VariantCopy" (ByVal pDstVariant As Long, Src As Variant)
Declare Function StringFromGUID2& Lib "ole32" (IID As Any, ByVal lpStr&, Optional ByVal cchmax& = 39)
Declare Function IsEqualGUID Lib "ole32" (Guid1 As Any, Guid2 As Any) As Long
Declare Function CLSIDFromString& Lib "ole32" (ByVal lpStr&, IID As Any)
Declare Function CoTaskMemAlloc& Lib "ole32" (ByVal sz&)
Declare Sub CoTaskMemFree Lib "ole32" (ByVal pMem&)

Public vbI As cInterfaces, IID_IUnknown As tIID

Sub Main()
  Set vbI = New cInterfaces
  IID_IUnknown = STRtoIID(vbI.sIID_IUnknown)
End Sub
 
Public Function IIDtoStr(IID As tIID) As String
  IIDtoStr = Space$(38)
  StringFromGUID2 IID, StrPtr(IIDtoStr)
End Function

Public Function STRtoIID(STR As String) As tIID
  If Len(STR) <> 38 Then Err.Raise vbObjectError, , "Invalid IID-String"
  CLSIDFromString StrPtr(STR), STRtoIID
End Function
 
Public Function InVBAStopModeOrResetting() As Boolean
  Static InitDone As Boolean, VBAEnv As Boolean
  If Not InitDone Then
    InitDone = True
    VBAEnv = GetModuleHandle("vba6.dll")
  End If
  If VBAEnv Then InVBAStopModeOrResetting = (EbMode = 0 Or EbIsResetting <> 0)
End Function

Public Function GetStringFromPointerW(ByVal WStrPtr As Long, Optional ByVal ExpectedMaxLen As Long = 1024) As String
Static sBuffer As String, psBuffer As Long, NullCharPos As Long
  If WStrPtr = 0 Then Exit Function
  If Len(sBuffer) < ExpectedMaxLen Then
    sBuffer = String$(ExpectedMaxLen, 0)
    psBuffer = StrPtr(sBuffer)
  End If
  
  If lstrcpynW(psBuffer, WStrPtr, ExpectedMaxLen) Then
    NullCharPos = InStr(sBuffer, vbNullChar) - 1
    If NullCharPos <= 0 Then NullCharPos = ExpectedMaxLen
    GetStringFromPointerW = Left$(sBuffer, NullCharPos)
  End If
End Function
