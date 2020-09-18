 
'Modified from src: https://www.vbforums.com/showthread.php?883497-Simple-Encryption

Private Type Blob
    Size As Long
    Data As Long
End Type

Dim EncrData As Blob

Private Declare Sub LocalFree Lib "kernel32.dll" (ByVal M As Long)
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CryptProtectData Lib "Crypt32.dll" (pDataIn As Blob, ByVal szDataDescr As Long, ByVal Entrophy As Long, ByVal pvReserved As Long, ByVal pPromptStruct As Long, ByVal dwFlags As Long, pDataOut As Blob) As Long
Private Declare Function CryptUnprotectData Lib "Crypt32.dll" (pDataIn As Blob, ByVal szDataDescr As Long, ByVal Entrophy As Long, ByVal pvReserved As Long, ByVal pPromptStruct As Long, ByVal dwFlags As Long, pDataOut As Blob) As Long
 'CryptProtectData(DataIn, null, DataPassword, null, null,null, DataOut)

Private Function Encrypt(Source As String) As Long
    Dim DataIn As Blob
    Dim lRet As Long
    DataIn.Size = LenB(Source)
    DataIn.Data = StrPtr(Source)
    lRet = CryptProtectData(DataIn, 0&, 0, 0, 0, 0, EncrData)
    If lRet = 0 Then Encrypt = Err.LastDllError
End Function
Private Function Decrypt(sDecr As String) As Boolean
    Dim DataTmp As Blob
    Call CryptUnprotectData(EncrData, 0, 0, 0, 0, 0, DataTmp) 'Decrypt data
    sDecr = String$(DataTmp.Size / 2, Chr$(0)) 'Prepare string & copy data
    CopyMemory ByVal StrPtr(sDecr), ByVal DataTmp.Data, DataTmp.Size
    LocalFree DataTmp.Data
End Function

Private Sub BytesToHex(sDescr As String, bArray() As Byte) as string
    Dim lPtr As Long
    If GetbSize(bArray) = 0 Then Exit Sub
    Dim s as string
    For lPtr = 0 To UBound(bArray)
        s = s &  Right$("0" & Hex$(bArray(lPtr)), 2) & " ";
    Next lPtr
    BytesToHex = s
End Sub
Private Function BytesToStr(sDescr As String, bArray() As Byte) as string
    Dim lPntr As Long
    Dim strBuff As String
    Dim iLines As Integer
    Dim M%, N%
    lPntr = GetbSize(bArray)
    If lPntr = 0 Then Exit Sub
    If lPntr Mod 16 = 0 Then
        iLines = lPntr / 16
    Else
        iLines = Int(lPntr / 16) + 1
    End If
    strBuff = String$(iLines * 50, Chr$(32))
    Mid$(strBuff, Len(strBuff) - 1, 2) = vbCrLf
    M% = 1
    For N% = 0 To UBound(bArray) Step 1
        Mid$(strBuff, M%, 3) = Right$("0" & Hex$(bArray(N%)), 2) & " "
        M% = M% + 3
        If (N% + 1) Mod 16 = 0 Then
            Mid$(strBuff, M%, 2) = vbCrLf
            M% = M% + 2
        End If
    Next N%
    If Len(sDescr) > 0 Then sDescr = sDescr & ":" & vbCrLf
    BytesToStr = sDescr & strBuff
End Function



 
Private Function GetbSize(bArray() As Byte) As Long
    On Error GoTo GetSizeErr
    GetbSize = UBound(bArray) + 1
    Exit Function
GetSizeErr:
    GetbSize = 0
End Function



Private Sub test()
    'Encrypt
    Dim sIN As String
    Dim bTmp() As Byte
    If Encrypt(txtString.Text) = 0 Then
        ReDim bTmp(EncrData.Size - 1)
        CopyMemory bTmp(0), ByVal EncrData.Data, EncrData.Size
        Debug.Print BytesToHex("Encrypted Data", bTmp)
    Else
        Debug.Print "Encryption Failed!"
    End If

    Debug.Assert false

    'Decrypt
    Dim sOut As String
    Call Decrypt(sOut)
    txtMsg.Text = txtMsg.Text & "Decrypted: " & sOut
    txtMsg.SelStart = Len(txtMsg.Text)

    Debug.assert false
End Sub