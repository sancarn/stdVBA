VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "String Evaluation"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Evaluate"
      Height          =   375
      Left            =   8040
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   9855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sText     As String
Dim sTextOrig As String
Dim sARY(2)   As String
Dim bText     As Byte
Dim lPtr      As Long
Dim lPtrOrig  As Long
Dim sPtr      As Long
Dim sPtrOrig  As Long
Dim pPtr      As Long
Dim pPtrOrig  As Long
Dim sLen      As Long
Dim pVal      As Long
Dim pValOrig  As Long
Dim i         As Integer
Dim vStr      As Variant

Private Type dcluStr
             Str0 As String
             Str1 As String
             Str2 As String
        End Type

Private Declare Sub CopyMemory Lib "KERNEL32" _
                    Alias "RtlMoveMemory" (hpvDest As Any, _
                                           hpvSource As Any, _
                                           ByVal cbCopy As Long)

Private Sub Command1_Click()
    
  If Text2.Text = "" Then
     MsgBox "Nothing to evaluate!", vbInformation, Me.Caption
     Text2.SetFocus
     Exit Sub
  End If
  
  If Command1.Caption = "Evaluate" Then
     Command1.Caption = "More..."
     Step1
  Else
     If Command1.Caption = "More..." Then
        Text2.Locked = True
        Command1.Caption = "Continue"
        Step2
     Else
        If Command1.Caption = "Continue" Then
           Command1.Caption = "Finally..."
           Step3
        Else
           If Command1.Caption = "Finally..." Then
              Command1.Caption = "Finished"
              Step4
           Else
              Step5
              Text2.Locked = False
              Text2.SetFocus
              Command1.Caption = "Evaluate"
           End If
        End If
     End If
  End If
  
End Sub

Private Sub Step1()
  
  sText = Text2.Text
  sPtr = StrPtr(sText)
  pPtr = VarPtr(sText)
  lPtr = sPtr - 4
  CopyMemory sLen, ByVal lPtr, 4
  CopyMemory pVal, ByVal pPtr, 4
  Text1.Text = "You entered: " & sText & vbCrLf & vbCrLf & _
               "The text was stored in the VB string ""sText""." & vbCrLf & _
               "The address of the BSTR [VarPtr(sText)] is: " & pPtr & vbCrLf & _
               "The address in the BSTR [StrPtr(sText)] is: " & pVal & vbCrLf & _
               "The address of the length [StrPtr(sText) - 4] is: " & lPtr & vbCrLf & vbCrLf & _
               "The stored length is: " & sLen & vbCrLf & _
               "LenB(sText) returns " & LenB(sText) & vbCrLf & _
               "Len(sText) returns " & Len(sText) & vbCrLf & vbCrLf & _
               "The contents of memory at " & lPtr & "-" & (lPtr + 3) & " is: "

  For i = 0 To 3
      CopyMemory bText, ByVal lPtr + i, 1
      Text1.Text = Text1.Text & bText & " "
  Next i
  Text1.Text = Text1.Text & vbCrLf & "NOTE: the value is stored in little-endian format!" & _
               vbCrLf & vbCrLf & "The contents of memory at " & _
               sPtr & "-" & (sPtr + sLen + 1) & " is: " & vbCrLf
  For i = 0 To sLen + 1
      CopyMemory bText, ByVal sPtr + i, 1
      Text1.Text = Text1.Text & bText & " "
  Next i
  Text1.Text = Text1.Text & vbCrLf & "NOTE: this includes the NULL terminator character!" & _
               vbCrLf & vbCrLf & "Enter another string and click 'More...'"
  
End Sub

Private Sub Step2()

  sTextOrig = sText     ' save the old string info
  sPtrOrig = sPtr
  pPtrOrig = pPtr
  lPtrOrig = lPtr
  pValOrig = pVal
  sText = Text2.Text    ' and get the new string
  sPtr = StrPtr(sText)
  pPtr = VarPtr(sText)
  lPtr = sPtr - 4
  CopyMemory sLen, ByVal lPtr, 4
  CopyMemory pVal, ByVal pPtr, 4
  Text1.Text = "You entered: " & sText & vbCrLf & vbCrLf & _
               "The text was stored in the VB string ""sText""." & vbCrLf & _
               "The address of the BSTR [VarPtr(sText)] is: " & pPtr & vbCrLf & _
               "The address in the BSTR [StrPtr(sText)] is: " & pVal & vbCrLf & _
               "The address of the length [StrPtr(sText) - 4] is: " & lPtr & vbCrLf & vbCrLf & _
               "Notice that the address of the BSTR [VarPtr(sText)] is still " & pPtrOrig & ", while the " & _
               "address in the BSTR [StrPtr(sText)] has changed from it's original value of " & pValOrig & ".  " & _
               "In fact, if we append the original string to the new string, the address in the BSTR for " & _
               "the resulting string should be yet another value." & vbCrLf & vbCrLf
  sText = sText & " " & sTextOrig
  Text1.Text = Text1.Text & "sText now contains " & sText & ".  The address of the BSTR [VarPtr(sText)] is still " & _
               VarPtr(sText) & ", but the address in the BSTR [StrPtr(sText)] is now " & StrPtr(sText) & "." & vbCrLf & vbCrLf & _
               "Click 'Continue' to look at string arrays."

End Sub

Private Sub Step3()
  Dim aVal(2) As Long
  Dim aLen(2) As Long
  
  ' get the first 3 words from the string
  vStr = Split(sTextOrig, " ")
  sARY(0) = vStr(0)
  sARY(1) = vStr(1)
  sARY(2) = vStr(2)
  'get the addresses in their BSTRs and the values of their lengths
  sPtr = VarPtr(sARY(0))
  CopyMemory aVal(0), ByVal sPtr, 4
  CopyMemory aLen(0), ByVal aVal(0) - 4, 4
  sPtr = sPtr + 4
  CopyMemory aVal(1), ByVal sPtr, 4
  CopyMemory aLen(1), ByVal aVal(1) - 4, 4
  sPtr = sPtr + 4
  CopyMemory aVal(2), ByVal sPtr, 4
  CopyMemory aLen(2), ByVal aVal(2) - 4, 4
  
  Text1.Text = "The original first 3 words entered have been stored in the string array sARY.  sARY is " & _
               "actually an array of 3 BSTRs, or Longs containing the memory adddresses " & _
               "of the actual strings." & vbCrLf & vbCrLf & _
               "The BSTR sARY(0) is located at " & VarPtr(sARY(0)) & " and contains the address " & aVal(0) & vbCrLf & _
               "The BSTR sARY(1) is located at " & VarPtr(sARY(1)) & " and contains the address " & aVal(1) & vbCrLf & _
               "The BSTR sARY(2) is located at " & VarPtr(sARY(2)) & " and contains the address " & aVal(2) & vbCrLf & _
               "Notice that the 3 BSTRs are adjacent to each other in memory but the addresses that they " & _
               "contain are scattered throughout memory." & vbCrLf & vbCrLf & _
               "The string at address " & aVal(0) & " has a LenB of " & aLen(0) & _
               " bytes" & vbCrLf & "The contents of memory at " & aVal(0) & " is: "
  For i = 0 To aLen(0) + 1
      CopyMemory bText, ByVal aVal(0) + i, 1
      Text1.Text = Text1.Text & bText & " "
  Next i
  Text1.Text = Text1.Text & vbCrLf & vbCrLf
  
  Text1.Text = Text1.Text & "The string at address " & aVal(1) & " has a LenB of " & aLen(1) & " bytes" & vbCrLf & _
               "The contents of memory at " & aVal(1) & " is: "
  For i = 0 To aLen(1) + 1
      CopyMemory bText, ByVal aVal(1) + i, 1
      Text1.Text = Text1.Text & bText & " "
  Next i
  Text1.Text = Text1.Text & vbCrLf & vbCrLf

  Text1.Text = Text1.Text & "The string at address " & aVal(2) & " has a LenB of " & aLen(2) & " bytes" & vbCrLf & _
               "The contents of memory at " & aVal(2) & " is: "
  For i = 0 To aLen(2) + 1
      CopyMemory bText, ByVal aVal(2) + i, 1
      Text1.Text = Text1.Text & bText & " "
  Next i
  Text1.Text = Text1.Text & vbCrLf & vbCrLf & "Click 'Finally...' to look at strings in a UDT"
  
End Sub

Private Sub Step4()
  Dim uVal(2) As Long
  Dim uLen(2) As Long
  Dim uSTR    As dcluStr
  
  With uSTR
       ' get the first 3 words from the string
       .Str0 = vStr(0)
       .Str1 = vStr(1)
       .Str2 = vStr(2)
       'get the addresses in their BSTRs and the values of their lengths
       sPtr = VarPtr(.Str0)
       CopyMemory uVal(0), ByVal sPtr, 4
       CopyMemory uLen(0), ByVal uVal(0) - 4, 4
       sPtr = sPtr + 4
       CopyMemory uVal(1), ByVal sPtr, 4
       CopyMemory uLen(1), ByVal uVal(1) - 4, 4
       sPtr = sPtr + 4
       CopyMemory uVal(2), ByVal sPtr, 4
       CopyMemory uLen(2), ByVal uVal(2) - 4, 4
       End With
   Text1.Text = "The first 3 words entered have now been stored in the UDT uSTR, defined as:" & vbCrLf & _
                "Private Type dcluStr" & vbCrLf & _
                "             Str0 As String" & vbCrLf & _
                "             Str1 As String" & vbCrLf & _
                "             Str2 As String" & vbCrLf & _
                "        End Type" & vbCrLf & vbCrLf & _
                "Something that we can do with a UDT that we could not do with an array is get " & _
                "its Len and LenB.  Since Str0, Str1 and Str2 are BSTRs, Len and LenB should both " & _
                "be equal to the number of bytes in 3 Longs, i.e. 12!  Len(uSTR) = " & Len(uSTR) & _
                " and LenB(uSTR) = " & LenB(uSTR) & vbCrLf & vbCrLf & _
                "Like the array sARY in the previous example, uSTR contains nothing more than 3 " & vbCrLf & _
                "BSTRs." & vbCrLf & vbCrLf & _
                "The BSTR Str0 is located at " & VarPtr(uSTR.Str0) & " and contains the address " & uVal(0) & vbCrLf & _
                "The BSTR Str1 is located at " & VarPtr(uSTR.Str1) & " and contains the address " & uVal(1) & vbCrLf & _
                "The BSTR Str2 is located at " & VarPtr(uSTR.Str2) & " and contains the address " & uVal(2) & vbCrLf & _
                "Notice that, as with aSTR, the 3 BSTRs are adjacent to each other in memory but the " & _
                "addresses that they contain are scattered throughout memory." & vbCrLf & vbCrLf & _
                "The string at address " & uVal(0) & " has a LenB of " & uLen(0) & _
                " bytes" & vbCrLf & "The contents of memory at " & uVal(0) & " is: "
  For i = 0 To uLen(0) + 1
      CopyMemory bText, ByVal uVal(0) + i, 1
      Text1.Text = Text1.Text & bText & " "
  Next i
  Text1.Text = Text1.Text & vbCrLf & vbCrLf
  
  Text1.Text = Text1.Text & "The string at address " & uVal(1) & " has a LenB of " & uLen(1) & " bytes" & vbCrLf & _
               "The contents of memory at " & uVal(1) & " is: "
  For i = 0 To uLen(1) + 1
      CopyMemory bText, ByVal uVal(1) + i, 1
      Text1.Text = Text1.Text & bText & " "
  Next i
  Text1.Text = Text1.Text & vbCrLf & vbCrLf

  Text1.Text = Text1.Text & "The string at address " & uVal(2) & " has a LenB of " & uLen(2) & " bytes" & vbCrLf & _
               "The contents of memory at " & uVal(2) & " is: "
  For i = 0 To uLen(2) + 1
      CopyMemory bText, ByVal uVal(2) + i, 1
      Text1.Text = Text1.Text & bText & " "
  Next i
  Text1.Text = Text1.Text & vbCrLf & vbCrLf & "Click 'Finished' to continue."


End Sub

Private Sub Step5()

  Text1.Text = "Enter some text in the field above and click 'Evaluate'."
  Text2.Text = ""
  
End Sub

Private Sub Form_Activate()

  Text2.SetFocus
  
End Sub

Private Sub Form_Load()

  Text1.Text = "Dim myString As String" & vbCrLf & "myString = ""VB strings are fun ;)""" & _
               vbCrLf & vbCrLf & "When we declare myString we are actually creating a " & _
               "data type of BSTR.  A BSTR is a pointer to a null-terminated Unicode " & _
               "character array that is preceded by a 4-byte length field." & vbCrLf & vbCrLf & _
               "So, this means that myString is really a Long (4 bytes) that contains the " & _
               "address of the first Unicode character in the string ""VB strings are fun ;)"".  " & _
               "This Unicode character array is terminated by a NULL character, Chr(0), and " & _
               "the length of the string is stored in the 4 bytes that preceed the character " & _
               "array in memory.  The length value is the number of bytes in the string (the " & _
               "value returned by LenB) and not the number of characters (the value returned by " & _
               "Len).  The NULL terminator is not included in the length." & vbCrLf & vbCrLf & _
               "Enter some text in the field above and click 'Evaluate'."

End Sub

Private Sub Text2_GotFocus()
   
   With Text2
     .SelStart = 0
     .SelLength = Len(.Text)
   End With
  
End Sub
