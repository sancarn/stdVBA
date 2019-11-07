VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Save/Load UDT"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUDT 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   1200
      Width           =   6015
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "UDT info:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   750
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Save/Load UDT arrays by DigiRev
'Example of how to save and load UDT arrays to disk.

Option Explicit

'An example UDT.
Private Type MY_UDT
    strName As String
    intAge As Integer
    strPassword As String
    bolDummy As Boolean
End Type

'Array of the UDT.
'The one we'll be saving and then loading.
Private udtTest() As MY_UDT

Private Sub cmdLoad_Click()
    LoadUDT App.Path & "\udt.txt"
    
    'Now display the UDT in the textbox.
    '-----------------------------------
    Dim intLoop As Integer, strName As String
    
    With txtUDT
        For intLoop = 0 To UBound(udtTest())
            .SelStart = Len(.Text)
            strName = "udtTest(" & intLoop & ")"
            .SelText = strName & vbCrLf & String$(Len(strName), "-") & vbCrLf
            .SelText = Space$(2) & ".bolDummy:    " & udtTest(intLoop).bolDummy & vbCrLf
            .SelText = Space$(2) & ".intAge:      " & udtTest(intLoop).intAge & vbCrLf
            .SelText = Space$(2) & ".strName:     " & udtTest(intLoop).strName & vbCrLf
            .SelText = Space$(2) & ".strPassword: " & udtTest(intLoop).strPassword & vbCrLf & vbCrLf
        Next intLoop
    End With
End Sub

'Save the UDT array.
Private Sub cmdSave_Click()
    Dim strPath As String 'Path to save the UDT array.
    
    'It's a binary file, but i used .txt so you can easily open it in notepad if
    'you're curious what a UDT array looks like when written to disk.
    strPath = App.Path & "\udt.txt"
    
    SaveUDT strPath
    
    MsgBox "Saved", vbInformation
    cmdLoad.Enabled = True
End Sub

Private Sub Form_Load()
    LoadSomeUDTs
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase udtTest()
End Sub

'Load the file back into the UDT array (udtTest()).
Private Sub LoadUDT(ByRef FilePath As String)
    'First thing we'll do is look at the 4 byte header we wrote to the file.
    'This will tell us how many items are in the array.
    'Then we can ReDim udtTest() to the correct dimensions.
    'Then we can use the Get # statement to dump the file right back into the UDT array.
    
    Dim intFF As Integer, strHeader As String
    
    intFF = FreeFile
    
    'Buffer strHeader to 4 bytes.
    strHeader = Space$(4)
    
    'Open file in binary read mode.
    Open FilePath For Binary Access Read As #intFF
        'Get 4 byte header.
        Get #intFF, 1, strHeader
        
        'Re-dimension the array.
        ReDim udtTest(CLng(strHeader)) As MY_UDT
        
        'Dump the file back into the UDT.
        Get #intFF, 5, udtTest()
    Close #intFF
End Sub

'Save UDT array to FilePath.
Private Sub SaveUDT(ByRef FilePath As String)
    'You can write a UDT array directly to a file using the Put # statement.
    'However, when loading them back from the file, we need to know how
    'many items there were in the array, so we can re-dimension it appropriately.
    
    'So we will write a short 4 byte header to the beginning of the file.
    'This will be a number telling us how many items were in the array.
    'Then we can ReDim() the array before loading it back from the file.
    
    'The header will always be 4 bytes, so for 3 item arrays, the header would be "0003"
    'This makes the max number of array items "9999" for this example.
    'You can easily modify the code to give you more.
    
    Dim intFF As Integer 'File handle to use.
    
    intFF = FreeFile 'Get available file handle.
    
    Open FilePath For Binary Access Write As #intFF
        Put #intFF, 1, BuildHeader(4) 'Write header.
        Put #intFF, LOF(intFF) + 1, udtTest() 'Write UDT array.
    Close #intFF
    
End Sub

'Pads a string with 0's until its 4 bytes long.
Private Function BuildHeader(ByVal Length As Long) As String
    BuildHeader = String$(Abs(4 - Len(CStr(UBound(udtTest())))), "0") & CStr(UBound(udtTest()))
End Function

'Just loads some stuff into the UDT array for testing.
Private Sub LoadSomeUDTs()
    ReDim udtTest(32) As MY_UDT '33 array items.
    
    Dim intLoop As Integer
    
    For intLoop = 0 To 32
        With udtTest(intLoop)
            .bolDummy = Int(Rnd * 1) - 1
            .intAge = Int(Rnd * 80) + 1
            .strName = "User " & intLoop
            .strPassword = String$(Int(Rnd * 32) + 1, Chr$(Int(Rnd * 255) + 1))
        End With
    Next intLoop
            
End Sub

