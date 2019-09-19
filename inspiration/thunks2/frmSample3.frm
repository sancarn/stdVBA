VERSION 5.00
Begin VB.Form frmSample3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple Callbacks"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5820
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Run Example"
      Height          =   825
      Left            =   4560
      TabIndex        =   3
      Top             =   135
      Width           =   960
   End
   Begin VB.ListBox List2 
      Height          =   2370
      Left            =   3060
      TabIndex        =   1
      Top             =   1020
      Width           =   2460
   End
   Begin VB.ListBox List1 
      Height          =   2370
      Left            =   300
      TabIndex        =   0
      Top             =   1020
      Width           =   2460
   End
   Begin VB.Label Label1 
      Caption         =   $"frmSample3.frx":0000
      Height          =   990
      Left            =   315
      TabIndex        =   2
      Top             =   135
      Width           =   4155
   End
End
Attribute VB_Name = "frmSample3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function EnumChildWindows Lib "user32.dll" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function DispCallFunc Lib "OleAut32.dll" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByVal paTypes As Long, ByVal paValues As Long, ByRef retVAR As Variant) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Sub Command1_Click()

    Dim cThunks As clsThunks, oTasker As Object
    Dim lCount As Long
    
    '/// these variables used to set up the call to QSort
    Dim hMod As Long, fPtr As Long
    Dim vData(0 To 199) As Double
    Dim pParams(0 To 3) As Variant, vRtn As Long
    Dim pVarTypes(0 To 3) As Integer, pPtrs(0 To 3) As Long
  
    Randomize Timer
    For lCount = 0 To 199: vData(lCount) = Rnd * 100000: Next
    
    hMod = LoadLibrary("msvcrt20.dll")
    Debug.Assert hMod <> 0
    fPtr = GetProcAddress(hMod, "qsort")
    Debug.Assert fPtr <> 0
    
    '/// create a CDecl callback for QSort
    Set cThunks = New clsThunks
    Set oTasker = cThunks.CreateTasker_Callback(Me, 2, 2, , True)
    
    pParams(0) = VarPtr(vData(0)): pParams(1) = 200&
    pParams(2) = 8&: pParams(3) = oTasker.AddressOf
    
    pVarTypes(0) = vbLong: pVarTypes(1) = vbLong
    pVarTypes(2) = vbLong: pVarTypes(3) = vbLong
    pPtrs(0) = VarPtr(pParams(0)): pPtrs(1) = VarPtr(pParams(1))
    pPtrs(2) = VarPtr(pParams(2)): pPtrs(3) = VarPtr(pParams(3))
    
    DispCallFunc 0&, fPtr, 1, vbLong, 4, VarPtr(pVarTypes(0)), VarPtr(pPtrs(0)), vRtn
    FreeLibrary hMod
    List2.Clear
    For lCount = 0 To 199: List2.AddItem vData(lCount): Next
    
    '//// create a stdCall callback for EnumChildWindows
    Set oTasker = cThunks.CreateTasker_Callback(Me, 1, 2)
    lCount = 0: List1.Clear
    EnumChildWindows 0, oTasker.AddressOf, VarPtr(lCount)
    MsgBox "There were " & lCount & " top level desktop windows enumerated", vbInformation + vbOKOnly, "Done"
    
End Sub

' //////////////// DO NOT ADD ANY NEW CODE FROM HERE TO THE END OF THE CODE PAGE \\\\\\\\\\\\\\\\\\\\
Private Function qsort_compare_up(ByRef arg1 As Double, ByRef arg2 As Double) As Long

  Select Case arg2 - arg1
  Case Is < 0: qsort_compare_up = 1
  Case Is > 0: qsort_compare_up = -1
  End Select
  
End Function        ' ordinal #2

Private Function myEnumCallback(ByVal hWnd As Long, ByRef uParam As Long) As Long

    If hWnd = Me.hWnd Then
        List1.AddItem "Me: " & hWnd
    Else
        List1.AddItem "hWnd: " & hWnd
    End If
    uParam = uParam + 1
    myEnumCallback = 1
    
End Function        ' ordinal #1
' //////////////////////////////// DO NOT ADD ANY NEW CODE BELOW THIS BANNER \\\\\\\\\\\\\\\\\\\\\\\\\\\\\

