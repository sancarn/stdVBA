VERSION 5.00
Begin VB.Form frmApplets 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Пример с апплетами"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOut 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4380
      Left            =   195
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   675
      Width           =   6615
   End
   Begin VB.Label Label1 
      Caption         =   "Этот пример поочерёдно загружает CPL-апплеты в память и вызывают у них функцию CplApplet используя указатель на функцию."
      Height          =   915
      Left            =   135
      TabIndex        =   0
      Top             =   165
      Width           =   6630
   End
End
Attribute VB_Name = "frmApplets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal LibName As String) As Long
Private Declare Sub FreeLibrary Lib "kernel32" (ByVal hModule As Long)
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal addr As String) As Long


Private Declare Function LoadStringW Lib "user32" (ByVal hInstance As Long, _
                                                   ByVal id As Long, _
                                                   ByVal psBuffer As Long, _
                                                   ByVal nBufferMax As Long) As Long
                                                     
  
Private Enum CPL_MSG
    CPL_EMPTY
    CPL_INIT
    CPL_GETCOUNT
    CPL_INQUIRE
    CPL_SELECT
    CPL_DBLCLK
    CPL_STOP
    CPL_EXIT
    CPL_NEWINQUIRE
End Enum

Private Type CPLINFO
    idIcon                   As Long
    idName                   As Long
    idInfo                   As Long
    lpData                   As Long
End Type

' Указатель:
Private Declare Function ptrMyFunc Lib "*" (ByVal hwnd As Long, _
                                            ByVal msg As CPL_MSG, _
                                            ByVal param1 As Long, _
                                            ByVal param2 As Long) As Long

' Тот же указатель с другим прототипом:
Private Declare Function ptrMyFuncStruct Lib "*" Alias "ptrMyFunc" _
                                           (ByVal hwnd As Long, _
                                            ByVal msg As CPL_MSG, _
                                            ByVal index As Long, _
                                            ByRef info As CPLINFO) As Long



Private Sub Form_Load()
    ChDrive Environ$("windir")
    ChDir Environ$("windir")
    ChDir "system32"
    
    Dim fn As String
    Dim hModule As Long
    Dim i As Long
    Dim sBuf As String
    Dim nBufLen  As Long
    

    
    fn = Dir("*.cpl", vbSystem Or vbHidden Or vbNormal Or vbReadOnly)
    
    Do
        AppendNL "Апплет '" + fn + "':"
        hModule = LoadLibraryA(fn)
        If hModule <> 0 Then
            FuncPointer("ptrMyFunc") = GetProcAddress(hModule, "CPlApplet")
            If (FuncPointer("ptrMyFunc") <> 0) Then
                If ptrMyFunc(Me.hwnd, CPL_INIT, 0, 0) <> 0 Then
                    
                    '
                    ' Проходимся по элементам апплета.
                    '
                    
                    Dim inf As CPLINFO
                    
                    For i = 0 To ptrMyFunc(Me.hwnd, CPL_GETCOUNT, 0, 0) - 1

                        Call ptrMyFuncStruct(Me.hwnd, _
                                             CPL_INQUIRE, _
                                             i, _
                                             inf)
                                            
                                                                    
                        sBuf = String(256, 0)
                        nBufLen = LoadStringW(hModule, inf.idName, StrPtr(sBuf), 256)
                        sBuf = Left$(sBuf, InStr(1, sBuf, Chr$(0)) - 1)
                        
                        AppendNL "  Диалог: " + sBuf
                    Next i
                    
                    Call ptrMyFunc(Me.hwnd, CPL_EXIT, 0, 0)
                Else
                    AppendNL "  Не удалось выполнить инициализацию."
                End If
            Else
                AppendNL "  Не удалось найти точку входа CPlApplet."
            End If
            FreeLibrary hModule
        Else
            AppendNL "  Не удалось выполнить LoadLibrary."
        End If
        fn = Dir()
    Loop Until Len(fn) = 0
    

End Sub

Private Sub AppendT(ByRef t As String)
    txtOut.Text = txtOut.Text + t
End Sub

Private Sub AppendNL(ByRef t As String)
    txtOut.Text = txtOut.Text + t + vbNewLine
End Sub
