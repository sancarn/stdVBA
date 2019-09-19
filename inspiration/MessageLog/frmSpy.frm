VERSION 5.00
Begin VB.Form frmSpy 
   Caption         =   "Message log by the trick"
   ClientHeight    =   7635
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8400
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   509
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   450
      Left            =   4125
      TabIndex        =   1
      Top             =   7020
      Width           =   1335
   End
   Begin VB.CommandButton cmdPick 
      Caption         =   "Drag on window"
      Height          =   450
      Left            =   2790
      TabIndex        =   0
      Top             =   7020
      Width           =   1335
   End
End
Attribute VB_Name = "frmSpy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Форма frmSpy.frm
' © Кривоус Анатолий Анатольевич (The trick), 2014
' Работает вместе с modEnumWindow.mod, modInjection.mod, modListView.mod

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long) As Long
Private Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function FrameRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private Const COLOR_WINDOWFRAME As Long = 6
Private Const HS_DIAGCROSS As Long = 5
Private Const SM_CXBORDER As Long = 5
Private Const SM_CYBORDER As Long = 6
Private Const SM_CXFRAME = 32
Private Const SM_CYFRAME = 33
Private Const R2_NOT As Long = 6
Private Const PS_INSIDEFRAME As Long = 6
Private Const NULL_BRUSH As Long = 5

Private Const ControlSpacing As Long = 5                                                        ' Расстояние между контролами

Dim hWndprev As Long                                                                            ' Хендл "захваченного" окна
Dim Pn As Long, Br As Long                                                                      ' Перо и кисть

' Установка позиции и размеров контролов в зависимости от размеров формы
Private Sub SetMetrics()
    Dim CsPx As Long, l As Long, t As Long, w As Long, h As Long

    w = Me.ScaleWidth - ControlSpacing * 2
    h = Me.ScaleHeight - ControlSpacing * 3 - cmdPick.Height
    
    cmdPick.Move (Me.ScaleWidth - cmdPick.Width - cmdStop.Width - ControlSpacing) \ 2, _
                 ControlSpacing * 2 + h
    cmdStop.Move cmdPick.Left + cmdPick.Width, _
                      cmdPick.Top
                      
    MoveWindow modListView.hListView, ControlSpacing, _
                         ControlSpacing, _
                         w, h, True
End Sub

' Нажали на кнопку - начинаем поиск
Private Sub cmdPick_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    hWndprev = 0
    cmdPick.Caption = "Drop on window"
End Sub
' Перемещение кнопки над окнами
Private Sub cmdPick_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If modInjection.hThread = 0 Then MarkWindow
End Sub
' Отпустили кнопку
Private Sub cmdPick_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    MarkWindow True
    cmdPick.Caption = "Drag on window"
    If hWndprev Then
        modInjection.Clear
        cmdPick.Enabled = Not modInjection.Hook(hWndprev)
    End If
End Sub
Private Sub cmdStop_Click()
    modInjection.Clear               ' Удаляем инъекцию
    hWndprev = 0
    cmdPick.Enabled = True
End Sub
Private Sub Form_Load()
    modListView.InitListView
    Pn = CreatePen(PS_INSIDEFRAME, GetSystemMetrics(SM_CXBORDER) * 3, vbBlack)
    Br = CreateHatchBrush(HS_DIAGCROSS, GetSysColor(COLOR_WINDOWFRAME))
End Sub

' Помечаем окно рамкой
Private Sub MarkWindow(Optional Cancel As Boolean)
    Dim hWndFp As Long, Buf As String, l As Long

    If Cancel And hWndprev Then DrawFrame hWndprev: Exit Sub    ' Это когда отпускаем кнопку, чтобы отменить изменения
    
    hWndFp = modEnumWindow.GetWindowFromCursorPos               ' Получаем окно под курсором
    
    If hWndFp <> hWndprev Then
        If hWndFp Then
            Buf = String(256, 0)
            l = GetClassName(hWndFp, Buf, 255)
            If l Then Buf = Left(Buf, l)
            Me.Caption = Hex(hWndFp) & " Class = '" & Buf & "' "
            DrawFrame hWndFp
        Else
            Me.Caption = "Message log by the trick"
        End If
        If hWndprev Then DrawFrame hWndprev
    End If
    
    hWndprev = hWndFp
End Sub
' Рисует рамку у окна
Private Sub DrawFrame(lhWnd As Long)
    Dim hDCWnd As Long, RC As RECT, oPn As Long, oBr As Long, _
        hRgn As Long, SzX As Long, SzY As Long

    hDCWnd = GetWindowDC(lhWnd)
    If hDCWnd = 0 Then Exit Sub
    SetROP2 hDCWnd, R2_NOT
    oPn = SelectObject(hDCWnd, Pn)
    oBr = SelectObject(hDCWnd, GetStockObject(NULL_BRUSH))
    
    hRgn = CreateRectRgn(0, 0, 0, 0)
    SzX = GetSystemMetrics(SM_CXBORDER) * 3
    SzY = GetSystemMetrics(SM_CYBORDER) * 3
    
    If GetWindowRgn(lhWnd, hRgn) Then
        FrameRgn hDCWnd, hRgn, oBr, SzX, SzY
    Else
        GetWindowRect lhWnd, RC
        If IsZoomed(lhWnd) Then
            RC.Left = GetSystemMetrics(SM_CXFRAME)
            RC.Top = GetSystemMetrics(SM_CYFRAME)
            RC.Right = RC.Right + RC.Left
            RC.Bottom = RC.Bottom + RC.Top
        End If
        Rectangle hDCWnd, 0, 0, RC.Right - RC.Left, RC.Bottom - RC.Top
    End If
    
    SelectObject hDCWnd, oBr
    SelectObject hDCWnd, oPn
    ReleaseDC lhWnd, hDCWnd
End Sub
Private Sub Form_Resize()
    If Me.ScaleHeight > (cmdPick.Height + ControlSpacing * 3) Then SetMetrics
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call cmdStop_Click
    DeleteObject Pn
    DeleteObject Br
    DestroyListView
End Sub
