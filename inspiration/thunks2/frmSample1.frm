VERSION 5.00
Begin VB.Form frmSample1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subclass plus Hook Sample"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5730
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
   ScaleHeight     =   2550
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboColors 
      Height          =   330
      Index           =   0
      Left            =   1590
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1575
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      Height          =   270
      Left            =   1290
      Top             =   1950
      Width           =   3180
   End
   Begin VB.Label Label2 
      Caption         =   "Gradients"
      Height          =   270
      Index           =   1
      Left            =   2940
      TabIndex        =   2
      Top             =   1080
      Width           =   2475
   End
   Begin VB.Label Label2 
      Caption         =   "Basic RGB colors"
      Height          =   270
      Index           =   0
      Left            =   225
      TabIndex        =   1
      Top             =   1080
      Width           =   2475
   End
   Begin VB.Label Label1 
      Caption         =   $"frmSample1.frx":0000
      Height          =   885
      Left            =   180
      TabIndex        =   0
      Top             =   195
      Width           =   5280
   End
End
Attribute VB_Name = "frmSample1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cThunks As clsThunks
Private oTasker As Object

'///////////// following used for custom drawing combobox, none of it for the thunks
' note about this sample project...
'   we have a hidden indexed combobox on the form
'       that is used to dynamically load additional combos from it
'       this allows us to use the normal events that VB offers
'   it doesn't matter what style that combo has assigned to it, because the
'       hook function changes the style as we desire when the control is created
'   another option is to use something like the following, per combobox
'       Dim WithEvents cboXYZ As ComboBox
'       -- then to create it: Set cboXYZ = Me.Controls.Add("VB.ComboBox, "cboXYZ")

' This owner-drawn sample was borrowed from vbForums in a posting by The Trick
' http://www.vbforums.com/showthread.php?789527-VB6-Combobox-for-color-selection

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type DRAWITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemAction As Long
    itemState As Long
    hwndItem As Long
    hDC As Long
    rcItem As RECT
    ItemData As Long
End Type
Private Type MEASUREITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemWidth As Long
    itemHeight As Long
    ItemData As Long
End Type
 
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function SetDCBrushColor Lib "gdi32" (ByVal hDC As Long, ByVal colorref As Long) As Long
Private Declare Function SetDCPenColor Lib "gdi32" (ByVal hDC As Long, ByVal colorref As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, lpStr As Any, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
 
Private Const TRANSPARENT As Long = 1
Private Const COLOR_WINDOW As Long = 5
Private Const COLOR_WINDOWTEXT As Long = 8
Private Const COLOR_HIGHLIGHT As Long = 13
Private Const COLOR_HIGHLIGHTTEXT As Long = 14
Private Const ODS_SELECTED As Long = &H1
Private Const DC_PEN As Long = 19
Private Const DC_BRUSH As Long = 18
Private Const WH_CBT As Long = 5
Private Const HCBT_CREATEWND As Long = 3
Private Const ODT_COMBOBOX As Long = 3
Private Const CBS_OWNERDRAWFIXED As Long = &H10&
Private Const CBS_DROPDOWNLIST As Long = &H3&
Private Const CBS_HASSTRINGS As Long = &H200&
Private Const WM_MEASUREITEM As Long = &H2C
Private Const WM_DRAWITEM = &H2B
Private Const GWL_STYLE As Long = &HFFFFFFF0
Private Const DT_SINGLELINE As Long = &H20, DT_VCENTER As Long = &H4
Private Const CB_GETLBTEXT As Long = &H148
Private Const CB_GETLBTEXTLEN As Long = &H149
 
Private Sub Form_Load()

    Dim oHook As Object, oCombo As ComboBox
    Set cThunks = New clsThunks
    
    ' create a subclass tasker @ ordinal #1, using extended mode for filtering
    ' we are subclassing the form because it is the container for the
    '   comboboxes & the drawing messages are sent to it. If we had
    '   these combos in a picbox or frame, we'd be subclassing that instead
    Set oTasker = cThunks.CreateTasker_Subclass(Me, 1, Me.hWnd, True)
    oTasker.AddFilterMessage 0, Array(WM_DRAWITEM, WM_MEASUREITEM)
    
    ' create hook tasker @ ordinal #2, will be released on exit of Form_Load
    Set oHook = cThunks.CreateTasker_Hook(Me, 2, WH_CBT, , True)
    oHook.AddFilterMessage 0, 1, HCBT_CREATEWND ' filtering available w/extended version
    
    ' create the first combobox, position it, fill it, and show it
    Load cboColors(1)   ' << hook catches it at this point
    oHook.Pause         ' just showing that we can turn it off & back on as needed
    Set oCombo = cboColors(1)
    With oCombo
        .Move Label2(0).Left, Label2(0).Top + Label2(0).Height
        .AddItem "Red": .ItemData(.NewIndex) = vbRed
        .AddItem "Blue": .ItemData(.NewIndex) = vbBlue
        .AddItem "Green": .ItemData(.NewIndex) = vbGreen
        .AddItem "Yellow": .ItemData(.NewIndex) = vbYellow
        .AddItem "Cyan": .ItemData(.NewIndex) = vbCyan
        .AddItem "Magenta": .ItemData(.NewIndex) = vbMagenta
        .AddItem "Black": .ItemData(.NewIndex) = vbBlack
        .AddItem "White": .ItemData(.NewIndex) = vbWhite
        .Visible = True
    End With
    
    ' create the second combobox, position it, fill it, and show it
    Dim r As Long, g As Long, b As Long, c As Long
    
    oHook.resume        ' turn hook back on
    Load cboColors(2)   ' << hook catches it at this point
    Set oHook = Nothing ' not really needed, oHook is released when routine exits
    Set oCombo = cboColors(2)
    With oCombo
        .Move Label2(1).Left, Label2(1).Top + Label2(1).Height
        Do
            c = RGB(r, g, b)
            .AddItem "0x" & Right$("0000000" & Hex$(c), 8)
            .ItemData(.NewIndex) = c
            r = r + &H40
            If r > 255 Then r = 0: g = g + &H40
            If g > 255 Then g = 0: b = b + &H40
            If b > 255 Then Exit Do
        Loop
        .Visible = True
    End With
    Set oCombo = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oTasker = Nothing
    Set cThunks = Nothing
End Sub

Private Sub cboColors_Click(Index As Integer)
    Shape1.BackColor = cboColors(Index).ItemData(cboColors(Index).ListIndex)
End Sub

' //////////////// DO NOT ADD ANY NEW CODE FROM HERE TO THE END OF THE CODE PAGE \\\\\\\\\\\\\\\\\\\\
Private Function myHookProc(ByVal nCode As Long, ByVal wParam As Long, _
                        ByVal lParam As Long, ByVal pTasker As Object, _
                        ByRef EatMessage As Boolean) As Long
    
    If nCode = HCBT_CREATEWND Then
        Dim sClass As String, l As Long, lStyle As Long
        sClass = Space$(256)
        l = GetClassName(wParam, sClass, 255)
        If l Then
            sClass = Left$(sClass, l)   ' test for design-time & run-time classes
            If StrComp(sClass, "ThunderComboBox", vbTextCompare) = 0 Or _
               StrComp(sClass, "ThunderRT6ComboBox", vbTextCompare) = 0 Then
                lStyle = GetWindowLong(wParam, GWL_STYLE)
                SetWindowLong wParam, GWL_STYLE, lStyle Or CBS_OWNERDRAWFIXED Or CBS_DROPDOWNLIST Or CBS_HASSTRINGS
            End If
        End If
    End If

End Function    ' ordinal #2

Private Function myWindowProc(ByVal hWnd As Long, ByVal uMsg As Long, _
                        ByVal wParam As Long, ByVal lParam As Long, _
                        ByVal hWndTag As Long, ByVal pTasker As Object, _
                        ByRef EatMessage As Boolean) As Long

    ' note: if we were owner-drawing many controls, we'd probably want to
    ' move this stuff to its own routine and test the passed hWnd parameter
    ' to verify it is an hWnd we are owner-drawing vs VB or some other 'owner'.
    ' For this simple example, we have no other owner-drawn comboboxes, so that
    ' sanity check isn't needed, i.e.,
    '   If Not (hWnd = cboColors(1).hWnd Or hWnd = cboColors(2).hWnd) Then Exit Function

    If uMsg = WM_DRAWITEM Then
        Dim drw As DRAWITEMSTRUCT
        Dim obr As Long, opn As Long, l As Long, s As String
        CopyMemory drw, ByVal lParam, Len(drw)
        If drw.CtlType = ODT_COMBOBOX Then
        
            ' handled, don't forward down the chain
            EatMessage = True: myWindowProc = EatMessage
            
            obr = SelectObject(drw.hDC, GetStockObject(DC_BRUSH))
            opn = SelectObject(drw.hDC, GetStockObject(DC_PEN))
            If (drw.itemState And ODS_SELECTED) Then
                SetDCBrushColor drw.hDC, GetSysColor(COLOR_HIGHLIGHT)
                SetDCPenColor drw.hDC, GetSysColor(COLOR_HIGHLIGHT)
                Rectangle drw.hDC, drw.rcItem.Left, drw.rcItem.Top, drw.rcItem.Right, drw.rcItem.Bottom
                SetDCPenColor drw.hDC, GetSysColor(COLOR_HIGHLIGHTTEXT)
                SetTextColor drw.hDC, GetSysColor(COLOR_HIGHLIGHTTEXT)
            Else
                SetDCBrushColor drw.hDC, GetSysColor(COLOR_WINDOW)
                SetDCPenColor drw.hDC, GetSysColor(COLOR_WINDOW)
                Rectangle drw.hDC, drw.rcItem.Left, drw.rcItem.Top, drw.rcItem.Right, drw.rcItem.Bottom
                SetDCPenColor drw.hDC, GetSysColor(COLOR_WINDOWTEXT)
                SetTextColor drw.hDC, GetSysColor(COLOR_WINDOWTEXT)
            End If
            SetBkMode drw.hDC, TRANSPARENT
            If drw.itemID >= 0 Then
                SetDCBrushColor drw.hDC, drw.ItemData
                Rectangle drw.hDC, drw.rcItem.Left + 3, drw.rcItem.Top + 3, drw.rcItem.Left + 70, drw.rcItem.Bottom - 3
                l = SendMessage(drw.hwndItem, CB_GETLBTEXTLEN, drw.itemID, ByVal 0)
                If l Then
                    s = Space$(l + 1)
                    l = SendMessage(drw.hwndItem, CB_GETLBTEXT, drw.itemID, ByVal s)
                    s = Left$(s, l)
                    drw.rcItem.Left = drw.rcItem.Left + 78
                End If
            Else
                drw.rcItem.Left = drw.rcItem.Left + 2
                s = "None"
            End If
            DrawText drw.hDC, ByVal s, Len(s), drw.rcItem, DT_VCENTER Or DT_SINGLELINE
            SelectObject drw.hDC, obr
            SelectObject drw.hDC, opn
        End If
        
    ElseIf uMsg = WM_MEASUREITEM Then
        Dim meas As MEASUREITEMSTRUCT, RC As RECT
        CopyMemory meas, ByVal lParam, Len(meas)
        If meas.CtlType = ODT_COMBOBOX Then
        
            ' handled, don't forward down the chain
            EatMessage = True: myWindowProc = EatMessage
            
            GetClientRect hWnd, RC
            meas.itemWidth = RC.Right - RC.Left
            CopyMemory ByVal lParam, meas, Len(meas)
        End If
    End If

End Function    ' ordinal #1
' //////////////////////////////// DO NOT ADD ANY NEW CODE BELOW THIS BANNER \\\\\\\\\\\\\\\\\\\\\\\\\\\\\

