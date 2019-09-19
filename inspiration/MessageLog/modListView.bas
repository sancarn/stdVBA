Attribute VB_Name = "modListView"
Option Explicit

' Дополнительный модуль для работы с SysListView32
' © Кривоус Анатолий Анатольевич (The trick), 2014

' Стандартный ListViewWndClass перерисовывает всю клиентскую область при добавлении
' новой записи, либо при прокрутке из-за этого происходит неприятное мерцание, иногда
' даже полностью "белеет" фон. Для предотвращения такого поведения я решил использовать
' SysListView32, который ведет себя правильно при добавлении записей и прокрутке
' Также я использую в качестве идентификации сообщений коллекцию, с ключом - номером сообщения
' поэтому не поддерживаются одинаковые номера сообщений, т.к. я делал этот пример только
' ради демонстрации внедрения

Private Type LVITEM
    mask       As Long
    iItem      As Long
    iSubItem   As Long
    State      As Long
    stateMask  As Long
    pszText    As String
    cchTextMax As Long
    iImage     As Long
    lParam     As Long
    iIndent    As Long
End Type
Private Type LVCOLUMN
    mask As Long
    fmt As Long
    CX As Long
    pszText As String
    cchTextMax As Long
    iSubItem As Long
    iImage As Long
    iOrder As Long
End Type
Private Type tagInitCommonControlsEx
    dwSize As Long
    dwICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32" (ByRef TLPINITCOMMONCONTROLSEX As tagInitCommonControlsEx) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Const ICC_WIN95_CLASSES = &HFF
Private Const WS_CHILD = &H40000000
Private Const WS_TABSTOP = &H10000
Private Const LVS_REPORT = &H1&
Private Const LVS_SINGLESEL = &H4&
Private Const WS_EX_CLIENTEDGE = &H200&
Private Const LVS_EX_FULLROWSELECT = &H20&
Private Const LVS_EX_GRIDLINES = &H1&
Private Const SW_SHOW = 5

Private Const LVM_FIRST = &H1000
Private Const LVM_INSERTCOLUMN = (LVM_FIRST + 27)
Private Const LVM_INSERTITEM = (LVM_FIRST + 7)
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 54)
Private Const LVM_GETITEMCOUNT = (LVM_FIRST + 4)
Private Const LVM_ENSUREVISIBLE = (LVM_FIRST + 19)
Private Const LVM_SETITEMTEXTA = (LVM_FIRST + 46)
Private Const LVCF_WIDTH = &H2
Private Const LVCF_TEXT = &H4
Private Const LVIF_TEXT = &H1

Public hListView As Long                                                                        ' Хендл
Public Dic As Collection                                                                        ' Список сообщений

' Инициализация ListView
Public Sub InitListView()
    Dim ExStyle As Long
    Dim LVStyle As Long
    Dim Col As LVCOLUMN
    Dim CC As tagInitCommonControlsEx
    
    CC.dwSize = Len(CC)
    CC.dwICC = ICC_WIN95_CLASSES
    
    If InitCommonControlsEx(CC) = 0 Then MsgBox "Error InitCommonControlsEx": End
    
    ExStyle = WS_EX_CLIENTEDGE                                                                  ' Рамка у ListView
    LVStyle = WS_CHILD Or WS_TABSTOP Or LVS_REPORT Or LVS_SINGLESEL                             ' Стиль Report и единственный выбор
    
    hListView = CreateWindowEx(ExStyle, "SysListView32", vbNullString, LVStyle, 5, 5, 100, 100, frmSpy.hwnd, 0, App.hInstance, ByVal 0)
    
    If hListView = 0 Then MsgBox "Error creating ListView " & Err.LastDllError, vbCritical: End ' Если не удалось создать - нет смысла
                                                                                                ' продолжать работать
    
    SendMessage hListView, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, _
                ByVal LVS_EX_FULLROWSELECT Or LVS_EX_GRIDLINES                                  ' Установка расширеных стилей:
                                                                                                ' выбор всей строки и сетка
    ' Вставляем колонки в ListView
    Col.mask = LVCF_TEXT Or LVCF_WIDTH
    Col.pszText = "№": Col.cchTextMax = Len(Col.pszText): Col.CX = 64
    SendMessage hListView, LVM_INSERTCOLUMN, 0, Col
    
    Col.pszText = "Message": Col.cchTextMax = Len(Col.pszText): Col.CX = 200
    SendMessage hListView, LVM_INSERTCOLUMN, 1, Col
    
    Col.pszText = "wParam": Col.cchTextMax = Len(Col.pszText): Col.CX = 100
    SendMessage hListView, LVM_INSERTCOLUMN, 2, Col
    
    Col.pszText = "lParam": Col.cchTextMax = Len(Col.pszText): Col.CX = 100
    SendMessage hListView, LVM_INSERTCOLUMN, 3, Col
    
    Call ShowWindow(hListView, SW_SHOW)                                                         ' Показываем окно
End Sub
' Уничтожение ListView
Public Sub DestroyListView()
    DestroyWindow hListView                                                                     ' Уничтожаем окно
    hListView = 0
End Sub
' Инициализация словаря
Public Sub DicInit()
    Dim fNum As Integer, s As String, key As String
    
    On Error GoTo Errorlabel
    
    fNum = FreeFile
    
    Open App.Path & "\WMList.txt" For Input As fNum
    
    Set Dic = New Collection
    
    Do Until EOF(fNum)
        Line Input #fNum, s
        key = "_" & Left$(s, 4)
        Dic.Add Mid$(s, 5), key
    Loop
    
    Close fNum
    
    Exit Sub
Errorlabel:
    MsgBox "Windows messages list loading error", vbExclamation
    Err.Clear
End Sub
' Добавить строку (без мерцания)
Public Function ItemAdd(ByVal Message As String, ByVal wParam As String, ByVal lParam As String) As Boolean
    Dim LV As LVITEM, i As Long
    
    i = SendMessage(hListView, LVM_GETITEMCOUNT, 0, ByVal 0&)
    
    With LV
      .pszText = i
      .iItem = i
      .cchTextMax = Len(.pszText)
      .mask = LVIF_TEXT
    End With
    
    SendMessage hListView, LVM_INSERTITEM, 0, LV
    LV.pszText = Message: LV.iSubItem = 1
    SendMessage hListView, LVM_SETITEMTEXTA, i, LV
    LV.pszText = wParam: LV.iSubItem = 2
    SendMessage hListView, LVM_SETITEMTEXTA, i, LV
    LV.pszText = lParam: LV.iSubItem = 3
    SendMessage hListView, LVM_SETITEMTEXTA, i, LV
    
    SendMessage hListView, LVM_ENSUREVISIBLE, i, ByVal True
End Function
' Возвращает имя сообщения по номеру
Public Function GetMessageName(ByVal Number As Long) As String
    On Error Resume Next
    Dim h As String
    
    h = "0000": Mid$(h, 5 - Len(Hex(Number))) = Hex(Number)
    GetMessageName = Dic.Item("_" & h)
    If Err.Number Then GetMessageName = h: Err.Clear
End Function
