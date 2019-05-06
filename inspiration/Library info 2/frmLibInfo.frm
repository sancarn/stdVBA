VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmLibInfo 
   Caption         =   "Library info 2"
   ClientHeight    =   8655
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   14790
   LinkTopic       =   "Form1"
   ScaleHeight     =   577
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   986
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar stbStatus 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   5
      Top             =   8340
      Width           =   14790
      _ExtentX        =   26088
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin MSComctlLib.ImageList iglLib 
      Left            =   6720
      Top             =   1290
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibInfo.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibInfo.frx":0322
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibInfo.frx":0644
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibInfo.frx":0966
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibInfo.frx":0CB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibInfo.frx":0FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibInfo.frx":12FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibInfo.frx":161E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibInfo.frx":1940
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibInfo.frx":1A52
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibInfo.frx":1B64
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibInfo.frx":1C76
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibInfo.frx":1FC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibInfo.frx":231A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibInfo.frx":242C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibInfo.frx":277E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibInfo.frx":2AD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibInfo.frx":2BE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibInfo.frx":2CF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibInfo.frx":2E06
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibInfo.frx":2F18
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwLib 
      Height          =   8400
      Left            =   90
      TabIndex        =   4
      Top             =   150
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   14817
      _Version        =   393217
      Indentation     =   354
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "iglLib"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lvwImplements 
      Height          =   1050
      Left            =   4560
      TabIndex        =   3
      Top             =   330
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   1852
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "iglLib"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   11359
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   9596
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMembers 
      Height          =   6810
      Left            =   4560
      TabIndex        =   0
      Top             =   1725
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   12012
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "iglLib"
      SmallIcons      =   "iglLib"
      ColHdrIcons     =   "iglLib"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label lblImplements 
      Caption         =   "Implements:"
      Height          =   285
      Left            =   4575
      TabIndex        =   2
      Top             =   75
      Width           =   975
   End
   Begin VB.Label lblMembers 
      Caption         =   "Members:"
      Height          =   285
      Left            =   4575
      TabIndex        =   1
      Top             =   1440
      Width           =   975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmLibInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Получение информации об импорте, экспорте и библиотеки типов
' © Кривоус Анатолий Анатольевич (The trick), 2014

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As Long
    lpstrCustomFilter As Long
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As Long
    nMaxFile As Long
    lpstrFileTitle As Long
    nMaxFileTitle As Long
    lpstrInitialDir As Long
    lpstrTitle As Long
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
End Type

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameW" (pOpenfilename As OPENFILENAME) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function VariantCopy Lib "oleaut32.dll" (pvargDest As Any, pvargSrc As Any) As Long

Private Lib         As ITypeLib
Private natInfo     As clsNativeInfo

Private Sub Form_Load()
    Set natInfo = New clsNativeInfo
End Sub
' Загрузить
Private Sub LoadLib(Path As String)
    Dim er      As ERROR_CODE
    ' Загрузить информацию о библиотеке
    er = LoadLibInfo(Path)
    ' Загрузить библиотеку типов
    If Not LoadTLB(Path) And er <> OK Then
        ' Не удалось загрузить ни TLB ни DLL (EXE)
        MsgBox natInfo.ErrorText(er), vbCritical, "Error"
        Exit Sub
    Else
        lvwImplements.ListItems.Clear
        lvwMembers.ListItems.Clear
    End If
End Sub
' Загрузить информацию о DLL
Private Function LoadLibInfo(Path As String) As ERROR_CODE
    Dim i           As Long
    Dim key         As String
    Dim itm         As String
    Dim icn         As Long
    Dim iInf        As ImportInfo
    Dim nInf        As clsNativeInfo
    ' Создаем парсер
    Set nInf = New clsNativeInfo
    ' Если неуспешно, то возвращаем код ошибки
    If Not nInf.Extract(Path) Then LoadLibInfo = nInf.ErrorCode: Exit Function
    Set natInfo = nInf
    ' Очищаем дерево
    tvwLib.Nodes.Clear
    ' Получаем имя файла
    itm = GetFileTitle(Path, True)
    ' Добавляем в корень дерева имя библиотеки
    tvwLib.Nodes.Add(, , "lib", natInfo.ModuleName & " (" & itm & ")", 8).Expanded = True
    ' Если есть экспорт, то добавляем узел
    If natInfo.ExportCount Then tvwLib.Nodes.Add("lib", tvwChild, "export", "Export", 1).Expanded = True
    ' Если есть импорт, то добавляем узел
    If natInfo.ImportCount Then
        tvwLib.Nodes.Add("lib", tvwChild, "import", "Import", 2).Expanded = True
        ' Проход по всем элементам импорта
        For i = 0 To natInfo.ImportCount - 1
            key = "imp_" & CStr(i)
            iInf = natInfo.Import(i)
            ' Добавляем в дерево имя библиотеки
            tvwLib.Nodes.Add "import", tvwChild, key, iInf.Name, 8
        Next
    End If
    ' Если есть отложенный импорт, то добавляем узел
    If natInfo.DelayImportCount Then
        tvwLib.Nodes.Add("lib", tvwChild, "dimport", "Delay import", 3).Expanded = True
        For i = 0 To natInfo.DelayImportCount - 1
            key = "dimp_" & CStr(i)
            iInf = natInfo.DelayImport(i)
            ' Добавляем в дерево имя библиотеки
            tvwLib.Nodes.Add "dimport", tvwChild, key, iInf.Name, 8
        Next
    End If
End Function
' Загрузить библиотеку типов
Private Function LoadTLB(Path As String) As Boolean
    Dim locLib      As ITypeLib
    Dim Name        As String
    Dim Desc        As String
    Dim i           As Long
    Dim inf         As ITypeInfo
    Dim key         As String
    Dim fl          As TYPEKIND
    
    ' Обработчик ошибок
    On Error Resume Next
    Set locLib = LoadTypeLibEx(Path, REGKIND_NONE)
    ' Если ошибка, значит нет библиотеки типов
    If Err.Number Then Err.Clear: Exit Function
    On Error GoTo 0
    ' Получаем описание
    locLib.GetDocumentation -1, Name, Desc, 0, vbNullString
    ' Добавляем узел
    If tvwLib.Nodes.Count = 0 Then
        tvwLib.Nodes.Add(, , "lib", GetFileTitle(Path), 8).Expanded = True
    End If
    tvwLib.Nodes.Add("lib", tvwChild, "com", Name & " (" & Desc & ")", 4).Expanded = True
    ' Проход по списку
    For i = 0 To locLib.GetTypeInfoCount - 1
        Set inf = locLib.GetTypeInfo(i)
        inf.GetDocumentation -1, Name, vbNullString, 0, vbNullString
        key = "co_" & CStr(i)
        fl = locLib.GetTypeInfoType(i)
        tvwLib.Nodes.Add "com", tvwChild, key, Name, fl + 9
    Next

    Set Lib = locLib
    
    LoadTLB = True
End Function
Private Sub LoadInterface(typ As ITypeInfo, Attr As TYPEATTR)
    Dim hType   As Long
    Dim i       As Long
    Dim impl    As ITypeInfo
    Dim Name    As String
    Dim Desc    As String
    Dim itm     As ListItem
    Dim ptr     As Long
    Dim finf    As FUNCDESC
    Dim icn     As Long
    Dim ret     As String
    ' Получение информации о реализуемых интерфейсах
    For i = 0 To Attr.cImplTypes - 1
        hType = typ.GetRefTypeOfImplType(i)
        Set impl = typ.GetRefTypeInfo(hType)
        impl.GetDocumentation -1, Name, Desc, 0, vbNullString
        Set itm = lvwImplements.ListItems.Add(, , Name, , 12)
        itm.SubItems(1) = Desc
    Next
    ' Получение информации о методах
    For i = 0 To Attr.cFuncs - 1
        ' Получаем описание метода
        ptr = typ.GetFuncDesc(i)
        CopyMemory finf, ByVal ptr, LenB(finf)
        ' Получаем имя и описание метода
        typ.GetDocumentation finf.memid, Name, Desc, 0, vbNullString
        ' Получаем тип метода и устанавливаем соответствующую иконку
        Select Case finf.invkind
        Case INVOKEKIND.INVOKE_FUNC: icn = 5
        Case INVOKEKIND.INVOKE_PROPERTYGET: icn = 18
        Case Else: icn = 19
        End Select
        ' Получаем тип возвращаемого значения
        ' Добавляем в список
        
        Set itm = lvwMembers.ListItems.Add(, , Name, , icn)
        If finf.wFuncFlags And FUNCFLAGS.FUNCFLAG_FRESTRICTED Then
            itm.ForeColor = &HA0A0A0
        ElseIf finf.wFuncFlags And FUNCFLAGS.FUNCFLAG_FHIDDEN Then
            itm.ForeColor = &HFF8080
        End If
        itm.SubItems(2) = Desc
        itm.SubItems(1) = GetTypeName(finf.elemdescFunc.tdesc, typ)
        typ.ReleaseFuncDesc ptr
    Next
End Sub
Private Sub LoadCoClass(typ As ITypeInfo, Attr As TYPEATTR)
    Dim ptr     As Long
    Dim icn     As Long
    Dim i       As Long
    Dim N       As Long
    Dim hType   As Long
    Dim isEvent As Boolean
    Dim Name    As String
    Dim Desc    As String
    Dim impl    As ITypeInfo
    Dim finf    As FUNCDESC
    Dim itm     As ListItem
    
    ' Получение информации о реализуемых интерфейсах
    For i = 0 To Attr.cImplTypes - 1
        hType = typ.GetRefTypeOfImplType(i)
        Set impl = typ.GetRefTypeInfo(hType)
        impl.GetDocumentation -1, Name, vbNullString, 0, vbNullString
        isEvent = typ.GetImplTypeFlags(i) = 3
        lvwImplements.ListItems.Add , , Name, , IIf(isEvent, 20, 12)
        ' Добавляем в список методы
        ' Получаем атрибуты
        ptr = impl.GetTypeAttr()
        CopyMemory Attr, ByVal ptr, Len(Attr)
        ' Получение информации о методах
        For N = 0 To Attr.cFuncs - 1
            ' Получаем описание метода
            ptr = impl.GetFuncDesc(N)
            CopyMemory finf, ByVal ptr, LenB(finf)
            ' Получаем имя и описание метода
            impl.GetDocumentation finf.memid, Name, Desc, 0, vbNullString
            ' Получаем тип метода и устанавливаем соответствующую иконку
            If isEvent Then
                icn = 20
            Else
                Select Case finf.invkind
                Case INVOKEKIND.INVOKE_FUNC: icn = 5
                Case INVOKEKIND.INVOKE_PROPERTYGET: icn = 18
                Case Else: icn = 19
                End Select
            End If
            ' Добавляем в список
            'Debug.Print finf.wFuncFlags, Name
            Set itm = lvwMembers.ListItems.Add(, , Name, , icn)
            If finf.wFuncFlags And FUNCFLAGS.FUNCFLAG_FRESTRICTED Then
                itm.ForeColor = &HA0A0A0
            ElseIf finf.wFuncFlags And FUNCFLAGS.FUNCFLAG_FHIDDEN Then
                itm.ForeColor = &HFF8080
            End If
            itm.SubItems(2) = Desc
            itm.SubItems(1) = GetTypeName(finf.elemdescFunc.tdesc, impl)
            impl.ReleaseFuncDesc ptr
        Next
        impl.ReleaseTypeAttr ptr
    Next
End Sub
' Загрузить тип
Private Sub LoadType(typ As ITypeInfo)
    Dim impl    As ITypeInfo
    Dim tAttr   As TYPEATTR
    Dim ptr     As Long
    Dim i       As Long
    Dim N       As Long
    Dim finf    As FUNCDESC
    Dim Name    As String
    Dim icn     As Long
    Dim hType   As Long
    Dim itm     As ListItem
    Dim Desc    As String
    Dim strGuid As String
    ' Получаем атрибуты
    ptr = typ.GetTypeAttr()
    CopyMemory tAttr, ByVal ptr, Len(tAttr)
    typ.ReleaseTypeAttr ptr
    
    strGuid = Space(38)
    StringFromGUID2 tAttr.iid, strGuid, 39

    ' Если это интерфейс
    If tAttr.TYPEKIND = TKIND_DISPATCH Or _
       tAttr.TYPEKIND = TKIND_INTERFACE Then
        ' Интерфейс
        LoadInterface typ, tAttr
        stbStatus.Panels.Add(, "guid", strGuid).AutoSize = sbrContents
    ElseIf tAttr.TYPEKIND = TKIND_COCLASS Then
        LoadCoClass typ, tAttr
        stbStatus.Panels.Add(, "guid", strGuid).AutoSize = sbrContents
    Else
        ' Если псевдоним
        If tAttr.TYPEKIND = TKIND_ALIAS Then
            ' Получаем тип переменной
            Do
                Select Case tAttr.tdescAlias.vt
                Case VARENUM.VT_USERDEFINED
                    ' Если псевдоним UDT
                    Set typ = typ.GetRefTypeInfo(tAttr.tdescAlias.pTypeDesc)
                    typ.GetDocumentation -1, Name, vbNullString, 0, vbNullString
                    ' Получаем атрибуты
                    ptr = typ.GetTypeAttr()
                    CopyMemory tAttr, ByVal ptr, Len(tAttr)
                    typ.ReleaseTypeAttr ptr
                    
                    lvwMembers.ListItems.Add(, , Name, , tAttr.TYPEKIND + 9).Tag = tAttr.tdescAlias.pTypeDesc
                    Exit Do
                Case VARENUM.VT_PTR
                    ' Если это ссылка, то получаем содержимое
                    CopyMemory tAttr.tdescAlias, ByVal tAttr.tdescAlias.pTypeDesc, Len(tAttr.tdescAlias)
                Case VARENUM.VT_CARRAY
                    ' Это массив
                    Dim arrDesc     As ARRAYDESC
                    CopyMemory arrDesc, ByVal tAttr.tdescAlias.pTypeDesc, Len(arrDesc)
                    Name = GetTypeName(arrDesc.tdescElem) & GetArray(tAttr.tdescAlias.pTypeDesc)
                    ' Добавляем в список
                    lvwMembers.ListItems.Add(, , Name, , 17).Tag = 0
                    Exit Do
                Case Else
                    ' Это стандартный тип
                    Name = GetTypeName(tAttr.tdescAlias)
                    ' Добавляем в список
                    lvwMembers.ListItems.Add(, , Name, , 17).Tag = 0
                    Exit Do
                End Select
            Loop
        Else
        ' Если не псевдоним
            Dim varInfo     As VARDESC
            Dim cnst        As Variant
            
            'Проходим по списку элементов переменных, констант
            For i = 0 To tAttr.cVars - 1
                ' Получаем описание
                ptr = typ.GetVarDesc(i)
                CopyMemory varInfo, ByVal ptr, LenB(varInfo)
                ' Получаем имя, описание
                typ.GetDocumentation varInfo.memid, Name, vbNullString, 0, vbNullString
                ' Получаем тип элемента и устанавливаем иконку
                Select Case varInfo.VARKIND
                Case VARKIND.VAR_CONST
                    icn = 21
                    cnst = Empty
                    VariantCopy cnst, ByVal varInfo.oInst_varValue
                    Select Case VarType(cnst)
                    Case VARENUM.VT_BSTR
                        Name = Name & " = " & Chr$(34) & cnst & Chr$(34)
                    Case VARENUM.VT_UI1, VARENUM.VT_I2, VARENUM.VT_I4
                        Name = Name & " = " & cnst & " (&H" & Hex(cnst) & ")"
                    Case Else
                        Name = Name & " = " & cnst
                    End Select
                Case VARKIND.VAR_PERINSTANCE
                    Select Case varInfo.elemdescVar.tdesc.vt
                    Case VARENUM.VT_CARRAY
                        ' Массив
                        Name = Name & GetArray(varInfo.elemdescVar.tdesc.pTypeDesc)
                    End Select
                    icn = 17
                Case Else: icn = 0
                End Select
                ' Добавляем в список
                Set itm = lvwMembers.ListItems.Add(, , Name, , icn)
                itm.Tag = varInfo.memid
                itm.SubItems(1) = GetTypeName(varInfo.elemdescVar.tdesc, typ)
                typ.ReleaseVarDesc ptr
            Next

            ' Проход по списку функций
            For i = 0 To tAttr.cFuncs - 1
                ' Получаем описание
                ptr = typ.GetFuncDesc(i)
                CopyMemory finf, ByVal ptr, LenB(finf)
                ' Получаем имя, описание
                typ.GetDocumentation finf.memid, Name, Desc, 0, vbNullString
                ' Добавляем в список
                Set itm = lvwMembers.ListItems.Add(, , Name, , 5)
                itm.SubItems(2) = Desc
                itm.SubItems(1) = GetTypeName(finf.elemdescFunc.tdesc, typ)
                typ.ReleaseFuncDesc ptr
            Next
        End If
    End If
End Sub

Private Sub Form_Resize()
    Dim h As Long
    h = ScaleHeight - stbStatus.Height
    If h <= 105 Or ScaleWidth <= 265 Then Exit Sub
    tvwLib.Move 5, 5, 250, h - 10
    lblImplements.Move 260, 5, ScaleWidth - 265
    lvwImplements.Move 260, 25, ScaleWidth - 265, 50
    lblMembers.Move 260, 80, ScaleWidth - 265
    lvwMembers.Move 260, 100, ScaleWidth - 265, h - 105
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuOpen_Click()
    Dim ofn     As OPENFILENAME, Out    As String, i        As Long
    
    ofn.nMaxFile = 260
    Out = String(260, vbNullChar)

    ofn.hwndOwner = hWnd
    ofn.lpstrTitle = StrPtr("Открыть файл")
    ofn.lpstrFile = StrPtr(Out)
    ofn.lStructSize = Len(ofn)
    ofn.lpstrFilter = StrPtr("Поддерживаемые файлы" & vbNullChar & "*.dll;*.ocx;*.exe;*.tlb" & vbNullChar)
    
    If GetOpenFileName(ofn) Then
        i = InStr(1, Out, vbNullChar, vbBinaryCompare)
        If i Then Out = Left$(Out, i - 1)
        LoadLib Out
    End If
End Sub

Private Sub tvwLib_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim i       As Long
    Dim icn     As Long
    Dim f       As Long
    Dim itm     As ListItem
    Dim iInf    As ImportInfo
    
    stbStatus.Panels.Clear
    ' Проверяем ключ
    Select Case Node.key
    Case "export"
        ' Клик по экспорту
        Dim eInf    As ExportInfo
        ' Очищаем список и добавляем колонки
        lvwMembers.ListItems.Clear
        lvwMembers.ColumnHeaders.Clear
        lvwMembers.ColumnHeaders.Add , , "Type", 16
        lvwMembers.ColumnHeaders.Add , , "Ord", 64
        lvwMembers.ColumnHeaders.Add , , "Name", 256
        lvwMembers.ColumnHeaders.Add , , "Entry point", 100
        ' Проход по всем элементам экспорта
        For i = 0 To natInfo.ExportCount - 1
            ' Получаем элемент импорта
            eInf = natInfo.Export(i)
            ' Если без имени, то иконка экспорта по ординалу
            If Len(eInf.Name) = 0 Then
                icn = 7
            Else
                ' Иконка для перенаправления
                If eInf.Forwarder Then icn = 6 Else icn = 5
            End If
            ' Добавляем в список тип
            Set itm = lvwMembers.ListItems.Add(, , , , icn)
            itm.ForeColor = vbBlack
            ' Добавляем в список ординал
            itm.SubItems(1) = eInf.Ordinal
            ' Если без имени
            If Len(eInf.Name) = 0 Then
                ' Нет имени
                itm.SubItems(2) = "-"
                ' Точка входа
                itm.SubItems(3) = FormatHex(Hex(eInf.EntryPoint))
            Else
                ' Если перенаправление
                If eInf.Forwarder Then
                    ' Без имени
                    itm.SubItems(2) = "-"
                    ' Точка входа
                    itm.SubItems(3) = eInf.Name
                Else
                    ' Имя
                    itm.SubItems(2) = eInf.Name
                    ' Точка входа
                    itm.SubItems(3) = FormatHex(Hex(eInf.EntryPoint))
                End If
            End If
        Next
    Case "lib"
        ' Очиска списка
        lvwMembers.ListItems.Clear
        lvwMembers.ColumnHeaders.Clear
    Case Else
        Select Case Node.Parent.key
        Case "import"
            ' Клик по импорту
            ' Очиска списка и добавление колонки с именем функции
            lvwMembers.ListItems.Clear
            lvwMembers.ColumnHeaders.Clear
            lvwMembers.ColumnHeaders.Add , , "Name", 300
            ' Находим индекс библиотеки
            i = Node.Index - Node.Parent.Index - 1
            ' Получаем список функций импортированных из этой библиотеки
            iInf = natInfo.Import(i)
            ' Проход по списку функций
            For f = 0 To iInf.Count - 1
                ' Добавляем в список
                lvwMembers.ListItems.Add , , iInf.Func(f), , 5
            Next
        Case "dimport"
            ' Клик по отложенному импорту
            ' Очиска списка и добавление колонки с именем функции
            lvwMembers.ListItems.Clear
            lvwMembers.ColumnHeaders.Clear
            lvwMembers.ColumnHeaders.Add , , "Name", 300
            ' Находим индекс библиотеки
            i = Node.Index - Node.Parent.Index - 1
            ' Получаем список функций импортированных из этой библиотеки
            iInf = natInfo.DelayImport(i)
            ' Проход по списку функций
            For f = 0 To iInf.Count - 1
                ' Добавляем в список
                lvwMembers.ListItems.Add , , iInf.Func(f), , 5
            Next
        Case "com"
            ' Клик по библиотеке типов
            Dim typ     As ITypeInfo
            lvwMembers.ListItems.Clear
            lvwMembers.ColumnHeaders.Clear
            lvwMembers.ColumnHeaders.Add , , "Name", 300
            lvwMembers.ColumnHeaders.Add , , "Type", 200
            lvwMembers.ColumnHeaders.Add , , "Description", 4000
            ' Находим индекс типа
            i = Node.Index - Node.Parent.Index - 1
            ' Очистка списков
            lvwImplements.ListItems.Clear
            lvwMembers.ListItems.Clear
            Set typ = Lib.GetTypeInfo(i)
            LoadType typ
        End Select
    End Select
End Sub
' Получить имя стандартного типа
Private Function GetTypeName(td As TYPEDESC, Optional typ As ITypeInfo) As String
    Dim Desc    As TYPEDESC
    Dim ardsc   As ARRAYDESC
    Dim ti      As ITypeInfo
    
    Select Case td.vt
    Case 1: GetTypeName = "NULL"
    Case 2: GetTypeName = "Integer"
    Case 3: GetTypeName = "Long"
    Case 4: GetTypeName = "Single"
    Case 5: GetTypeName = "Double"
    Case 6: GetTypeName = "Currency"
    Case 7: GetTypeName = "Date"
    Case 8: GetTypeName = "String"
    Case 9: GetTypeName = "Object(IDispatch)"
    Case 10: GetTypeName = "Error"
    Case 11: GetTypeName = "Boolean"
    Case 12: GetTypeName = "Variant"
    Case 13: GetTypeName = "IUnknown"
    Case 14: GetTypeName = "Decimal"
    Case 16: GetTypeName = "SByte"
    Case 17: GetTypeName = "Byte"
    Case 18: GetTypeName = "UShort"
    Case 19: GetTypeName = "ULong"
    Case 20: GetTypeName = "Int64"
    Case 21: GetTypeName = "UInt64"
    Case 22: GetTypeName = "Long" ' "Int"
    Case 23: GetTypeName = "UInt"
    Case 24: GetTypeName = "Void"
    Case 25: GetTypeName = "HRESULT"
    Case 26
        ' PTR
        CopyMemory Desc, ByVal td.pTypeDesc, Len(Desc)
        GetTypeName = GetTypeName(Desc, typ)
    Case 27: GetTypeName = "SAFEARRAY"
    Case 28
        ' CArray
        CopyMemory ardsc, ByVal td.pTypeDesc, Len(ardsc)
        GetTypeName = GetTypeName(ardsc.tdescElem, typ)
    Case 29
        ' UDT
        Set ti = typ.GetRefTypeInfo(td.pTypeDesc)
        ti.GetDocumentation -1, GetTypeName, vbNullString, 0, vbNullString
    Case 30: GetTypeName = "LPStr"
    Case 31: GetTypeName = "LPWStr"
    Case 36: GetTypeName = "RECORD"
    Case 37: GetTypeName = "INT_PTR"
    Case 38: GetTypeName = "UINT_PTR"
    Case 64: GetTypeName = "FILETIME"
    Case 65: GetTypeName = "Blob"
    Case 66: GetTypeName = "Stream"
    Case 67: GetTypeName = "Storage"
    Case 68: GetTypeName = "STREAMED_OBJECT"
    Case 69: GetTypeName = "STORED_OBJECT"
    Case 70: GetTypeName = "BLOB_OBJECT"
    Case 71: GetTypeName = "CF"
    Case 72: GetTypeName = "CLSID"
    Case 73: GetTypeName = "VERSIONED_STREAM"
    Case &HFFF: GetTypeName = "BSTR_BLOB"
    End Select
End Function

' Получить размерности массива
Private Function GetArray(ByVal ptr As Long) As String
    Dim arr     As ARRAYDESC
    Dim bnd()   As SAFEARRAYBOUND
    Dim i       As Long
    
    CopyMemory arr, ByVal ptr, Len(arr)
    ReDim bnd(arr.cDims)
    CopyMemory bnd(0), ByVal ptr + LenB(arr), Len(bnd(0)) * (arr.cDims + 1)
    
    GetArray = "("
    For i = 0 To UBound(bnd)
        If i Then GetArray = GetArray & ", "
        If bnd(i).lLbound = 0 Then
            GetArray = GetArray & bnd(i).cElements - 1
        Else
            GetArray = GetArray & bnd(i).lLbound & " To " & bnd(i).cElements + bnd(i).lLbound - 1
        End If
    Next
    GetArray = GetArray & ")"
End Function

Private Function FormatHex(V As String) As String
    Dim i       As Long
    FormatHex = V
    For i = Len(V) To 7
        FormatHex = "0" & FormatHex
    Next
    FormatHex = "&H" & FormatHex
End Function

Private Function GetFileTitle(Path As String, Optional UseExtension As Boolean = False) As String
    Dim L As Long, P As Long
    L = InStrRev(Path, "\")
    If UseExtension Then P = Len(Path) + 1 Else P = InStrRev(Path, ".")
    If P > L Then
        L = IIf(L = 0, 1, L + 1)
        GetFileTitle = Mid$(Path, L, P - L)
    ElseIf P = L Then
        GetFileTitle = Path
    Else
        GetFileTitle = Mid$(Path, L + 1)
    End If
End Function
