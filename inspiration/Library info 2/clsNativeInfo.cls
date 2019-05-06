VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNativeInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Класс clsNativeInfo.cls для получения информации об импорте, экспорте
' © Кривоус Анатолий Анатольевич (The trick), 2014

Private Type IMAGE_DOS_HEADER
    e_magic         As Integer
    e_cblp          As Integer
    e_cp            As Integer
    e_crlc          As Integer
    e_cparhdr       As Integer
    e_minalloc      As Integer
    e_maxalloc      As Integer
    e_ss            As Integer
    e_sp            As Integer
    e_csum          As Integer
    e_ip            As Integer
    e_cs            As Integer
    e_lfarlc        As Integer
    e_ovno          As Integer
    e_res(0 To 3)   As Integer
    e_oemid         As Integer
    e_oeminfo       As Integer
    e_res2(0 To 9)  As Integer
    e_lfanew        As Long
End Type
Private Type IMAGE_DATA_DIRECTORY
    VirtualAddress  As Long
    Size            As Long
End Type
Private Type IMAGE_OPTIONAL_HEADER
    Magic                           As Integer
    MajorLinkerVersion              As Byte
    MinorLinkerVersion              As Byte
    SizeOfCode                      As Long
    SizeOfInitializedData           As Long
    SizeOfUnitializedData           As Long
    AddressOfEntryPoint             As Long
    BaseOfCode                      As Long
    BaseOfData                      As Long
    ImageBase                       As Long
    SectionAlignment                As Long
    FileAlignment                   As Long
    MajorOperatingSystemVersion     As Integer
    MinorOperatingSystemVersion     As Integer
    MajorImageVersion               As Integer
    MinorImageVersion               As Integer
    MajorSubsystemVersion           As Integer
    MinorSubsystemVersion           As Integer
    W32VersionValue                 As Long
    SizeOfImage                     As Long
    SizeOfHeaders                   As Long
    CheckSum                        As Long
    SubSystem                       As Integer
    DllCharacteristics              As Integer
    SizeOfStackReserve              As Long
    SizeOfStackCommit               As Long
    SizeOfHeapReserve               As Long
    SizeOfHeapCommit                As Long
    LoaderFlags                     As Long
    NumberOfRvaAndSizes             As Long
End Type
Private Type IMAGE_FILE_HEADER
    Machine                 As Integer
    NumberOfSections        As Integer
    TimeDateStamp           As Long
    PointerToSymbolTable    As Long
    NumberOfSymbols         As Long
    SizeOfOptionalHeader    As Integer
    Characteristics         As Integer
End Type
Private Type IMAGE_NT_HEADERS
    Signature               As Long
    FileHeader              As IMAGE_FILE_HEADER
    OptionalHeader          As IMAGE_OPTIONAL_HEADER
End Type
Private Type IMAGE_SECTION_HEADER
    SectionName(7)          As Byte
    VirtualSize             As Long
    VirtualAddress          As Long
    SizeOfRawData           As Long
    PointerToRawData        As Long
    PointerToRelocations    As Long
    PointerToLinenumbers    As Long
    NumberOfRelocations     As Integer
    NumberOfLinenumbers     As Integer
    Characteristics         As Long
End Type
Private Type IMAGE_IMPORT_DESCRIPTOR
    Characteristics         As Long
    TimeDateStamp           As Long
    ForwarderChain          As Long
    pName                   As Long
    FirstThunk              As Long
End Type
Private Type IMAGE_EXPORT_DIRECTORY
    Characteristics         As Long
    TimeDateStamp           As Long
    MajorVersion            As Integer
    MinorVersion            As Integer
    pName                   As Long
    Base                    As Long
    NumberOfFunctions       As Long
    NumberOfNames           As Long
    AddressOfFunctions      As Long
    AddressOfNames          As Long
    AddressOfNameOrdinals   As Long
End Type
Private Type IMAGE_DELAY_IMPORT_DESCRIPTOR
    grAttr      As Long
    rvaDllName  As Long
    phMod       As Long
    pIAT        As Long
    pINT        As Long
    pBoundIAT   As Long
    pUnloadIAT  As Long
    dwTimeStamp As Long
End Type

Private Type ptDat
    Prv1    As Long
    Prv2    As Long
End Type

Private Const IMAGE_DIRECTORY_ENTRY_EXPORT          As Long = 0
Private Const IMAGE_DIRECTORY_ENTRY_IMPORT          As Long = 1
Private Const IMAGE_DIRECTORY_ENTRY_DELAY_IMPORT    As Long = 13
Private Const GENERIC_READ                          As Long = &H80000000
Private Const FILE_SHARE_READ                       As Long = &H1
Private Const PAGE_READONLY                         As Long = 2&
Private Const FILE_MAP_READ                         As Long = &H4
Private Const INVALID_HANDLE_VALUE                  As Long = -1
Private Const OPEN_EXISTING                         As Long = 3
Private Const FILE_ATTRIBUTE_NORMAL                 As Long = &H80
Private Const IMAGE_DOS_SIGNATURE                   As Long = &H5A4D
Private Const IMAGE_NT_SIGNATURE                    As Long = &H4550&
Private Const IMAGE_NT_OPTIONAL_HDR32_MAGIC         As Long = &H10B&

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetMem4 Lib "msvbvm60" (Src As Any, Dst As Any) As Long
Private Declare Function lstrcpyn Lib "kernel32" Alias "lstrcpynA" (lpString1 As Any, lpString2 As Any, ByVal iMaxLength As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (lpString As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingW" (ByVal hFile As Long, lpFileMappingAttributes As Any, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As Long) As Long
Private Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" (ByVal lpBaseAddress As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function IsBadReadPtr Lib "kernel32" (lp As Any, ByVal ucb As Long) As Long

Public Enum ERROR_CODE
    OK
    ERROR_LOADING
    INCORRECT_FILE
    ONLY32BITS
    WIN32ERROR
    IN_HEADER = &H10&
    IN_EXPORT = &H20&
    IN_IMPORT = &H30&
    IN_DELAY_IMPORT = &H40&
End Enum

Private mError              As ERROR_CODE
Private mBase               As Long
Private mNTHdr              As IMAGE_NT_HEADERS
Private mDatDirCount        As Long
Private mDatDir()           As IMAGE_DATA_DIRECTORY
Private mExport()           As ExportInfo
Private mExportCount        As Long
Private mImport()           As ImportInfo
Private mImportCount        As Long
Private mDelayImport()      As ImportInfo
Private mDelayImportCount   As Long
Private mModName            As String

Dim lpFirstSection          As Long

Friend Property Get ErrorCode() As ERROR_CODE
    ErrorCode = mError
End Property
Friend Property Get ErrorText(ByVal ErrorNumber As Long) As String
    Select Case ErrorNumber And &HF
    Case ERROR_CODE.ERROR_LOADING:      ErrorText = "Error loading"
    Case ERROR_CODE.WIN32ERROR:         ErrorText = "Win32 error"
    Case ERROR_CODE.ONLY32BITS:         ErrorText = "Supported only 32bits"
    Case ERROR_CODE.INCORRECT_FILE:     ErrorText = "Incorrect file"
    End Select
    ' Где произошла ошибка
    Select Case ErrorNumber And &HF0
    Case ERROR_CODE.IN_HEADER:          ErrorText = ErrorText & " header."
    Case ERROR_CODE.IN_EXPORT:          ErrorText = ErrorText & " export table."
    Case ERROR_CODE.IN_IMPORT:          ErrorText = ErrorText & " import table."
    Case ERROR_CODE.IN_DELAY_IMPORT:    ErrorText = ErrorText & " delay-import table."
    Case Else:                          ErrorText = ErrorText & "."
    End Select
End Property
Friend Property Get ModuleName() As String
    If mError Then Exit Property
    ModuleName = mModName
End Property
Friend Property Get Export(ByVal Index As Long) As ExportInfo
    If CBool(mError) Or Index < 0 Or Index > mExportCount Then
        Err.Raise 5, , "Invalid export index": Exit Property
    End If
    Export = mExport(Index)
End Property
Friend Property Get ExportCount() As Long
    If mError Then Exit Property
    ExportCount = mExportCount
End Property
Friend Property Get Import(ByVal Index As Long) As ImportInfo
    If CBool(mError) Or Index < 0 Or Index > mImportCount Then
        Err.Raise 5, , "Invalid import index": Exit Property
    End If
    Import = mImport(Index)
End Property
Friend Property Get ImportCount() As Long
    If mError Then Exit Property
    ImportCount = mImportCount
End Property
Friend Property Get DelayImport(ByVal Index As Long) As ImportInfo
    If CBool(mError) Or Index < 0 Or Index > mDelayImportCount Then
        Err.Raise 5, , "Invalid import index": Exit Property
    End If
    DelayImport = mDelayImport(Index)
End Property
Friend Property Get DelayImportCount() As Long
    If mError Then Exit Property
    DelayImportCount = mDelayImportCount
End Property
' Извлеч информацию из PE файла
Friend Function Extract(FileName As String) As Boolean
    ' Сброс ошибки
    mError = OK
    ' Загружаем образ
    If Not LoadPE(FileName) Then Exit Function
    ' Загружаем заголовок
    If Not GetHeader() Then ClosePE: Exit Function
    ' Загружаем таблицу экспорта
    If Not GetExport() Then ClosePE: Exit Function
    ' Загружаем таблицу имопорта
    If Not GetImport() Then ClosePE: Exit Function
    ' Загружаем таблицу отложенного импорта
    If Not GetDelayImport() Then ClosePE: Exit Function
    ' Закрываем файл
    ClosePE
    Extract = True
End Function

' Проецирует файл на АП процесса, возвращает статус
Private Function LoadPE(FileName As String) As Boolean
    Dim hFile       As Long
    Dim hMap        As Long
    ' Открываем файл
    hFile = CreateFile(StrPtr(FileName), GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    ' Если ошибка, выходим
    If hFile = INVALID_HANDLE_VALUE Then mError = ERROR_LOADING: Exit Function
    ' Создаем файловое отображение
    hMap = CreateFileMapping(hFile, ByVal 0&, PAGE_READONLY, 0, 0, 0)
    ' Закрываем дескриптор открытого файла
    CloseHandle hFile
    ' Если ошибка, то выходим
    If hMap = 0 Then mError = WIN32ERROR: Exit Function
    ' Проецируем файл
    mBase = MapViewOfFile(hMap, FILE_MAP_READ, 0, 0, 0)
    ' Закрываем дескриптор отображения
    CloseHandle hMap
    ' Проверка указателя
    If mBase = 0 Then mError = WIN32ERROR: Exit Function
    ' Возвращаем успех
    LoadPE = True
End Function
' Закрыть файл
Private Sub ClosePE()
    UnmapViewOfFile mBase
    mBase = 0
End Sub
' Загрузить  заголовок и описатели каталогов
Private Function GetHeader() As Boolean
    Dim dosHdr()    As IMAGE_DOS_HEADER
    Dim dPtr()      As Long
    Dim dOld        As ptDat
    ReDim dosHdr(0):    ReDim dPtr(0)
    ' Создаем указатель на заголовок
    dOld = PtGet(dPtr, Not Not dosHdr): dPtr(0) = mBase
    ' Проверка корректности указателя
    If IsBadReadPtr(ByVal dPtr(0), Len(dosHdr(0))) = 0 Then
        ' Проверяем сигнатуру MZ и выравнивание
        If dosHdr(0).e_magic = IMAGE_DOS_SIGNATURE And (dosHdr(0).e_lfanew And &H3) = 0 Then
            ' Проверяем PE заголовок
            Dim ntHdr()    As IMAGE_NT_HEADERS
            Dim nPtr()     As Long
            Dim nOld        As ptDat
            ReDim ntHdr(0):    ReDim nPtr(0)
            ' Устанавливаем указатель на PE заголовок
            nOld = PtGet(nPtr, Not Not ntHdr): nPtr(0) = mBase + dosHdr(0).e_lfanew
            ' Проверка корректности указателя
            If IsBadReadPtr(ByVal nPtr(0), Len(ntHdr(0))) = 0 Then
                ' Проверка сигнатуры PE и размера опционального заголовка
                If ntHdr(0).Signature = IMAGE_NT_SIGNATURE And _
                   ntHdr(0).FileHeader.SizeOfOptionalHeader >= Len(ntHdr(0).OptionalHeader) Then
                    'Проверка разрядности приложения
                    If ntHdr(0).OptionalHeader.Magic = IMAGE_NT_OPTIONAL_HDR32_MAGIC Then
                        ' Поддерживается только 32-разрядные приложения
                        If ntHdr(0).OptionalHeader.NumberOfRvaAndSizes >= 0 Then
                            Dim dirCt       As Long
                            Dim lpDatDir    As Long
                            ' Определяем смещение описателей каталогов
                            lpDatDir = nPtr(0) + Len(ntHdr(0))
                            ' Определяем количество каталогов
                            dirCt = ntHdr(0).OptionalHeader.NumberOfRvaAndSizes
                            ' Определяем смещение описателей секций
                            lpFirstSection = lpDatDir + Len(mDatDir(0)) * dirCt
                            ' Если есть каталоги
                            If dirCt Then
                                ' Проверяем корректность указателя
                                If IsBadReadPtr(ByVal lpDatDir, Len(mDatDir(0)) * dirCt) = 0 Then
                                    ReDim mDatDir(dirCt - 1)
                                    CopyMemory mDatDir(0), ByVal lpDatDir, Len(mDatDir(0)) * dirCt
                                    mDatDirCount = dirCt
                                    mNTHdr = ntHdr(0)
                                    GetHeader = True
                                Else: mError = INCORRECT_FILE Or IN_HEADER
                                End If
                            Else
                                Erase mDatDir()
                                mDatDirCount = 0
                                mNTHdr = ntHdr(0)
                                GetHeader = True
                            End If
                        Else: mError = INCORRECT_FILE Or IN_HEADER
                        End If
                    Else: mError = ONLY32BITS Or IN_HEADER
                    End If
                Else: mError = INCORRECT_FILE Or IN_HEADER
                End If
            Else: mError = INCORRECT_FILE Or IN_HEADER
            End If
            ' Освобождаем указатель
            PtRelease nPtr, nOld
        Else: mError = INCORRECT_FILE Or IN_HEADER
        End If
    Else: mError = INCORRECT_FILE Or IN_HEADER
    End If
    ' Освобождаем указатель
    PtRelease dPtr, dOld
End Function
' Загрузить таблицу экспорта
Private Function GetExport() As Boolean
    Dim ptr     As Long
    Dim rva     As Long
    Dim sz      As Long
    ' Очищаем список функций
    Erase mExport(): mExportCount = 0
    ' Если нет таблицы экспорта, выходим
    If mDatDirCount = 0 Then GetExport = True: Exit Function
    ' Получаем RAW секции
    rva = mDatDir(IMAGE_DIRECTORY_ENTRY_EXPORT).VirtualAddress
    sz = mDatDir(IMAGE_DIRECTORY_ENTRY_EXPORT).Size
    ptr = RVA2RAW(rva) + mBase
    If sz > 0 And rva > 0 Then
        Dim tbl         As IMAGE_EXPORT_DIRECTORY
        ' Проверяем соответствие размеров
        If Len(tbl) > sz Then mError = INCORRECT_FILE Or IN_EXPORT: Exit Function
        ' Проверяем корректность указателя и размеров
        If IsBadReadPtr(ByVal ptr, sz) Then mError = INCORRECT_FILE Or IN_EXPORT: Exit Function
        ' Копируем таблицу
        CopyMemory tbl, ByVal ptr, Len(tbl)
        ' Получаем имя модуля
        mModName = GetString(RVA2RAW(tbl.pName) + mBase)
        ' Получаем количество экспортируемых функций
        If tbl.NumberOfFunctions < 0 Then mError = INCORRECT_FILE Or IN_EXPORT: Exit Function
        If tbl.NumberOfFunctions Then
            ' Если есть экспортируемые функции
            Dim i           As Long, N          As Long, f          As Long
            Dim aPt()       As Long, adr()      As Long, ao         As ptDat
            Dim oPt()       As Long, ord()      As Integer, oo      As ptDat
            Dim nPt()       As Long, nam()      As Long, no         As ptDat
            Dim rawNam      As Long, rawOrd     As Long
            ReDim aPt(0):   ReDim adr(0)
            ReDim oPt(0):   ReDim ord(0)
            ReDim nPt(0):   ReDim nam(0)
            ' Выделяем место под таблицу
            ReDim mExport(tbl.NumberOfFunctions - 1)
            ' Установка указателя на начало функций
            ao = PtGet(aPt, Not Not adr):       aPt(0) = RVA2RAW(tbl.AddressOfFunctions) + mBase
            ' Установка указателя на список ординалов
            oo = PtGet(oPt, Not Not ord)
            ' Установка указателя для экспртируемых имен
            no = PtGet(nPt, Not Not nam)
            rawNam = RVA2RAW(tbl.AddressOfNames)
            rawOrd = RVA2RAW(tbl.AddressOfNameOrdinals)
            ' Проход по списку функций
            For i = 0 To tbl.NumberOfFunctions - 1
                ' Записываем ординал
                mExport(f).Ordinal = i + tbl.Base
                ' Проверка типа указателя
                Select Case True
                Case adr(0) = 0
                    ' Не используется
                Case (adr(0) >= rva And adr(0) < rva + sz)
                    ' Перенаправление
                    mExport(f).Name = GetString(RVA2RAW(adr(0)) + mBase)
                    ' Проверка корректности
                    If Len(mExport(f).Name) = 0 Then
                        mError = INCORRECT_FILE Or IN_EXPORT
                        Exit For
                    End If
                    mExport(f).Forwarder = True
                    f = f + 1
                Case Else
                    ' Обычный
                    mExport(f).EntryPoint = adr(0)
                    oPt(0) = rawOrd + mBase
                    ' Проход по таблице имен в поисках имени
                    For N = 0 To tbl.NumberOfNames - 1
                        If ord(0) = i Then
                            ' Нашли имя, соответствующее ординалу
                            nPt(0) = rawNam + N * 4 + mBase
                            mExport(f).Name = GetString(RVA2RAW(nam(0)) + mBase)
                            ' Проверка корректности
                            If Len(mExport(f).Name) = 0 Then
                                mError = INCORRECT_FILE Or IN_EXPORT
                                i = tbl.NumberOfFunctions
                            End If
                            Exit For
                        End If
                        ' Указатель на следующий ординал
                        oPt(0) = oPt(0) + 2
                    Next
                    f = f + 1
                End Select
                ' Указатель на следующий элемент
                aPt(0) = aPt(0) + 4
            Next
            mExportCount = f
            ' Восстановление указателей
            PtRelease aPt, ao
            PtRelease oPt, oo
            PtRelease nPt, no
        Else: GetExport = True: Exit Function
        End If
    Else: GetExport = True: Exit Function
    End If
    GetExport = mError = OK
End Function
' Загрузить таблицу импорта
Private Function GetImport() As Boolean
    Dim pDsc()     As Long, dsc()     As IMAGE_IMPORT_DESCRIPTOR, od      As ptDat
    Dim ptr     As Long
    Dim rva     As Long
    Dim sz      As Long
    ReDim pDsc(0):  ReDim dsc(0)
    ' Очищаем список функций
    Erase mImport(): mImportCount = 0
    ' Если нет таблицы импорта, выходим
    If mDatDirCount <= IMAGE_DIRECTORY_ENTRY_IMPORT Then GetImport = True: Exit Function
    ' Получаем VA секции
    rva = mDatDir(IMAGE_DIRECTORY_ENTRY_IMPORT).VirtualAddress
    sz = mDatDir(IMAGE_DIRECTORY_ENTRY_IMPORT).Size
    ptr = RVA2RAW(rva) + mBase
    ' Если присутствует
    If sz > 0 And rva > 0 Then
        Dim i       As Long
        ' Устанавливаем указатель на описатель
        od = PtGet(pDsc, Not Not dsc): pDsc(0) = ptr
        ' Проход по описателям
        Do Until dsc(0).Characteristics = 0 And _
                 dsc(0).FirstThunk = 0 And _
                 dsc(0).ForwarderChain = 0 And _
                 dsc(0).pName = 0 And _
                 dsc(0).TimeDateStamp = 0
            ' Если вышли за границу переопределяем массив
            If i = mImportCount Then
                mImportCount = mImportCount + 10
                ReDim Preserve mImport(mImportCount)
            End If
            ' Проверяем корректность имени
            If dsc(0).pName = 0 Then
                mError = INCORRECT_FILE Or IN_IMPORT
                Exit Do
            End If
            ' Получаем имя библиотеки
            mImport(i).Name = GetString(RVA2RAW(dsc(0).pName) + mBase)
            ' Проверяем корректность
            If Len(mImport(i).Name) = 0 Then
                mError = INCORRECT_FILE Or IN_IMPORT
                Exit Do
            End If
            ' Получаем имена функций в данной библиотеке
            If Not GetFunctions(IIf(dsc(0).Characteristics, dsc(0).Characteristics, dsc(0).FirstThunk), mImport(i).Func()) Then
                mError = mError Or IN_IMPORT
                Exit Do
            End If
            ' Получаем реальное количество функций
            If Not Not mImport(i).Func() Then
                mImport(i).Count = UBound(mImport(i).Func) + 1
            Else: mImport(i).Count = 0
            End If
            ' Переход к следующему элементу
            i = i + 1
            pDsc(0) = pDsc(0) + Len(dsc(0))
        Loop
        ' Получаем количество
        mImportCount = i
        ' Освобождаем указатель
        PtRelease pDsc, od
    Else: GetImport = True: Exit Function
    End If
    GetImport = mError = OK
End Function
' Загрузить таблицу отложенного импорта
Private Function GetDelayImport() As Boolean
    Dim pDsc()      As Long, dsc()     As IMAGE_DELAY_IMPORT_DESCRIPTOR, od      As ptDat
    Dim ptr     As Long
    Dim rva     As Long
    Dim sz      As Long
    ReDim pDsc(0):  ReDim dsc(0)
    ' Очищаем список функций
    Erase mDelayImport(): mDelayImportCount = 0
    ' Если нет таблицы импорта, выходим
    If mDatDirCount <= IMAGE_DIRECTORY_ENTRY_DELAY_IMPORT Then GetDelayImport = True: Exit Function
    ' Получаем VA секции
    rva = mDatDir(IMAGE_DIRECTORY_ENTRY_DELAY_IMPORT).VirtualAddress
    sz = mDatDir(IMAGE_DIRECTORY_ENTRY_DELAY_IMPORT).Size
    ptr = RVA2RAW(rva) + mBase
    ' Если присутствует
    If sz > 0 And rva > 0 Then
        Dim i       As Long
        ' Устанавливаем указатель на описатель
        od = PtGet(pDsc, Not Not dsc): pDsc(0) = ptr
        ' Проход по описателям
        Do While dsc(0).rvaDllName
            ' Если вышли за границу переопределяем массив
            If i = mDelayImportCount Then
                mDelayImportCount = mDelayImportCount + 10
                ReDim Preserve mDelayImport(mDelayImportCount)
            End If
            ' Если содержит RVA
            If dsc(0).grAttr And 1 Then
                ' Получаем имя библиотеки
                mDelayImport(i).Name = GetString(RVA2RAW(dsc(0).rvaDllName) + mBase)
                ' Проверка корректности имени
                If Len(mDelayImport(i).Name) = 0 Then
                    mError = INCORRECT_FILE Or IN_DELAY_IMPORT
                    Exit Do
                End If
                ' Получаем список функций
                GetFunctions dsc(0).pINT, mDelayImport(i).Func()
            Else
                ' Получаем имя библиотеки
                mDelayImport(i).Name = GetString(RVA2RAW(dsc(0).rvaDllName - mBase) + mBase)
                ' Проверка корректности имени
                If Len(mDelayImport(i).Name) = 0 Then
                    mError = INCORRECT_FILE Or IN_DELAY_IMPORT
                    Exit Do
                End If
                ' Получаем список функций
                If Not GetFunctions(dsc(0).pINT, mDelayImport(i).Func()) Then
                    mError = mError Or IN_DELAY_IMPORT
                    Exit Do
                End If
            End If
            ' Получаем реальное количество функций
            If Not Not mDelayImport(i).Func() Then
                mDelayImport(i).Count = UBound(mDelayImport(i).Func) + 1
            Else: mDelayImport(i).Count = 0
            End If
            ' Переход к следующему элементу
            i = i + 1
            pDsc(0) = pDsc(0) + Len(dsc(0))
        Loop
        ' Получаем количество
        mDelayImportCount = i
        ' Освобождаем указатель
        PtRelease pDsc, od
    Else: GetDelayImport = True: Exit Function
    End If
    GetDelayImport = mError = OK
End Function
' Извлечь имена функций
Private Function GetFunctions(ByVal rva As Long, ByRef Names() As String) As Boolean
    Dim pTnk()      As Long, thnk()         As Long, ot     As ptDat
    Dim i           As Long
    Dim Count       As Long
    ReDim pTnk(0):  ReDim thnk(0)
    ' Устанавливаем указатель на начало IMAGE_THUNK_DATA
    ot = PtGet(pTnk, Not Not thnk)
    pTnk(0) = RVA2RAW(rva) + mBase
    ' Проход по именам
    Do While thnk(0)
        ' Если вышли за границу переопределяем массив
        If i = Count Then
            Count = Count + 20
            ReDim Preserve Names(Count)
        End If
        ' Проверяем тип импорта
        If thnk(0) < 0 Then
            ' Импорт по ординалу
            Names(i) = "#" & CStr(thnk(0) And &H7FFFFFFF)
        Else
            ' Импорт по имени
            Names(i) = GetString(RVA2RAW(thnk(0)) + mBase + 2)
            ' Проверка корректности имени
            If Len(Names(i)) = 0 Then
                mError = INCORRECT_FILE
                Exit Do
            End If
        End If
        i = i + 1
        pTnk(0) = pTnk(0) + 4
    Loop
    ' Получаем реальное количество элементов
    If i Then ReDim Preserve Names(i - 1)
    ' Освобождаем указатель
    PtRelease pTnk, ot
    GetFunctions = mError = OK
End Function
' Получить строку
Private Function GetString(ByVal ptr As Long) As String
    Dim L   As Long
    L = lstrlen(ByVal ptr)
    If L Then
        GetString = Space(L)
        If lstrcpyn(ByVal GetString, ByVal ptr, L + 1) = 0 Then GetString = vbNullString
    End If
End Function
' RVA to RAW
Private Function RVA2RAW(ByVal rva As Long) As Long
    Dim i       As Long
    Dim pSec()  As Long, sec()      As IMAGE_SECTION_HEADER, os     As ptDat
    ' Установка указателя на начало секций
    ReDim pSec(0):  ReDim sec(0)
    os = PtGet(pSec, Not Not sec): pSec(0) = lpFirstSection
    ' Проход по всем секциям
    For i = 0 To mNTHdr.FileHeader.NumberOfSections - 1
        ' Если RVA находится внутри секции
        If rva >= sec(0).VirtualAddress And rva < sec(0).VirtualAddress + sec(0).VirtualSize Then
            RVA2RAW = sec(0).PointerToRawData + (rva - sec(0).VirtualAddress)
            ' Освобождаем указатель
            PtRelease pSec, os
            Exit Function
        End If
        pSec(0) = pSec(0) + Len(sec(0))
    Next
    RVA2RAW = rva
    ' Освобождаем указатель
    PtRelease pSec, os
End Function
' Создать указатель
Private Function PtGet(Pointer() As Long, ByVal VarAddr As Long) As ptDat
    Dim i As Long
    i = (Not Not Pointer) + &HC
    GetMem4 ByVal i, PtGet.Prv1
    GetMem4 VarAddr + &HC, ByVal i
    PtGet.Prv2 = Pointer(0)
End Function
' Освободить указатель
Private Sub PtRelease(Pointer() As Long, prev As ptDat)
    Pointer(0) = prev.Prv2
    GetMem4 prev.Prv1, ByVal (Not Not Pointer) + &HC
End Sub
