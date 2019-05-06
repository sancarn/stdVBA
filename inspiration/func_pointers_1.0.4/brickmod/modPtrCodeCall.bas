Attribute VB_Name = "FuncPointerSupport"
Option Explicit
' ******************************************************************************
' *     Fire-Lines © 2008-2012                                                 *
' *     Модуль:                                                                *
' *         FunctPointerSupport                                                *
' *     Описание:                                                              *
' *         Функции в этом модуле обеспечивают поддержку указателей на функции *
' *         через Declare.                                                     *
' *     Автор:                                                                 *
' *         Владислав Петровский (firehacker)                                  *
' *     История изменений:                                                     *
' *         *   2008-06-29  firehacker  Файл создан.                           *
' *         *   2010-02-02  firehacker  Реализована поддержка указателей       *
' *                                     в скомпилированном проекте.            *
' *         *   2011-02-27  firehacker  Реализована поддержка указателей       *
' *                                     в режиме отладки в среде разработки.   *
' *         *   2011-03-06  firehacker  Исправлен баг: для ANSI-строк исполь-  *
' *                                     зовался SysAllocString, который на     *
' *                                     входе ожидает увидеть UCS2-строку.     *
' *                                     В результате указатели с именем с      *
' *                                     нечётной длиной могли не работать.     *
' *         *   2011-03-10  firehacker  Исправлен баг: LocalResolver, содер-   *
' *                                     жал ветку выполнения, которая ничего   *
' *                                     не делала и не возвращала никакого     *
' *                                     значения — если функция импортируется  *
' *                                     не по ординалу, но имя не нашлось в    *
' *                                     наборе Undiscovered-указателей.        *
' *         *   2011-03-10  firehacker  Исправлен баг: некорректная обработка  *
' *                                     ошибки «Specified pointer is not set». *
' *         *   2012-03-30  firehacker  Исправлен баг: после отключения модуля *
' *                                     с помощью MagicPointersOnOff, и после- *
' *                                     дующего включения, указатели, вызов    *
' *                                     через которые уже осуществлялся ранее, *
' *                                     становились «не изменяемыми», то есть  *
' *                                     сохраняли оставшееся с прошлого раза   *
' *                                     значение, даже после присвоения нового.*
' *     Текущая версия: 1.0.4                                                  *
' *                                                                            *
' ******************************************************************************

'
' Если вы работаете под VB5 (а не 6), замените следующую константу:
'
#Const VB_VERSION = 6 ' = 5


'
' Объявление функций, которые используются и в скомпилированном файле
' и в режиме отладки под IDE.
'
Private Declare Function SysAllocString Lib "oleaut32" (ByVal lpszString As Long) As String
Private Declare Sub SysFreeString Lib "oleaut32" (ByVal lpbstr As Long)
Private Declare Function lstrcmpiA Lib "kernel32" (ByVal lpStr1 As Long, ByVal lpStr2 As String) As Long

'
' Объявление функций, которые будут использоваться только в скомпилированном
' файле.
'
#If VB_VERSION = 6 Then
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal pdwFrom As Long, ByRef pdwTo As Long)
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal pdwTo As Long, ByVal dwNewVal As Long)
Private Declare Function DllFunctionCall Lib "msvbvm60" (ByRef pLookupEntry As EB_DELAYED_IMPORT_LOOKUP_ENTRY) As Long
Private Declare Function AryPtr Lib "msvbvm60" Alias "VarPtr" (Ary() As Any) As Long
#ElseIf VB_VERSION = 5 Then
Private Declare Sub GetMem4 Lib "msvbvm50" (ByVal pdwFrom As Long, ByRef pdwTo As Long)
Private Declare Sub PutMem4 Lib "msvbvm50" (ByVal pdwTo As Long, ByVal dwNewVal As Long)
Private Declare Function DllFunctionCall Lib "msvbvm50" (ByRef pLookupEntry As EB_DELAYED_IMPORT_LOOKUP_ENTRY) As Long
Private Declare Function AryPtr Lib "msvbvm60" Alias "VarPtr" (Ary() As Any) As Long
#End If
Private Declare Sub VirtualProtect Lib "kernel32" (ByVal pRegion As Long, ByVal lSize As Long, ByVal lNewProtection As Long, ByRef lpOldProtection As Long)

'
' Объявление функций, которые будут использоваться только в режиме отладки
' в среде (IDE). Поэтому не нужно думать, что скомпилированный файл становится
' зависимым от vbaX.dll — не становится!
'
Private Declare Sub RtlMoveMemory Lib "kernel32" (dst As Any, src As Any, ByVal sz As Long)
#If VB_VERSION = 6 Then
Private Declare Sub GetMem1 Lib "msvbvm60" (ByVal pdwFrom As Long, ByRef pdwTo As Byte)
Private Declare Sub EbGetExecutingProj Lib "vba6" (ByRef hProject As Long)
Private Declare Sub TipReleaseProject Lib "vba6" (ByVal hProject As Long)
Private Declare Sub TipGetModule Lib "vba6" (ByVal hProject As Long, ByVal iModule As Long, ByVal emod As Long, ByRef hModule As Long)
Private Declare Sub TipReleaseModule Lib "vba6" (ByVal hModule As Long)
Private Declare Sub TipGetModuleCount Lib "vba6" (ByVal hProject As Long, ByVal x As Long, ByRef lCount As Long)
Private Declare Sub TipGetModuleName Lib "vba6" (ByVal hModule As Long, ByRef pbstrName As Long)
Private Declare Sub EbGetModuleFlags Lib "vba6" (ByVal hModule As Long, ByRef F As EB_MODULE_FLAGS)
Private Declare Sub EbMemberBeginQuery Lib "vba6" (ByVal hModule As Long, ByVal iMember As Long)
Private Declare Sub EbMemberEndQuery Lib "vba6" ()
Private Declare Function EbMemberGetCount Lib "vba6" () As Long
Private Declare Function EbMemberGetName Lib "vba6" (ByVal i As Long) As Long
Private Declare Function EbMemberGetMemberkind Lib "vba6" (ByVal i As Long) As EB_MEMBER_KIND
Private Declare Function EbMemberGetMemid Lib "vba6" (ByVal i As Long, ByRef memid As Long) As Long
Private Declare Function EbMemberGetHostString Lib "vba6" (ByVal i As Long) As Long
Private Declare Function TipGetFunctionIdOfMod Lib "vba6" (ByVal hModule As Long, ByVal bstrName As Long, ByRef bstrId As Long) As Long
Private Declare Function TipGetLpfnOfFunctionId Lib "vba6" (ByVal hProject As Long, ByVal bstrId As Long, ByRef lpAddress As Long) As Long
Private Declare Function TipGetTypeLibOfHProject Lib "vba6" (ByVal hProject As Long, ByRef pTL As IUnknown) As Long
#ElseIf VB_VERSION = 5 Then
Private Declare Sub GetMem1 Lib "msvbvm50" (ByVal pdwFrom As Long, ByRef pdwTo As Byte)
Private Declare Sub EbGetExecutingProj Lib "vba5" (ByRef hProject As Long)
Private Declare Sub TipReleaseProject Lib "vba5" (ByVal hProject As Long)
Private Declare Sub TipGetModule Lib "vba5" (ByVal hProject As Long, ByVal iModule As Long, ByVal emod As Long, ByRef hModule As Long)
Private Declare Sub TipReleaseModule Lib "vba5" (ByVal hModule As Long)
Private Declare Sub TipGetModuleCount Lib "vba5" (ByVal hProject As Long, ByVal x As Long, ByRef lCount As Long)
Private Declare Sub TipGetModuleName Lib "vba5" (ByVal hModule As Long, ByRef pbstrName As Long)
Private Declare Sub EbGetModuleFlags Lib "vba5" (ByVal hModule As Long, ByRef F As EB_MODULE_FLAGS)
Private Declare Sub EbMemberBeginQuery Lib "vba5" (ByVal hModule As Long, ByVal iMember As Long)
Private Declare Sub EbMemberEndQuery Lib "vba5" ()
Private Declare Function EbMemberGetCount Lib "vba5" () As Long
Private Declare Function EbMemberGetName Lib "vba5" (ByVal i As Long) As Long
Private Declare Function EbMemberGetMemberkind Lib "vba5" (ByVal i As Long) As EB_MEMBER_KIND
Private Declare Function EbMemberGetFuncflags Lib "vba5" (ByVal i As Long) As Long
Private Declare Function EbMemberGetMemid Lib "vba5" (ByVal i As Long, ByRef memid As Long) As Long
Private Declare Function EbMemberGetHostString Lib "vba5" (ByVal i As Long) As Long
Private Declare Function TipGetFunctionIdOfMod Lib "vba5" (ByVal hModule As Long, ByVal bstrName As Long, ByRef bstrId As Long) As Long
Private Declare Function TipGetLpfnOfFunctionId Lib "vba5" (ByVal hProject As Long, ByVal bstrId As Long, ByRef lpAddress As Long) As Long
Private Declare Function TipGetTypeLibOfHProject Lib "vba5" (ByVal hProject As Long, ByRef pTL As IUnknown) As Long
#Else
    UNSUPPORTED VERSION
#End If

'
' Указатели.
'
Private Declare Function TlFindName Lib "*" (ByVal pTL As IUnknown, ByVal bstrName As Long, ByVal hash As Long, ByRef pTI As IUnknown, ByRef memid As Long, ByRef cFound As Integer) As Long
Private Declare Function TiAddressOfMember Lib "*" (ByVal pTI As IUnknown, ByVal memid As Long, ByVal invkind As Long, ByRef addr As Long) As Long

'
' Константы и типы, которые будут использованы и в скомпилированном файле,
' и при отладке под IDE.
'
Private Const Offs_EbImportModule               As Long = 4
Private Const Offs_EbImportFunc                 As Long = 8

Private Const EB_DELAYED_IMPORT_FLAG_BY_NAME    As Integer = &H400
Private Const EB_DELAYED_IMPORT_FLAG_BY_ORDINAL As Integer = &H200

Private Type EB_DELAYED_IMPORT_LOOKUP_ENTRY
    lpszModuleName              As Long
    lpszFunctionName            As Long
    sImportFlags                As Integer
    sFunctionOrdinal            As Integer
    lpImportAddressesEntry      As Long
End Type


'
' Константы и типы, которые будут использоваться в скомпилированном файле.
'
Private Const PAGE_EXECUTE_READWRITE            As Long = &H40

Private Const Offs_e_lfanew         As Long = &H3C
Private Const Offs_ImportTableRVA   As Long = &H80
Private Const Offs_ImportTableSize  As Long = &H84
Private Const Sz_ImportDescriptor   As Long = 20

Private Const Nm_LibraryKeyName     As String = "*"
Private Const Nm_DllExtension       As String = ".dll"
#If VB_VERSION = 6 Then
Private Const Nm_RtlModuleName      As String = "msvbvm60"
#ElseIf VB_VERSION = 5 Then
Private Const Nm_RtlModuleName      As String = "msvbvm50"
#End If

Private Type IMAGE_IMPORT_DESCRIPTOR
    rvaLookupTable              As Long
    TimeDateStampt              As Long
    ForwarderChain              As Long
    rvaModuleName               As Long
    rvaFirstThunk               As Long
End Type

Private Type SA_DESCRIPTOR_FOR_VECTOR
    cDimensions                 As Integer
    fFeatures                   As Integer
    cbItemSize                  As Long
    cLocks                      As Long
    pData                       As Long
    cL1Elements                 As Long
    iL1LBound                   As Long
End Type

'
' Константы и типы, которые будут использоваться только при отладке под IDE.
'
Private Const Offs_EbThunkDfcAddr As Long = &H11

Private Enum EB_MODULE_FLAGS
    EB_MODULE_STATIC = 1
    EB_MODULE_CLASS = 2
    EB_MODULE_STANDARD = 1024
End Enum

Private Enum EB_MEMBER_KIND
    EB_CONST
    EB_VAR
    EB_PROC
    EB_EVENT
End Enum

Private Enum ITL ' ITypeLib
    FindName = 11
End Enum

Private Enum ITI ' ITypeInfo
    AddressOfMember = 15
End Enum

Const EB_MODULE_MASK_ALL            As Long = &HFFFFFFFF
Const INVOKE_FUNC                   As Long = 1

Private m_DiscoveredPointers        As Collection
Private m_UndiscoveredPointers      As Collection

Private Function LetTrue(b As Boolean) As Boolean: b = True: LetTrue = True: End Function
Public Function L_(ByVal param As Long) As Long: L_ = param: End Function

Public Sub MagicPointersOnOff(ByVal bEngage As Boolean)
    Dim IDE_MODE As Boolean: Debug.Assert LetTrue(IDE_MODE)
    
    Static bEngaged         As Boolean
    Static lpIATEntry       As Long
    Static lIATOrigValue    As Long
    
    Dim i                   As Long
    Dim j                   As Long
    
    '
    ' У этой функции одна основная задача: установить и снять частичный (мо-
    ' дульно локальный) перехват функции DllFunctionCall.
    ' Как в скомпилированном файле, так и при отладке в IDE выполняется этот
    ' перехват, однако совершенно разными способами.
    '
    
    
    If bEngage Then
        If bEngaged Then Exit Sub
        
        If Not IDE_MODE Then
            '
            ' В скомпилированном файле выполнить перехват всех вызовов функции
            ' DllFunctionCall из текущего модуля легко: достаточно подменить
            ' значение ячейки таблицы импорта текущего модуля.
            '
            
            '
            ' Если это первый вызов, ячейка IAT ещё не найдена: ищем её.
            '
            
            If lpIATEntry = 0 Then
                
                
                Dim hMyModule           As Long
                Dim vaPEHeader          As Long
                Dim vaImportTable       As Long
                Dim ImportTable()       As IMAGE_IMPORT_DESCRIPTOR
                Dim saImportTable       As SA_DESCRIPTOR_FOR_VECTOR
                Dim iRtlDescriptor      As Long
           
                '
                ' Находим таблице дескрипторов импорта и инициализируем SA-дес-
                ' криптор.
                '
                
                hMyModule = App.hInstance
                GetMem4 hMyModule + Offs_e_lfanew, vaPEHeader
                vaPEHeader = vaPEHeader + hMyModule

                With saImportTable
                    .cDimensions = 1
                    .cbItemSize = Sz_ImportDescriptor
                    
                    GetMem4 vaPEHeader + Offs_ImportTableRVA, .pData
                    GetMem4 vaPEHeader + Offs_ImportTableSize, .cL1Elements
                    .pData = .pData + hMyModule
                End With

                PutMem4 AryPtr(ImportTable), VarPtr(saImportTable)
            
                iRtlDescriptor = -1
                
                '
                ' Обходим все дескрипторы импорта и находим дескрипторв, соот-
                ' ветствующий дескриптору msvbvmX0.dll
                '
                
                For i = 0 To saImportTable.cL1Elements - 1
                    If ImportTable(i).rvaLookupTable = 0 Then Exit For
                    Dim vaModuleName As Long
                    vaModuleName = ImportTable(i).rvaModuleName + hMyModule
                    If 0 = lstrcmpiA(vaModuleName, _
                                     Nm_RtlModuleName & Nm_DllExtension) Then
                        iRtlDescriptor = i
                    ElseIf 0 = lstrcmpiA(vaModuleName, _
                                         Nm_RtlModuleName) = 0 Then
                        iRtlDescriptor = i
                    End If
                Next i
            
                If iRtlDescriptor = -1 Then
                    '
                    ' Не нашли дескриптор: вероятно наш файл упаковали.
                    '
                    
                    PutMem4 AryPtr(ImportTable), 0&
                    Err.Raise 5, , _
                        "Runtime library have not been found in module imports!"
                    Exit Sub
                End If
                
                '
                ' Обходим таблицу Lookup-ов и ищем в ней Lookup, соответствующий
                ' импорту функции DllFunctionCall.
                '
        
                Dim lpLookupCursor  As Long
                Dim lpszProcName    As Long
                
                With ImportTable(iRtlDescriptor)
                    
                    lpLookupCursor = .rvaLookupTable + hMyModule
                
                    Do
                        GetMem4 lpLookupCursor, lpszProcName
                        If lpszProcName = 0& Then Exit Do
                        If lpszProcName > 0& Then
                            lpszProcName = lpszProcName + hMyModule + 2&
                            If 0 = lstrcmpiA(lpszProcName, _
                                             "DllFunctionCall") Then Exit Do
                        End If
                        lpLookupCursor = lpLookupCursor + 4&
                    Loop
                    
                    '
                    ' Не нашли Lookup: вероятно наш файл упакован.
                    '
                    
                    If lpszProcName = 0 Then
                        PutMem4 AryPtr(ImportTable), 0&
                        Err.Raise 5, , _
                                  "This module does not import DllFunctionCall."
                        Exit Sub
                    End If
                    
                    lpIATEntry = .rvaFirstThunk + (lpLookupCursor - .rvaLookupTable)
                End With
            
                Dim DFCDiscoveryBuffer(0 To 2) As Long
                Dim DFCDiscoveryLookup         As EB_DELAYED_IMPORT_LOOKUP_ENTRY
            
                '
                ' Делаем DFC-Discovery: вызов DFC, приводящий к тому, что адрес
                ' самой DFC попадает в VB-шную таблицу импорта именно сейчас.
                ' Но мы не можем вызвать DFC просто так (иначе будет Run-time
                ' error), поэтому заставляем DFC определить (и вернуть) адрес
                ' самой себя (так как её адрес нам потом понадобится).
                '
                
                With DFCDiscoveryLookup
                    .lpszModuleName = StrPtr(StrConv(Nm_RtlModuleName, vbFromUnicode))
                    .lpszFunctionName = StrPtr(StrConv("DllFunctionCall", vbFromUnicode))
                    .sImportFlags = EB_DELAYED_IMPORT_FLAG_BY_NAME
                    .lpImportAddressesEntry = VarPtr(DFCDiscoveryBuffer(0))
                End With
                
                lIATOrigValue = DllFunctionCall(DFCDiscoveryLookup)
            
                PutMem4 AryPtr(ImportTable), 0&
            End If

            '
            ' К этоу моменту адрес ячейки IAT известен: перезаписываем, и этим
            ' выполняем модуле-локальный перехват.
            '
            
            VirtualProtect lpIATEntry, 4, PAGE_EXECUTE_READWRITE, i
            PutMem4 lpIATEntry, AddressOf LocalResolver
            VirtualProtect lpIATEntry, 4, i, i
        
        Else ' if IDE_MODE
            '
            ' Под IDE надо сделать ровно то же самое: перехватить функцию
            ' DllFunctionCall модуле-локально.
            ' В скомпилированном файле это делалось легко: перезаписью ячейки
            ' IAT. А вот под IDE есть проблема: DFC не импортируется через им-
            ' порт, её адрес жестко вшит в динамически генерируемые куски кода.
            ' Так что единственный приемлемый способ её перехватить: найти эти
            ' динамически генерируемые кусочки кода (переходнички) и перезапи-
            ' сать адрес в них.
            '
            ' Нахождения таких переходников требует использования недокументи-
            ' рованных функций из VBAx.DLL.
            '
             
            Const S_OK                      As Long = 0
            Dim hres                        As Long ' HRESULT, на самом деле
            Dim hCurProject                 As Long
            Dim hCurModule                  As Long
            Dim CurProjectTL                As IUnknown
            Dim CurModuleTI                 As IUnknown
            Dim CurMemberAddress            As Long
            Dim CurMemberMemId              As Long
            Dim CurMemberName               As String
            Dim nModulesCount               As Long
            Dim nMembersCount               As Long
            Dim fMatchFlag                  As EB_MODULE_FLAGS
            Dim fModuleFlags                As EB_MODULE_FLAGS
            
            '
            ' Получаем хендл текщуго проекта, получаем число модулей в нём
            ' и проходимся по модулям. Делаем два прохода: сначала нужно пере-
            ' хватить DFC в обычных модулях (включая этот, в первую очередь).
            ' Затем, используя возможность вызова по указателю в этом модуле,
            ' делаем всё необходимое для её активации в модулях классов.
            ' Получаем также библиотеку типов проекта.
            '
            
            EbGetExecutingProj hCurProject
            TipGetTypeLibOfHProject hCurProject, CurProjectTL
            
            TipGetModuleCount hCurProject, EB_MODULE_MASK_ALL, nModulesCount
            
            '
            ' Перебираем обычные модули.
            '
            
            For fMatchFlag = EB_MODULE_STATIC To EB_MODULE_CLASS
                For i = 0 To nModulesCount - 1
                    TipGetModule hCurProject, i, EB_MODULE_MASK_ALL, hCurModule
                    
                    EbGetModuleFlags hCurModule, fModuleFlags
                    If fModuleFlags And fMatchFlag Then
                        
                        If fModuleFlags And EB_MODULE_CLASS Then
                            Dim bstrModuleName As Long
                            
                            TipGetModuleName hCurModule, bstrModuleName
                            
                            FuncPointer("TlFindName") = PointerFromMethIndex(CurProjectTL, _
                                                                             ITL.FindName)
                            
                            TlFindName CurProjectTL, bstrModuleName, 0, _
                                       CurModuleTI, 1, 1
                        End If
                        
                        EbMemberBeginQuery hCurModule, 0
                        
                        nMembersCount = EbMemberGetCount()
                        
                        '
                        ' Перебираем члены модуля, рассматриваем только
                        ' процедуры.
                        '
                        
                        For j = 0 To nMembersCount - 1
                            If EbMemberGetMemberkind(j) = EB_PROC Then
                                Dim bProcIsDeclare As Boolean
                                
                                If fMatchFlag = EB_MODULE_STATIC Then
                                    
                                    '
                                    ' Для обычных модулей TipGetFunctionIdOfMod
                                    ' работает, а для модулей классов по невыяс-
                                    ' ненной пока причине: нет. Поэтому для них
                                    ' используется другой метод.
                                    '
                                    
                                    Dim bstrIdOfMember As Long
                                    
                                    CurMemberName = AsciizToBSTR(EbMemberGetName(j))
                                    
                                    TipGetFunctionIdOfMod _
                                        hCurModule, _
                                        StrPtr(CurMemberName), _
                                        bstrIdOfMember

                                    TipGetLpfnOfFunctionId hCurProject, _
                                                           bstrIdOfMember, _
                                                           CurMemberAddress

                                    SysFreeString bstrIdOfMember
                                    
                                    bProcIsDeclare = CheckThunkSigAndLookup(CurMemberAddress)
                                    
                                ElseIf fMatchFlag = EB_MODULE_CLASS Then
                                    
                                    FuncPointer("TiAddressOfMember") = _
                                        PointerFromMethIndex(CurModuleTI, _
                                                             ITI.AddressOfMember)
                                                             
                                    TiAddressOfMember CurModuleTI, _
                                                      EbMemberGetMemid(j, 0), _
                                                      INVOKE_FUNC, _
                                                      CurMemberAddress
                                                      
                                    bProcIsDeclare = CheckThunkSigAndLookup(CurMemberAddress)
                                End If
                                
                                '
                                ' Делаем непосредственно перехват перезаписыва-
                                ' нием адреса функции DllFunctionCall.
                                '
                                
                                If bProcIsDeclare Then
                                    PutMem4 CurMemberAddress + Offs_EbThunkDfcAddr, _
                                            AddressOf LocalResolver
                                End If
                            End If
                        Next j
                        
                        If fModuleFlags And EB_MODULE_CLASS Then
                            Set CurModuleTI = Nothing
                        End If
                        
                        EbMemberEndQuery
                    End If
                    
                    TipReleaseModule hCurModule
                    
                Next i
                
                If fMatchFlag = EB_MODULE_STATIC Then
                    Set m_DiscoveredPointers = New Collection
                    Set m_UndiscoveredPointers = New Collection
                End If
            Next fMatchFlag
            
            Set CurProjectTL = Nothing
            
            TipReleaseProject hCurProject
            
            
        End If
        
        Set m_DiscoveredPointers = New Collection
        Set m_UndiscoveredPointers = New Collection
        
        bEngaged = True
        
    Else
        If Not bEngaged Then Exit Sub
        
        If Not IDE_MODE Then
            VirtualProtect lpIATEntry, 4, PAGE_EXECUTE_READWRITE, i
            PutMem4 lpIATEntry, lIATOrigValue
            VirtualProtect lpIATEntry, 4, i, i
        End If
        
        For i = 1 To m_DiscoveredPointers.Count
            PutMem4 m_DiscoveredPointers(i) + Offs_EbImportFunc, 0
        Next i
        
        Set m_DiscoveredPointers = Nothing
        Set m_UndiscoveredPointers = Nothing
        
        bEngaged = False
    End If
End Sub

Private Function AsciizToBSTR(ByVal lpsz As Long) As String
    Dim c As Long
    Dim b As Byte
    
    Do
        GetMem1 ByVal lpsz + c, b
        If b = 0 Then
            Exit Do
        Else
            c = c + 1
        End If
    Loop
        
    If c = 0 Then Exit Function

    AsciizToBSTR = Space(c)
    
    Do
        GetMem1 ByVal lpsz + c - 1, b
        Mid$(AsciizToBSTR, c, 1) = Chr$(b)
        c = c - 1
    Loop Until c = 0
End Function

Private Function CheckThunkSigAndLookup(ByVal lpThunkAddress As Long) As Boolean
    '
    ' Функция нужна только для работы в режиме отладки под IDE. В скомпилиро-
    ' ванный файл она НЕ ПОПАДЁТ.
    '
    
    Dim l As Long, le As EB_DELAYED_IMPORT_LOOKUP_ENTRY
    GetMem4 lpThunkAddress, l: If (l And &HFF) <> &HA1 Then Exit Function
    GetMem4 lpThunkAddress + 5, l: If l <> &H274C00B Then Exit Function
    GetMem4 lpThunkAddress + 21, l: If l <> &HE0FFD0FF Then Exit Function
    GetMem4 lpThunkAddress + 12, l
    RtlMoveMemory le, ByVal l, Len(le)
    CheckThunkSigAndLookup = 0 = lstrcmpiA(le.lpszModuleName, Nm_LibraryKeyName)
End Function

Private Function PointerFromMethIndex(ByVal pInterface As IUnknown, ByVal nIndex As Long) As Long
    '
    ' Функция нужна только для работы в режиме отладки под IDE. В скомпилиро-
    ' ванный файл она НЕ ПОПАДЁТ.
    '
    
    GetMem4 ObjPtr(pInterface), PointerFromMethIndex
    GetMem4 PointerFromMethIndex + nIndex * 4, PointerFromMethIndex
End Function

Private Function LocalResolver(ByRef pLookupEntry As EB_DELAYED_IMPORT_LOOKUP_ENTRY) As Long
    
    If Not pLookupEntry.sImportFlags And EB_DELAYED_IMPORT_FLAG_BY_ORDINAL Then
        If lstrcmpiA(pLookupEntry.lpszModuleName, Nm_LibraryKeyName) = 0 Then
            Dim lpAddress       As Long
            Dim sPointerName    As String
            
            sPointerName = AsciizToBSTR(pLookupEntry.lpszFunctionName)
            On Error Resume Next
            Err.Clear
            lpAddress = m_UndiscoveredPointers(sPointerName)
           
            If Err.Number <> 0 Then
                On Error GoTo 0
                On Error Resume Next
                lpAddress = m_DiscoveredPointers(sPointerName)
                If Err.Number <> 0 Then
                    On Error GoTo 0
                    Err.Raise 453, , "Specified pointer is not set."
                    Exit Function
                End If
            End If
           
            If lpAddress = 0 Then
                Err.Raise 453, , "Null pointer call."
                Exit Function
            End If

            m_UndiscoveredPointers.Remove sPointerName
            m_DiscoveredPointers.Add pLookupEntry.lpImportAddressesEntry, sPointerName
            
            PutMem4 pLookupEntry.lpImportAddressesEntry + Offs_EbImportFunc, lpAddress
            
            LocalResolver = lpAddress
            
            Exit Function
        End If
    End If
    
    LocalResolver = DllFunctionCall(pLookupEntry)

End Function

Public Property Get FuncPointer(ByVal sPointerName As String) As Long
    On Error GoTo try_discovered
        FuncPointer = m_UndiscoveredPointers(sPointerName)
        Exit Property
    On Error GoTo 0
    
try_discovered:
    On Error GoTo not_found_error_handler
        Dim lpImportAddressesEntry As Long
        lpImportAddressesEntry = m_DiscoveredPointers(sPointerName)
        GetMem4 lpImportAddressesEntry + Offs_EbImportFunc, FuncPointer
        Exit Property
    On Error GoTo 0
    
not_found_error_handler:
#If EMPTY_POINTER_CAUSES_ERROR Then
    Err.Raise 453, , "Specified pointer is not set."
#End If
End Property

Public Property Let FuncPointer(ByVal sPointerName As String, ByVal vNewValue As Long)
    On Error Resume Next
    
    Dim lpImportAddressesEntry As Long
    
    Err.Clear
    lpImportAddressesEntry = m_DiscoveredPointers(sPointerName)
    If Err.Number = 0 Then
        PutMem4 lpImportAddressesEntry + Offs_EbImportFunc, vNewValue
        Exit Property
    ElseIf Err.Number = 5 Then
        m_UndiscoveredPointers.Remove sPointerName
        m_UndiscoveredPointers.Add vNewValue, sPointerName
    Else
        MsgBox Err.Description
    End If
End Property


