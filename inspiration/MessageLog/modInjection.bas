Attribute VB_Name = "modInjection"
Option Explicit

' Модуль для внедрения в чужой процесс и подмены оконной процедуры, с целью получить все сообщения пересылемые окну
' © Кривоус Анатолий Анатольевич (The trick), 2014


' ******                  ******* **                                         *            **
'  *    *                 *  *  *  *                        *                              *
'  *    *                    *     *                        *                              *
'  *    * *** ***            *     * **    *****           ****   *** **   ***     *****   *  **
'  *****   *   *             *     **  *  *     *           *       **  *    *    *     *  *  *
'  *    *  *   *             *     *   *  *******           *       *        *    *        * *
'  *    *   * *              *     *   *  *                 *       *        *    *        ***
'  *    *   * *              *     *   *  *     *           *  *    *        *    *     *  *  *
' ******     *              ***   *** ***  *****             **   *****    *****   *****  **   **
'            *
'          **

Private Type MessageInfo                    ' Эту структуру передаем в качестве параметра нашему окну
    Msg As Long
    wParam As Long
    lParam As Long
End Type
Private Type TrickThreadData
    SrcWnd As Long                          ' Хендл сабклассируемого окна
    DesthWnd As Long                        ' Хендл окна frmSpy
    EventHandle As Long                     ' Хендл события, отвечающего за завершение потока
    AddrWindowProc As Long                  ' Адрес функции WindowProc в чужом процессе
    AddrStructure As Long                   ' Адрес этой структуры
    Msg As MessageInfo                      ' Для передачи указателя COPYDATASTRUCT
End Type
Private Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function VirtualAllocEx Lib "kernel32.dll" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFreeEx Lib "kernel32.dll" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As Long, lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60" (src As Any, dst As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (lpEventAttributes As Any, ByVal bManualReset As Long, ByVal bInitialState As Long, lpName As Any) As Long
Private Declare Function PulseEvent Lib "kernel32" (ByVal hEvent As Long) As Long
Private Declare Function DuplicateHandle Lib "kernel32" (ByVal hSourceProcessHandle As Long, ByVal hSourceHandle As Long, ByVal hTargetProcessHandle As Long, lpTargetHandle As Long, ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwOptions As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_COPYDATA = &H4A
Private Const GWL_WNDPROC = (-4)
Private Const DUPLICATE_SAME_ACCESS = &H2
Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Const MEM_COMMIT = &H1000&
Private Const MEM_RESERVE = &H2000&
Private Const MEM_RELEASE = &H8000&
Private Const PAGE_EXECUTE_READWRITE = &H40&
Private Const INFINITE = -1&

Private Const Prop As String = "pInject"                    ' 7 символов + \0, итого 8 байт, вполне помещается в переменную типа Currency
Private Const PropCur As Currency = 3276038452689.5472@     ' Строка Prop в виде Currecy числа

Public hProcess As Long                                     ' Хендл процесса, в который внедряемся
Public hThread As Long                                      ' Хендл потока, который мы создадим в чужом процессе
Public TID As Long                                          ' Идентификатор этого потока
Public lpProc As Long                                       ' Адрес функции InjectionProc
Public Size As Long                                         ' Размер данных и кода, внедряемого в процесс
Public hEvent As Long                                       ' Описатель события в нашем процессе

Dim lpPrevWndProc As Long                                   ' Адрес оконной процедуры frmSpy (изначальный)

' Функция внедряет код в чужой процесс
Public Function Hook(hwnd As Long) As Boolean
    Dim Buf() As Byte, ret As Long, PID As Long, DupHandle As Long, nearWndProc As Long, _
        FuncOf() As Long, FuncAddr() As Long, hMod As Long, lpFunc As Long, i As Long, lpData As Long
        
    If hProcess Then Clear                   ' Если перехват был, то убираем
    GetWindowThreadProcessId hwnd, PID
    
    ' Инициализация словаря
    If modListView.Dic Is Nothing Then modListView.DicInit
    
    If PID Then hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, PID) Else Exit Function

    ' Создаем событие для управления потоком
    hEvent = CreateEvent(ByVal 0, 1, 0, ByVal 0)

    If hEvent = 0 Then Clear: Exit Function
    ' Создаем дубликат описателя события для процесса
    If DuplicateHandle(GetCurrentProcess(), hEvent, hProcess, DupHandle, 0, False, DUPLICATE_SAME_ACCESS) = 0 Then Clear: Exit Function

    ' Определяем размер для внедренного кода
    lpData = AddrOf(AddressOf AddrOf) - AddrOf(AddressOf InjectionProc)
    ' Определяем относительное смещение функции WindowProc от данных
    nearWndProc = AddrOf(AddressOf AddrOf) - AddrOf(AddressOf WindowProc)
    ' Определяем размер данных и кода
    Size = lpData + 32

    ' Выделяем память в чужом процессе
    lpProc = VirtualAllocEx(hProcess, ByVal 0, Size, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If lpProc = 0 Then MsgBox "Error allocate memory", vbCritical: Clear: Exit Function

    ' Определяем смещения для псевдофункций API относительно начала данных
    ReDim FuncOf(9)
    FuncOf(0) = AddrOf(AddressOf myCopyMemory) - AddrOf(AddressOf InjectionProc)
    FuncOf(1) = AddrOf(AddressOf myCopyMemory2) - AddrOf(AddressOf InjectionProc)
    FuncOf(2) = AddrOf(AddressOf myCloseHandle) - AddrOf(AddressOf InjectionProc)
    FuncOf(3) = AddrOf(AddressOf myWaitForSingleObject) - AddrOf(AddressOf InjectionProc)
    FuncOf(4) = AddrOf(AddressOf mySetProp) - AddrOf(AddressOf InjectionProc)
    FuncOf(5) = AddrOf(AddressOf myGetProp) - AddrOf(AddressOf InjectionProc)
    FuncOf(6) = AddrOf(AddressOf myRemoveProp) - AddrOf(AddressOf InjectionProc)
    FuncOf(7) = AddrOf(AddressOf mySetWindowLong) - AddrOf(AddressOf InjectionProc)
    FuncOf(8) = AddrOf(AddressOf mySendMessage) - AddrOf(AddressOf InjectionProc)
    FuncOf(9) = AddrOf(AddressOf myCallWindowProc) - AddrOf(AddressOf InjectionProc)

    ' Определяем адреса API функций, для системных библиотек их образы спроецированы по одному и томуже адресу что и у нас
    ReDim FuncAddr(9)
    hMod = GetModuleHandle("kernel32")
    FuncAddr(0) = GetProcAddress(hMod, "RtlMoveMemory")
    FuncAddr(1) = FuncAddr(0)
    FuncAddr(2) = GetProcAddress(hMod, "CloseHandle")
    FuncAddr(3) = GetProcAddress(hMod, "WaitForSingleObject")
    hMod = GetModuleHandle("user32")
    FuncAddr(4) = GetProcAddress(hMod, "SetPropA")
    FuncAddr(5) = GetProcAddress(hMod, "GetPropA")
    FuncAddr(6) = GetProcAddress(hMod, "RemovePropA")
    FuncAddr(7) = GetProcAddress(hMod, "SetWindowLongA")
    FuncAddr(8) = GetProcAddress(hMod, "SendMessageA")
    FuncAddr(9) = GetProcAddress(hMod, "CallWindowProcA")

    ' Копируем код
    ReDim Buf(Size - 1)
    CopyMemory Buf(0), ByVal AddrOf(AddressOf InjectionProc), lpData

    ' Модифицируем код для вызова API вместо наших пустышек
    For i = 0 To UBound(FuncOf)
        Buf(FuncOf(i)) = &HE9                                                   ' JMP
        GetMem4 (FuncAddr(i) - FuncOf(i) - lpProc) - 5, Buf(FuncOf(i) + 1)      ' near (относительный прыжок на API функцию)
    Next

    ' Копируем данные
    GetMem4 hwnd, Buf(lpData)                                                   ' Хендл сабклассируемого окна
    GetMem4 frmSpy.hwnd, Buf(lpData + 4)                                        ' Хендл окна-приемника
    GetMem4 DupHandle, Buf(lpData + 8)                                          ' Хендл события
    GetMem4 lpProc + lpData - nearWndProc, Buf(lpData + 12)                     ' Адрес WindowProc в чужом процессе
    GetMem4 lpProc + lpData, Buf(lpData + 16)                                   ' Адрес этой структуры в чужом процессе
    
    ' Делаем инъекцию
    If WriteProcessMemory(hProcess, lpProc, Buf(0), Size, ret) Then
        If ret <> Size Then MsgBox "Error write process", vbCritical: Clear: Exit Function
        ' Запускаем код инъекции
        hThread = CreateRemoteThread(hProcess, ByVal 0, 0, lpProc, ByVal lpProc + Size - 32, 0, TID)
        If hThread = 0 Then MsgBox "Error create thread", vbCritical: Clear: Exit Function
    End If
    
    lpPrevWndProc = SetWindowLong(frmSpy.hwnd, GWL_WNDPROC, AddressOf SpyWindowProc)     ' Сабклассим наше окно
    
    Hook = True
End Function

' Удалить инъекцию
Public Sub Clear()
    If lpPrevWndProc Then
        SetWindowLong frmSpy.hwnd, GWL_WNDPROC, lpPrevWndProc       ' Убираем сабклассинг
        lpPrevWndProc = 0
    End If
    If hThread Then
        PulseEvent hEvent                                           ' Запускаем завершение потока
        WaitForSingleObject hThread, INFINITE                       ' Ждем завершения потока (замораживаемся)
        CloseHandle hThread                                         ' Закрываем описатель потока
        hThread = 0
    End If
    If lpProc Then
        Call VirtualFreeEx(hProcess, ByVal lpProc, 0, MEM_RELEASE)  ' Освобождаем выделенную память
    End If
    If hProcess Then
        CloseHandle hProcess                                        ' Закрываем описатель процесса
        hProcess = 0
    End If
    If hEvent Then
        CloseHandle hEvent                                          ' Закрываем описатель события (объект тоже удалится)
        hEvent = 0
    End If
End Sub
' Оконная процедура для отслеживания сообщений из нашего процесса
Private Function SpyWindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim CDS As COPYDATASTRUCT, Info As MessageInfo
    
    If Msg = WM_COPYDATA Then
        ' Получили сообщение из того процесса!!!
        CopyMemory CDS, ByVal lParam, Len(CDS)
        CopyMemory Info, ByVal CDS.lpData, CDS.cbData
        ItemAdd modListView.GetMessageName(Info.Msg), Info.wParam, Info.lParam
    End If
    
    ' Обрабатываем как и раньше
    SpyWindowProc = CallWindowProc(lpPrevWndProc, hwnd, Msg, wParam, ByVal lParam)
End Function

' Данный код выполняется в АП чужого процесса, поэтому он не имеет понятия ни о каких глобальных или локальных переменных
' уровня этого модуля, единственная область памяти с которой он может работать передаеться ему указателем на структуру
' TrickThreadData, который в последствии сохраняется в свойстве окна 'pInject'. Вызов наших функций, ведет к перенапрвлению
' к соответствующим API функциям. Здесь выполняется код, который вообще не использует рантайм. Для использования функций
' рантайма (сейчас о функциях которые не требуют инициализацию контекста потока), нужно его предварительно загрузить, через
' LoadLibrary() и получить адреса функций через GetProcAddress(). Все символьные имена и переменные, нужно хранить в
' в выделенной для этого предварительно памяти. Так что обращение к любой глобальной переменной или константе
' (пример s$="VB6 best language") может вызвать ошибку доступа

' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
' Процедура, выполняемая в чужом процессе, передаем ей указатель на данные
Private Sub InjectionProc(Dat As TrickThreadData)
    Dim lpOldProc As Long
    ' Мы в чужом процессе ))
    mySetProp Dat.SrcWnd, PropCur, Dat.AddrStructure                         ' Устанавливаем окну свойство с указателем на данные
    lpOldProc = mySetWindowLong(Dat.SrcWnd, GWL_WNDPROC, Dat.AddrWindowProc) ' Устанавливаем окну новый оконный обработчик
    ' Вместо нового адреса процедуры пишем старое
    Dat.AddrWindowProc = lpOldProc
    ' Замораживаем поток
    myWaitForSingleObject Dat.EventHandle, INFINITE
    ' Поток разморожен, значит надо возвращать все на место
    mySetWindowLong Dat.SrcWnd, GWL_WNDPROC, Dat.AddrWindowProc
    myRemoveProp Dat.SrcWnd, PropCur
    ' Закрываем описатель события
    myCloseHandle Dat.EventHandle
    ' Все поток закончен, теперь Clear разморозится и очистит занимаемую память
End Sub

' Прцедуры вызова соответствующих API c помощью сплайсинга
Private Function myCopyMemory(dst As TrickThreadData, ByVal src As Long, ByVal Length As Long) As Long
    myCopyMemory = -1
End Function
Private Function myCopyMemory2(ByVal dst As Long, src As TrickThreadData, ByVal Length As Long) As Long
    myCopyMemory2 = -2
End Function
Private Function mySetProp(ByVal hwnd As Long, ByRef Name As Currency, ByVal Value As Long) As Long
    mySetProp = -3
End Function
Private Function myGetProp(ByVal hwnd As Long, ByRef Name As Currency) As Long
    myGetProp = -4
End Function
Private Function myRemoveProp(ByVal hwnd As Long, ByRef Name As Currency) As Long
    myRemoveProp = -5
End Function
Private Function mySetWindowLong(ByVal hwnd As Long, ByVal Index As Long, ByVal Data As Long) As Long
    mySetWindowLong = -6
End Function
Private Function myWaitForSingleObject(ByVal hEvent As Long, ByVal Millisecond As Long) As Long
    myWaitForSingleObject = -7
End Function
Private Function mySendMessage(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As COPYDATASTRUCT) As Long
    mySendMessage = -8
End Function
Private Function myCallWindowProc(ByVal addr As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    myCallWindowProc = -9
End Function
Private Function myCloseHandle(ByVal Handle As Long) As Long
    myCloseHandle = -10
End Function
' Оконная функция, которая будет работать в чужом процессе
Private Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim lpDat As Long, Dat As TrickThreadData, CDS As COPYDATASTRUCT
    
    lpDat = myGetProp(hwnd, PropCur)

    myCopyMemory Dat, lpDat, Len(Dat)                   ' Копируем параметры
    
    ' Устанавливаем параметры сообщения
    Dat.Msg.Msg = uMsg
    Dat.Msg.wParam = wParam
    Dat.Msg.lParam = lParam
    
    myCopyMemory2 lpDat, Dat, Len(Dat)                  ' Копируем параметры обратно
    
    CDS.cbData = Len(Dat.Msg)
    CDS.lpData = lpDat + 20                             ' Смещение структуры MessageInfo, относительно данных
    
    ' Отправляем нашему окну уведомление
    mySendMessage Dat.DesthWnd, WM_COPYDATA, hwnd, CDS
    
    ' Вызываем процедуру по умолчанию
    WindowProc = myCallWindowProc(Dat.AddrWindowProc, hwnd, uMsg, wParam, lParam)
End Function
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
' Эта функция также служит маркером конца функции и в процесс не копируется
Private Function AddrOf(Value As Long) As Long
    AddrOf = Value
End Function
