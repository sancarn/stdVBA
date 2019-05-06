Attribute VB_Name = "modNativeInfo"
Option Explicit

' Декларации

' Информация об элементе экспорта
Public Type ExportInfo
    EntryPoint      As Long     ' Точка входа
    Forwarder       As Boolean  ' Если перенаправление
    Ordinal         As Long     ' Ординал
    Name            As String   ' Имя
End Type
' Информация об элементе импорта
Public Type ImportInfo
    Name            As String   ' Имя библиотеки
    Count           As Long     ' Количество импортируемых функций из этой библиотеки
    Func()          As String   ' Список функций
End Type
' Информация о DLL
Public Type NativeInfo
    ImportCount     As Long     ' Количество элементов импорта
    ExportCount     As Long     ' Количество элементов экспорта
    DelImpCount     As Long     ' Количество элементов отложенного импорта
    Import()        As ImportInfo
    Export()        As ExportInfo
    DelayImport()   As ImportInfo
End Type
Public Type tagSAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

