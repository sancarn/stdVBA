Attribute VB_Name = "SampleFunctions"
Private Declare Sub fpEnumProc Lib "*" (sFileName As Any)

Public Function DoAdd(ByVal a As Long, ByVal b As Long) As Long
    DoAdd = a + b
End Function

Public Function DoSub(ByVal a As Long, ByVal b As Long) As Long
    DoSub = a - b
End Function




Public Sub OurProgram()
    Debug.Print "Сейчас мы перечислим файлы в текущей папке:"
    TheirEnumFiles AddressOf OurCallback
    Debug.Print "Для каждого файла был вызван по указателю OurCallback": Stop
End Sub

Public Sub OurCallback(ByRef sFileName As String)
    Debug.Print "file: '"; sFileName; "'"
End Sub


Public Sub TheirEnumFiles(ByVal pEnumProc As Long)
    Dim fn As String
    
    FuncPointer("fpEnumProc") = pEnumProc
    fn = Dir("*")
    Do
        Call fpEnumProc(ByVal VarPtr(fn))
        fn = Dir()
    Loop Until Len(fn) = 0
    
End Sub
