VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Console 
   Caption         =   "Console"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9765.001
   OleObjectBlob   =   "Console.frx":0000
   ShowModal       =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "Console"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Private and Public Variables
    
    Private       CurrentLineIndex    As Long              ' Index of the CurrentLine
    Private Const Recognizer          As String = "\>>>"   ' Used to recognize when CurrentLine should start
    Private       PasteStarter        As Boolean           ' Determines if starter should be printed or not
    Private       UserInput           As Variant           ' Value the user put in
    Private       LastError           As Variant           ' Used for more information about last Error
    Private       MultilineTemp       As Boolean           ' Used to stop Printstarter for the "MULTILINE" Special keyword and stops handleenter to run code until there is no "_" character at the end

    Private       ConsoleVariables()  As Variant           ' Defined as (1, n), where 0,n is the variable name and 1,n is its value. Used to save userinput into a variable with name

    ' Check if the console awaits pre-declared answer, user input or just logging
    Private WorkMode As Long
    Private Enum WorkModeEnum
        Logging = 0
        UserInputt = 1
        PreDeclaredAnswer = 2
        UserLog = 3
        MultilineMode = 4
    End Enum

    


    ' Dependenant on Microsoft Visual Studio Extensebility 5.3
    Private Const Intellisense_Active As Boolean = True
    Private       VariableArray()     As Variant          ' Big Array storing all Variables, Subs and Functions of all classes, modules and forms
    Private       Intellisense_Index  As Long
    Private Const MaxArg              As Long = 33        ' 

    ' Dependant on sancarn´s stdLambda class
    Private Const stdLambda_Active    As Boolean = True

    'Intellisense Colors
    Private in_Basic       As Long
    Private in_Procedure   As Long
    Private in_Operator    As Long 'smooooooth operatooooor
    Private in_Datatype    As Long
    Private in_Value       As Long
    Private in_String      As Long
    Private in_Statement   As Long
    Private in_Keyword     As Long
    Private in_Parantheses As Long
    Private in_Variable    As Long
'







' Public Console Functions

    Public Property Get GetName(Index As Long) As Variant
        GetName = ConsoleVariables(0, Index)
    End Property

    Public Property Get GetValue(Optional Name As Variant = Empty, Optional Index As Variant = Empty) As Variant
        Dim i As Long
        If Index <> Empty Then
            GetValue = ConsoleVariables(1, Index)
            Exit Function
        Else
            For i = 0 To Ubound(ConsoleVariables, 2)
                If ConsoleVariables(0, i) = Name Then
                    GetValue = ConsoleVariables(1, i)
                    Exit Function
                End If
            Next
        End If
        GetValue = Empty
    End Property


    ' Check for UserInput
    ' Answers needs to be of same dimension as AllowedValues
    Public Function GetUserInput(Message As Variant, Optional InputType As String = "VARIANT") As Variant

        PrintConsole Message
        WorkMode = WorkModeEnum.UserInputt
        PasteStarter = False
        Do While WorkMode = WorkModeEnum.UserInputt
            DoEvents
            If UserInput <> "" Then
                UserInput = Replace(UserInput, Message, "")
                GetUserInput = UserInput
                WorkMode = WorkModeEnum.Logging
                UserInput = ""
            End If
        Loop
        PasteStarter = True

    End Function

    ' Check for UserInput
        ' Answers needs to be of same dimension as AllowedValues
    Public Function CheckPredeclaredAnswer(Message As Variant, AllowedValues As Variant, Optional Answers As Variant = Empty) As Variant

        Dim i As Variant
        Dim Found As Boolean
        Dim Index As Long

        Message = Message & "("
        For Each i In AllowedValues
            Message = Message & Cstr(i) & "|"
        Next i
        Message = Message & ") "
        PrintConsole Message

        WorkMode = WorkModeEnum.PreDeclaredAnswer
        PasteStarter = False
        Do While WorkMode = WorkModeEnum.PreDeclaredAnswer
            Index = 0
            DoEvents
            If UserInput <> "" Then
                UserInput = Replace(UserInput, Message, "")
                For Each i In AllowedValues
                    If Cstr(i) = UserInput Then
                        CheckPredeclaredAnswer = i
                        Found = True
                        WorkMode = WorkModeEnum.Logging
                        Exit For
                    End If
                    Index = Index + 1
                Next i
                If Found <> True Then
                    PrintEnter "Value not Valid"
                    PrintConsole Message
                End If
                UserInput = ""
            End If
        Loop
        PasteStarter = True
        PrintEnter Answers(Index)
        PrintConsole PrintStarter

    End Function

    Public Function PrintStarter() As Variant
        PrintStarter = ThisWorkbook.Path & Recognizer
    End Function

    Public Sub PrintEnter(Text As Variant, Optional Color As Variant)
        PrintConsole Text & vbcrlf, Color
    End Sub

    Public Sub PrintConsole(Text As Variant, Optional Color As Variant)
        
        Dim i As Long
        If IsMissing(Color) Then Color = in_Basic
        If ISArray(Color) Then
            If Ubound(Color) + 1 < Len(Text) Then
                LastError = 4
                HandleLastError
                Exit Sub
            End If
            ConsoleText.SelLength = 0
            For i = 1 To Len(Text)
                ConsoleText.SelStart = Len(ConsoleText.Text)
                ConsoleText.SelColor = Color(i - 1)
                ConsoleText.SelText = Mid(Text, i, 1)
            Next i
        Else
            ConsoleText.SelStart = Len(ConsoleText.Text)
            ConsoleText.SelLength = 0
            ConsoleText.SelColor = Color
            ConsoleText.SelText = Text
        End If
        SetUpNewLine

    End Sub
'

' Private Console Functions

    Private Sub UserForm_Initialize()
        AssignColor
        PasteStarter = True
        ConsoleText.Text = GetStartText
        ConsoleText.SelStart = 0
        ConsoleText.SelLength = Len(ConsoleText.Text)
        ConsoleText.SelColor = in_Basic
        CurrentLineIndex = UBound(Split(ConsoleText.Text, vbCrLf))
        ScrollHeight = 5000
        ScrollWidth = 3000
        ReDim ConsoleVariables(1, 0)
        ' Dependenant on Microsoft Visual Studio Extensebility 5.3
        If Intellisense_Active = True Then
            GetAllProcedures
        End If
    End Sub

    ' Just doing &H00FFFFFF will (for some reason) become a negative number, so to secure a positive number this is done
    Private Sub AssignColor()
        in_Basic        = &H10FFFFFF - &H10000000
        in_Procedure    = &H1000FFFF - &H10000000
        in_Operator     = &H100000FF - &H10000000
        in_Datatype     = &H1000AA00 - &H10000000
        in_Value        = &H1000FF00 - &H10000000
        in_String       = &H1000AAFF - &H10000000
        in_Statement    = &H10FF00FF - &H10000000
        in_Keyword      = &H10FF0000 - &H10000000
        in_Parantheses  = &H1000AAAA - &H10000000
        in_Variable     = &H10FFFF00 - &H10000000
    End Sub

    Private Sub SetValue(Name As Variant, Value As Variant, Optional Index As Long = Empty)
        Dim i As Long
        If Index <> Empty Then
            ConsoleVariables(1, Index) = Value
            Exit Sub
        End If
        For i = 0 To Ubound(ConsoleVariables, 2)
            If ConsoleVariables(0, i) = Name Or ConsoleVariables(0, i) = Empty Then
                ConsoleVariables(0, i) = Name
                ConsoleVariables(1, i) = Value
                Exit Sub
            End If
        Next
        ReDim Preserve ConsoleVariables(1, Ubound(ConsoleVariables, 2) + 1)
        ConsoleVariables(0, Ubound(ConsoleVariables, 2)) = Name
        ConsoleVariables(1, Ubound(ConsoleVariables, 2)) = Value

    End Sub

    Private Sub ConsoleText_KeyUp(pKey As Long, ByVal ShiftKey As Integer)
        
        Dim Temp As Variant
        Temp = Split(ConsoleText.Text, vbCrLf)
        Select Case pKey
            Case vbKeyReturn
                HandleEnter
                SetPositions
            Case vbKeyUp
                If CurrentLineIndex > 1 Then
                    CurrentLineIndex = CurrentLineIndex - 1
                    ConsoleText.SelStart = Len(ConsoleText.Text) - Len(Temp(Ubound(Temp))) - 5
                    ConsoleText.SelLength = Len(ConsoleText.Text)
                    ConsoleText.SelText = PrintStarter & GetLine(ConsoleText.Text, CurrentLineIndex)
                End If
            Case vbKeyDown
                If CurrentLineIndex < UBound(Temp) Then
                    CurrentLineIndex = CurrentLineIndex + 1
                    ConsoleText.SelStart = Len(ConsoleText.Text) - Len(Temp(Ubound(Temp)))
                    ConsoleText.SelLength = Len(ConsoleText.Text)
                    ConsoleText.SelText = PrintStarter & GetLine(ConsoleText.Text, CurrentLineIndex) - 5
                End If
            Case Else
                HandleOtherKeys pKey, ShiftKey
        End Select

    End Sub

    ' Module Code
    Private Function GetLine(Text As String, Index As Long) As String
        Dim Lines() As String
        Dim SearchString As String
        Dim ReplaceString As String
        Dim i As Variant
        Lines = Split(Text, vbCrLf)
        If Index > 0 And Index <= UBound(Lines) + 1 Then
            SearchString = Lines(Index)
            If InStr(1, SearchString, Recognizer) = 0 Then
                ReplaceString = ""
            Else
                ReplaceString = Mid(SearchString, 1, InStr(1, SearchString, Recognizer) - 1 + Len(Recognizer))
            End If
            GetLine = Replace(SearchString, ReplaceString, "")
        Else
            GetLine = "Line number out of range"
        End If
    End Function

    Private Function GetWord(Text As String, Optional Index As Long = Empty) As String
        Dim Words() As String
        Words = Split(Text, " ")
        If Index = Empty Then Index = Ubound(Words)
        If Ubound(Words) > -1 Then GetWord = Words(Index)
    End Function

    Private Sub SetUpNewLine()
        Dim Temp As Variant
        Temp = Split(ConsoleText.Text, vbCrLf)
        CurrentLineIndex = Ubound(Temp)
    End Sub

    Private Sub HandleEnter()

        Dim i As Long
        Dim Line As String
        SetUpNewLine
        Line = GetLine(ConsoleText.Text, CurrentLineIndex - 1)
        MulitlineEnd:
        Select Case WorkMode
            Case WorkModeEnum.Logging
                If InStr(1, Line, "==") <> 0 Then
                    HandleConsoleVariable Line
                Else
                    PrintEnter HandleCode(SplitString(Line, "; "))
                End If
            Case WorkModeEnum.UserInputt, WorkModeEnum.PredeclaredAnswer
                UserInput = Replace(Line, vbCrLf, "")
            Case WorkModeEnum.UserLog

            Case WorkModeEnum.MultilineMode
                If Mid(Line, Len(Line), 1) <> "_" Then
                    Dim Temp(2) As String '0 is final string, 1 fusion of final string with current string, 2 is current string
                    Dim TempCount As Long
                    Temp(0) = Line
                    TempCount = CurrentLineIndex - 1
                    Temp(2) = GetLine(ConsoleText.Text, TempCount - 1) ' needed for initialization
                    If Len(Temp(2)) = 0 Then 
                        Do Until Mid(Temp(2), Len(Temp(2)), 1) <> "_"
                            Temp(1) = Temp(2) & Temp(0)
                            Temp(0) = Temp(1)
                            Temp(1) = ""
                            TempCount = TempCount - 1
                            Temp(2) = GetLine(ConsoleText.Text, TempCount - 1)
                            If Len(Temp(2)) = 0 Then Exit Do
                        Loop
                    End If
                    Temp(0) = Replace(Temp(0), "_", "")
                    Line = Temp(0)
                    WorkMode = WorkModeEnum.Logging
                    GoTo MulitlineEnd
                End If
        End Select
        If PasteStarter = True And Workmode <> WorkModeEnum.MultilineMode Then PrintConsole PrintStarter

    End Sub

    Private Function HandleCode(Arguments() As Variant) As Variant

        Dim Temp As Variant
        Temp = HandleSpecial(Arguments)
        Select Case Temp
            Case 1: HandleCode = "Success": LastError = 1: Exit Function
            Case 2: HandleCode = "": Exit Function
            Case 3: 
                If IsDate(Arguments(0)) Then
                    HandleCode = CDate(Arguments(0))
                ElseIf IsNumeric(Arguments(0)) Then
                    HandleCode = CDbl(Arguments(0))
                Else
                    Dim i As Long
                    Dim Found As Boolean
                    For i = 0 To Ubound(ConsoleVariables, 2)
                        If Arguments(0) = ConsoleVariables(0, i) Then
                            HandleCode = Arguments(0)
                            Found = True
                        End If
                    Next
                    If Found <> True Then HandleCode = RunApplication(Arguments)
                End If
        End Select
        If HandleCode = Empty Then
            HandleCode = "Success"
            LastError = 1
        End If

    End Function

    Private Function RunApplication(Arguments() As Variant) As Variant

        On Error GoTo Error
        Select Case UBound(Arguments)
            Case 00:   RunApplication = Application.Run(Arguments(0))
            Case 01:   RunApplication = Application.Run(Arguments(0), Arguments(1))
            Case 02:   RunApplication = Application.Run(Arguments(0), Arguments(1), Arguments(2))
            Case 03:   RunApplication = Application.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3))
            Case 04:   RunApplication = Application.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4))
            Case 05:   RunApplication = Application.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5))
            Case 06:   RunApplication = Application.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6))
            Case 07:   RunApplication = Application.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7))
            Case 08:   RunApplication = Application.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8))
            Case 09:   RunApplication = Application.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9))
            Case 10:   RunApplication = Application.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10))
            Case 11:   RunApplication = Application.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11))
            Case 12:   RunApplication = Application.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12))
            Case 13:   RunApplication = Application.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13))
            Case 14:   RunApplication = Application.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14))
            Case 15:   RunApplication = Application.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15))
            Case 16:   RunApplication = Application.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16))
            Case 17:   RunApplication = Application.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17))
            Case 18:   RunApplication = Application.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18))
            Case 19:   RunApplication = Application.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19))
            Case 20:   RunApplication = Application.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20))
            Case 21:   RunApplication = Application.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21))
            Case 22:   RunApplication = Application.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22))
            Case 23:   RunApplication = Application.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23))
            Case 24:   RunApplication = Application.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23), Arguments(24))
            Case 25:   RunApplication = Application.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23), Arguments(24), Arguments(25))
            Case 26:   RunApplication = Application.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23), Arguments(24), Arguments(25), Arguments(26))
            Case 27:   RunApplication = Application.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23), Arguments(24), Arguments(25), Arguments(26), Arguments(27))
            Case 28:   RunApplication = Application.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23), Arguments(24), Arguments(25), Arguments(26), Arguments(27), Arguments(28))
            Case 29:   RunApplication = Application.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23), Arguments(24), Arguments(25), Arguments(26), Arguments(27), Arguments(28), Arguments(29))
            Case 30:   RunApplication = Application.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23), Arguments(24), Arguments(25), Arguments(26), Arguments(27), Arguments(28), Arguments(29), Arguments(30))
            Case Else: RunApplication = "Too many Arguments": LastError = 3
        End Select
        If IsError(RunApplication) Then
            GoTo Error:
        Else
            Exit Function
        End If
        Error:
        RunApplication = "Could not run Procedure. Procedure might not exist"
        LastError = 2

    End Function

    Private Function HandleSpecial(Arguments() As Variant) As Variant
        Select Case True
            Case UCase(CStr(Arguments(0))) Like "HELP"      : HandleSpecial = 1: HandleHelp
            Case UCase(CStr(Arguments(0))) Like "CLEAR"     : HandleSpecial = 2: HandleClear
            Case UCase(CStr(Arguments(0))) Like "MULTILINE" : HandleSpecial = 2: Workmode = WorkModeEnum.MultilineMode
            Case UCase(CStr(Arguments(0))) Like "INFO"      : HandleSpecial = 1: HandleLastError
            Case UCase(CStr(Arguments(0))) Like "[?]*"      : HandleSpecial = 4
            Case Else                                       : HandleSpecial = 3
        End Select
    End Function

    Private Sub HandleConsoleVariable(Line As Variant)

        Dim AssignOperator As Long    : AssignOperator = InStr(1, Line, "==")
        Dim Name           As Variant : Name           = Mid(Line, 1, AssignOperator - 1)
        Dim RightSide      As String  : RightSide      = Mid(Line, AssignOperator + 2)
        Dim Value          As Variant : Value          = HandleCode(SplitString(RightSide, "; "))
        Dim i              As Long
        Dim VariableFound  As Boolean

        If Value <> "Could not run Procedure. Procedure might not exist" Or IsNumeric(Value) = False Then
            For i = 0 To Ubound(ConsoleVariables, 2)
                If Value = ConsoleVariables(0, i) Then
                    Value = ConsoleVariables(1, i)
                    VariableFound = True
                End If
            Next
            If VariableFound <> True Then Value = RightSide
        End If
        SetValue Name, Value

    End Sub

    Private Function SplitString(Text As Variant, SplitText As Variant) As Variant()
        Dim Temp() As String
        Dim ReturnArray() As Variant
        Dim i As Long

        If InStr(1, CStr(Text), CStr(SplitText)) <> 0 Then
            Temp = Split(CStr(Text), CStr(SplitText))
            ReDim ReturnArray(Ubound(Temp))
            For i = 0 To Ubound(ReturnArray)
                ReturnArray(i) = CStr(Temp(i))
            Next i
        Else
            ReDim ReturnArray(0)
            ReturnArray(0) = Text
        End If
        SplitString = ReturnArray
    End Function

    Private Sub HandleLastError()
        Dim Message As String
        Dim Color As Long

        Color = &H100000FF - &H10000000
        Select Case LastError
        Case Empty:
            Message = "No previous Error detected"
        Case 1
            Message = "The last run line was executed without problems"
        Case 2
            Message = "The last run line couldnt be executed. Some Problems could be:"         & vbcrlf & _
                      "    1. Line wasnt written correctly"                                    & vbcrlf & _
                      "    2. Code doesnt exist"                                               & vbcrlf & _
                      "    3. There exists more than one publ1c procedure with the same name"  & vbcrlf & _
                      "    4. The Procedure has the same name as the component it sits in"     & vbcrlf & _
                      "    5. The Workbook with its VBProject isnt open"                       & vbcrlf & _
                      "    6. The parameters were passed wrong"                                & vbcrlf '1 in publ1c to not mess with GetAllProcedures
        Case 3
            Message = " You passed too many arguments, VBA limits ParamArray Arguments to 30 (1 up to 30)"
        Case 4
            Message = "PrintConsole or PrintEnter didnt recieve equal or more elements than the text passed"                 & vbcrlf & _
                      "To ensure this not happening pass 0 elements for in_basic, 1 to paint all chars to that color,"       & vbcrlf & _
                      "or if you want all individually then pass equal or more elements through the optional color argument"
        End Select
        PrintEnter Message, Color
    End Sub

    Private Static Function HandleOtherKeys(pKey As Long, ByVal ShiftKey As Integer)

        Static CapitalKey As Boolean
        Dim asciiChar As String
        Dim CurrentWord As String
        Dim CurrentLine As String
        
        ' Adjust for Shift key (Uppercase letters, special characters)
        CurrentLine = GetLine(ConsoleText.Text, CurrentLineIndex)
        CurrentWord = GetWord(CurrentLine)
        If pKey = vbKeyCapital Then
            CapitalKey = CapitalKey Xor True
            GoTo SkipKey
        End If
        If CapitalKey = True Then ShiftKey = 1
        Select Case ShiftKey
            Case 0
                ' Base character
                Select Case pKey
                    Case vbKeyA To vbKeyZ:      asciiChar = LCase(Chr(pKey))
                    Case vbKey0 To vbKey9:      asciiChar = Chr(pKey)
                    Case vbKeySpace:            asciiChar = " "
                    Case vbKeyBack:             asciiChar = Chr(8) ' Backspace
                    Case vbKeyReturn:           asciiChar = Chr(13) ' Carriage Return
                    Case vbKeyTab:              asciiChar = Chr(9) ' Tab
                    Case vbKeyMultiply:         asciiChar = "*"
                    Case vbKeyAdd, 187:         asciiChar = "+"
                    Case vbKeySubtract, 189:    asciiChar = "-"
                    Case vbKeyDecimal, 190:     asciiChar = "."
                    Case vbKeyDivide:           asciiChar = "/"
                    Case 188:                   asciiChar = ","
                    Case 191:                   asciiChar = "#"
                    Case 226:                   asciiChar = "<"
                    Case vbKeyRight:            asciiChar = "RIGHT"
                End Select
            Case = 1
                Select Case pKey
                    Case vbKeyA To vbKeyZ:      asciiChar = UCase(asciiChar)
                    Case vbKey1:                asciiChar = "!"
                    Case vbKey2:                asciiChar = Chr(34) ' """
                    Case vbKey3:                asciiChar = "§"
                    Case vbKey4:                asciiChar = "$"
                    Case vbKey5:                asciiChar = "%"
                    Case vbKey6:                asciiChar = "&"
                    Case vbKey7:                asciiChar = "/"
                    Case vbKey8:                asciiChar = "("
                    Case vbKey9:                asciiChar = ")"
                    Case vbKey0:                asciiChar = "="
                    Case 187:                   asciiChar = "*"
                    Case 188:                   asciiChar = ";"
                    Case 189:                   asciiChar = "_"
                    Case 190:                   asciiChar = ":"
                    Case 191:                   asciiChar = "'"
                    Case 226:                   asciiChar = ">"
                End Select
            Case 2

            Case 3
                    Case 226:                   asciiChar = "|"
        End Select
        Select Case asciiChar
            Case Chr(9), "RIGHT" ' TAB
                If Intellisense_Active = True Then IntellisenseList.Visible = True: IntellisenseList.SetFocus
            Case Else
                SetUp_IntelliSenseList CurrentWord
        End Select
        SkipKey:
        SetPositions
        ColorWord

    End Function

    Private Sub SetPositions()
        Dim Temp()       As String: Temp = Split(ConsoleText.Text, vbCrLf)
        Dim FactorHeight As Double: FactorHeight = Height / 4
        Dim FactorWidth  As Double: FactorWidth = Width / 8
        Dim ListOffset   As Double: ListOffset = 1.45
        Dim ListFactor   As Double: ListFactor = 1.35
        Dim CurrentLine  As String: CurrentLine = GetLine(ConsoleText.Text, UBound(Temp))
        ScrollTop = UBound(Temp) * 10 * ListFactor - FactorHeight - 100
        ScrollLeft = Len(CurrentLine) * 10
        IntellisenseList.Top = ScrollTop + (FactorWidth * ListOffset) + 100
        IntellisenseList.Left = ScrollLeft
        IntellisenseList.Visible = True
    End Sub

    Private Sub ColorWord()

        Dim Temp As String
        Dim Para_Counter As Long
        Dim Color As Long
        Dim Lines() As String
        Dim Words() As String
        Dim CurrentWord As String
        Dim PreviousWord As String
        Dim Is_String As Boolean
        Dim CurrentLine As String
        Dim CurrentLinePoint As Long
        Dim Tempp As Long
        Dim i As Long, j As Long

        CurrentLine = GetLine(ConsoleText.Text, CurrentLineIndex)
        Lines = Split(ConsoleText.Text, vbCrLf)
        CurrentLinePoint = Len(ConsoleText.Text) - Len(Lines(Ubound(Lines))) - 4
        Tempp = InStr(1, Lines(Ubound(Lines)), Recognizer)
        If Tempp <> 0 Then CurrentLinePoint = CurrentLinePoint + Tempp + Len(Recognizer)
        Words = Split(CurrentLine, " ")

        For i = 0 To UBound(Words)
            Dim TempCounter As Long
            For j = 0 To i - 1
                TempCounter = TempCounter + Len(Words(j)) + 1 ' one for space
            Next j
            ConsoleText.SelStart = CurrentLinePoint + TempCounter
            TempCounter = 0
            ConsoleText.SelLength = Len(Words(i))
            Select Case UCase(Words(i))
                Case "IF", "THEN", "ELSE", "END", "FOR", "EACH", "NEXT", "DO", "WHILE", "LOOP", "SELECT", "CASE", "EXIT", "CONTINUE"
                    Color = in_Statement
                Case "DIM", "PUBLIC", "PRIVATE", "GLOBAL", "TRUE", "FALSE", "FUNCTION", "SUB", "REDIM", "PRESERVE"
                    Color = in_Keyword
                Case "+", "*", "/", "-", "^", ":", ";", "<", ">", "=", "!", "|", "<>", "NOT", "AND", "OR", "XOR", "!=", "++", "||", "&", "&&", "=>", "=<", "<=", ">="
                    Color = in_Operator
                Case Else
                    If Ucase(PreviousWord) = "AS" Then
                        Color = in_Datatype
                    ElseIf Words(i) Like "*(*)" Then
                        Temp = Mid(Words(i), 1, InStr(1, Words(i), "(") - 1)
                        ConsoleText.SelLength = Len(Temp)
                        Color = in_Procedure
                    ElseIf InArray(Words(i), VariableArray) Then
                        Color = in_Variable
                    Else
                        Color = in_Basic
                    End If
            End Select
            ConsoleText.SelColor = Color
            PreviousWord = Words(i)
        Next i
        Color = 0
        For i = 1 To Len(CurrentLine)
            If Is_String = True Then
                Color = in_String
            Else
                Select Case Mid(CurrentLine, i, 1)
                    Case = "("
                        Color = in_Parantheses
                        Para_Counter = Para_Counter + 1
                        If Para_Counter Mod 2 = 0 Then Color = Color + &H00333300
                        ConsoleText.SelColor = Color
                    Case = ")"
                        Color = in_Parantheses
                        If Para_Counter Mod 2 = 0 Then Color = Color + &H00333300
                        ConsoleText.SelColor = Color
                        Para_Counter = Para_Counter - 1
                    Case "+", "*", "/", "-", "^", ":", ";", "<", ">", "=", "!", "|"
                        Color = in_Operator
                    Case Chr(34)
                        Color = in_String
                        Is_String = Is_String Xor True
                    Case Else
                        Color = 0
                End Select
            End If
            If Color <> 0 Then
                ConsoleText.SelStart = CurrentLinePoint + i - 1 ' 1 comes from offset between .text and .seltext
                ConsoleText.SelLength = 1
                ConsoleText.SelColor = Color
            End If
        Next i
        ConsoleText.SelStart = Len(ConsoleText.Text)
        ConsoleText.SelColor = in_Basic

    End Sub

    Private Function GetStartText() As String
        GetStartText =                   _
        "VBA Console [Version 1.0]" & Chr(13) & Chr(10) & _
        "No Rights reserved"        & Chr(13) & Chr(10) & _
        Chr(13) & Chr(10)                               & _
        PrintStarter
    End Function

    Private Function HandleClear()
        ConsoleText.Text = "_"
        ConsoleText.SelStart = 0
        ConsoleText.SelLength = Len(ConsoleText.Text)
        ConsoleText.SelColor = in_Basic
        SetUpNewLine
    End Function

    Private Sub HandleHelp()

        Dim Message As String
        Dim Color As Long

        Color = &H100000FF - &H10000000
        Message = _
        "--------------------------------------------------"                                           & vbcrlf & _
        "This Console can do the following:"                                                           & vbcrlf & _
        "1. It can be used as a form to showw messages, ask questions to the user or get a user input" & vbcrlf & _
        "2. It can be used to showw and log errors and handle them by user input"                      & vbcrlf & _
        "3. It can run Procedures with up to 29 arguments"                                             & vbcrlf & _
        ""                                                                                             & vbcrlf & _
        "HOW TO USE IT:"                                                                               & vbcrlf & _
        "   Run a Procedure:"                                                                          & vbcrlf & _
        "       To run a procedure you have to write the name of said procedure (Case sensitive)"      & vbcrlf & _
        "       If you want to pass parameters you have to write    |; | between every parameter"      & vbcrlf & _
        "       Example:"                                                                              & vbcrlf & _
        "           Say; THIS IS A PARAMETER; THIS IS ANOTHER PARAMETER"                               & vbcrlf & _
        ""                                                                                             & vbcrlf & _
        "   Ask a question:"                                                                           & vbcrlf & _
        "       Use CheckPredeclaredAnswer"                                                            & vbcrlf & _
        "           Param1 = Message to be showwn"                                                     & vbcrlf & _
        "           Param2 = Array of Values, which are acceptable answers"                            & vbcrlf & _
        "           Param3 = Array of Messages, which showw a text according to answer in Param2"      & vbcrlf & _
        "       The Function will loop until one of the acceptable answers is typed"                   & vbcrlf & _
        "--------------------------------------------------"                                           & vbcrlf
        PrintEnter Message, Color

    End Sub

    Private Function InStrAll(Text As String, SearchText As String, Optional StartIndex As Long = 1, Optional EndIndex As Long = 0, Optional StartFinding As Long = 0, Optional ReturnCount As Long = 255, Optional Line As Long = 0, Optional BreakText As String = Empty) As Long()
        
        Dim ReturnArray() As Long
        Dim Lines() As String
        Dim CurrentValue As Long: CurrentValue = 0
        Dim Found As Long: Found = 0
        Dim Saved As Long: Saved = -1
        Dim j As Long: j = 0
        Dim i As Long: i = StartIndex
        Dim EndLine As Long
        If EndIndex = 0 Then EndIndex = Len(Text)
        If BreakText <> Empty Then
            Lines = Split(Text, BreakText)
            EndLine = Line
        Else
            ReDim Lines(0)
            Lines(0) = Text
            EndLine = Ubound(Lines)
        End If
        
        For j = Line To EndLine
            Do Until i = EndIndex
                CurrentValue = 0
                CurrentValue = InStr(i, Lines(j), SearchText)
                If CurrentValue <> 0 Then 
                    Found = Found + 1
                Else
                    Exit Do
                End If
                If Found >= StartFinding Then
                    Saved = Saved + 1
                    ReDim Preserve ReturnArray(Saved)
                    ReturnArray(Saved) = CurrentValue
                End If
                If Saved = ReturnCount Then Exit For
                
                i = CurrentValue + Len(SearchText)
                If i > Len(Lines(j)) Then Exit Do
            Loop
            i = 1
        Next j
        InStrAll = ReturnArray

    End Function

'



' Following Part is Intellisense, which is dependant on Microsoft Visual Studio extensebility 5.3
    Private Sub GetAllProcedures()

        Dim WB As Workbook
        Dim VBProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        Dim CodeMod As VBIDE.CodeModule
        
        Dim CurrentRow As String
        Dim StartPoint As Long
        Dim EndPoint As Long
        Dim Name As String
        Dim ReturnType As Variant
        Dim TempArg() As String

        Dim i As Integer
        Dim j As Integer
        Dim ProcedureCount As Long
        Dim Temp As Long

        ' This is to get the second last space, which indicates, that its the startpoint for the returntype
        Dim TempArray() As String

        ProcedureCount = 1000
        ReDim VariableArray(MaxArg, ProcedureCount)
        ProcedureCount = 0

        For Each WB In Workbooks
            Set VBProj = WB.VBProject
            For Each VBComp In VBProj.VBComponents
                    Set CodeMod = VBComp.CodeModule
                    For i = 1 To CodeMod.CountOfLines
                        CurrentRow = CodeMod.Lines(i, 1)
                        If UCase(CurrentRow) Like "*PUBLIC *" And InStr(1, CurrentRow, "'") = 0  And Not UCase(CurrentRow) Like "*" & Chr(34) & "*PUBLIC*" & Chr(34) & "*" Then
                            If (UCase(CurrentRow) Like "* FUNCTION *" Or UCase(CurrentRow) Like "* SUB *") Then
                                ' A Procedure
                                '                          |----------|
                                '   Public Static Function VariableName(Arg1 As Variant, Arg2 As Variant) As Variant
                                Select Case True
                                    Case UCase(CurrentRow) Like "*PUBLIC STATIC SUB *(*)*"      : StartPoint = InStr(1, UCase(CurrentRow), "PUBLIC STATIC SUB ")      + Len("PUBLIC STATIC SUB ")      : EndPoint = InStr(1, UCase(CurrentRow), "("): Name = Mid(CurrentRow, StartPoint, EndPoint - StartPoint)   
                                    Case UCase(CurrentRow) Like "*PUBLIC SUB *(*)*"             : StartPoint = InStr(1, UCase(CurrentRow), "PUBLIC SUB ")             + Len("PUBLIC SUB ")             : EndPoint = InStr(1, UCase(CurrentRow), "("): Name = Mid(CurrentRow, StartPoint, EndPoint - StartPoint)   
                                    Case UCase(CurrentRow) Like "*PUBLIC STATIC FUNCTION *(*)*" : StartPoint = InStr(1, UCase(CurrentRow), "PUBLIC STATIC FUNCTION ") + Len("PUBLIC STATIC FUNCTION ") : EndPoint = InStr(1, UCase(CurrentRow), "("): Name = Mid(CurrentRow, StartPoint, EndPoint - StartPoint)   
                                    Case UCase(CurrentRow) Like "*PUBLIC FUNCTION *(*)*"        : StartPoint = InStr(1, UCase(CurrentRow), "PUBLIC FUNCTION ")        + Len("PUBLIC FUNCTION ")        : EndPoint = InStr(1, UCase(CurrentRow), "("): Name = Mid(CurrentRow, StartPoint, EndPoint - StartPoint)
                                    Case UCase(CurrentRow) Like "*PUBLIC PROPERTY GET *(*)*"    : StartPoint = InStr(1, UCase(CurrentRow), "PUBLIC PROPERTY GET ")    + Len("PUBLIC PROPERTY GET ")    : EndPoint = InStr(1, UCase(CurrentRow), "("): Name = Mid(CurrentRow, StartPoint, EndPoint - StartPoint)
                                    Case UCase(CurrentRow) Like "*PUBLIC PROPERTY SET *(*)*"    : StartPoint = InStr(1, UCase(CurrentRow), "PUBLIC PROPERTY SET ")    + Len("PUBLIC PROPERTY SET ")    : EndPoint = InStr(1, UCase(CurrentRow), "("): Name = Mid(CurrentRow, StartPoint, EndPoint - StartPoint)   
                                    Case Else
                                End Select
                                '                                       |------------------------------|
                                '   Public Static Function VariableName(Arg1 As Variant, Arg2 As Variant) As Variant
                                    StartPoint = InStr(1, CurrentRow, "(")
                                    EndPoint   = InStr(1, CurrentRow, ")")
                                    If StartPoint + 1 <> EndPoint Then
                                        TempArg = Split(Mid(CurrentRow, StartPoint + 1, EndPoint - StartPoint - 1), ",")
                                        For j = 0 To Ubound(TempArg)
                                            VariableArray(j + 4, ProcedureCount) = TempArg(j)
                                        Next
                                    End If
                                    Temp = 1
                            Else
                                ' A Variable
                                '          |----------|
                                '   Public VariableName As Variant
                                Select Case True
                                    Case UCase(CurrentRow) Like "*PUBLIC CONST *": StartPoint = InStr(1, UCase(CurrentRow), "*PUBLIC CONST *") + Len("PUBLIC CONST "): EndPoint = InStr(1, UCase(CurrentRow), " AS "): Name = Mid(CurrentRow, StartPoint + 1, EndPoint - StartPoint - 1)
                                    Case UCase(CurrentRow) Like "*PUBLIC *":       StartPoint = InStr(1, UCase(CurrentRow), "*PUBLIC *")       + Len("PUBLIC "):       EndPoint = InStr(1, UCase(CurrentRow), " AS "): Name = Mid(CurrentRow, StartPoint + 1, EndPoint - StartPoint - 1)
                                    Case Else
                                End Select
                                Temp = 1
                            End If
                                '                                                                        |---------|
                                '   Public Static Function VariableName(Arg1 As Variant, Arg2 As Variant) As Variant
                                TempArray = Split(CurrentRow, " ")
                                ReturnType = TempArray(Ubound(TempArray) - 1) & " " & TempArray(Ubound(TempArray))
                                ' If last character is ")", then it returns nothing
                                If Mid(TempArray(Ubound(TempArray)), Len(TempArray(Ubound(TempArray))), 1) = ")" Then ReturnType = "Void"

                                VariableArray(0, ProcedureCount) = VBProj.Name
                                VariableArray(1, ProcedureCount) = VBComp.Name
                                VariableArray(2, ProcedureCount) = Name
                                VariableArray(3, ProcedureCount) = ReturnType
                                ProcedureCount = ProcedureCount + Temp
                                Temp = 0
                                ' Add another 1000 Procedures
                                If ProcedureCount > UBound(VariableArray, 2) Then
                                    ReDim Preserve VariableArray(MaxArg, UBound(VariableArray, 2) + 1000)
                                End If
                        End If
                        NoLines:
                    Next
            Next
        Next
                
    End Sub
    
    Private Sub Close_IntelliSenseList()
        IntelliSenseList.Clear
        IntelliSenseList.Visible = False
        ConsoleText.SetFocus
        ConsoleText.SelStart = Len(ConsoleText.Text)
    End Sub

    Private Sub SetUp_IntelliSenseList(Text As String)
        
        Dim i As Long, j As Long, x As Long
        Dim StartPoint As Long
        Dim EndPoint As Long
        Dim Found() As String
        Dim Foundd As Boolean
        Dim Count As Long
        Dim Words() As String
        Dim AbstractionFound As Boolean


        Dim AbstractionDepth As Long: AbstractionDepth = 0
        Dim LeftNoStringValue As Long:  LeftNoStringValue = InStr(1, Text, Chr(34))
        Dim RightNoStringValue As Long: RightNoStringValue = InStr(LeftNoStringValue + 1, Text, Chr(34))
        '|---------------|      &    |----|     Dont know why this should be needed, but just in case
        'Project.Module.Fu"dfdf.sfd "nction
        If LeftNoStringValue <> 0 Then Text = Left(Text, LeftNoStringValue) & Right(Text, Len(Text) - RightNoStringValue)
        AbstractionDepth = Len(Text) - Len(Replace(Text,".",""))
        If AbstractionDepth > 2 Then Exit Sub ' Failsave for too much abstraction

        Words = Split(Text, ".")
        If Ubound(Words) = -1 Then Exit Sub
        Text = Words(Ubound(Words))

        IntelliSenseList.Clear
        Redim Found(0)
        If AbstractionDepth > 0 Then
            For x = AbstractionDepth To 2
                For i = 0 To Ubound(VariableArray, 2)
                    If UCase(VariableArray(x - 1, i)) = UCase(Words(Ubound(Words) - 1)) Then
                        StartPoint = i
                        AbstractionFound = True
                        AbstractionDepth = x
                        Exit For
                    End If
                Next
                EndPoint = StartPoint
                If i =< Ubound(VariableArray, 2) Then 
                    Do While UCase(VariableArray(x - 1, i)) = UCase(Words(Ubound(Words) - 1))
                        EndPoint = i
                        i = i + 1
                        If i > Ubound(VariableArray, 2) Then Exit Do
                    Loop
                End If
                If AbstractionFound = True Then Exit For
            Next
        End If
        If AbstractionFound <> True Then
            StartPoint = 0
            EndPoint = Ubound(VariableArray, 2)
        End If
        For x = AbstractionDepth To 2
            For i = StartPoint To EndPoint
                If UCase(VariableArray(x, i)) Like Ucase(Text) & "*" Then
                    For j = 0 To Ubound(Found)
                        If Found(j) = VariableArray(x, i) Then
                            Foundd = True
                            Exit For
                        End If
                    Next
                    If Foundd <> True Then
                        IntelliSenseList.AddItem
                        IntelliSenseList.List(Ubound(Found), 0) = VariableArray(x, i)
                        If x = 2 Then IntelliSenseList.List(Ubound(Found), 1) = GetProcedureText(i) 
                        Found(Ubound(Found)) = VariableArray(x, i)
                        Redim Preserve Found(Ubound(Found) + 1)
                    End If
                    Foundd = False
                End If
            Next
        Next

    End Sub

    Function InArray(stringToBeFound As String, arr As Variant) As Boolean
        Dim Element As Variant
        For Each Element in arr
            If UCase(Cstr(Element)) = UCase(stringToBeFound) Then
                InArray = True
                Exit Function
                End If
        Next Element
    End Function
    
    Private Function GetProcedureText(Index As Long) As String
        Dim i As Long
        Dim Text As String
        Text                  = VariableArray(0, Index) & " " & _
                                VariableArray(1, Index) & " " & _
                                VariableArray(2, Index) & " " & _
                                "("
        For i = 4 To MaxArg
            Text = Text & ", " & VariableArray(i, Index)
        Next
        Text                  = VariableArray(3, Index) & _
                                ")"
        GetProcedureText = Text
    End Function

    Private Sub IntelliSenseList_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
        
        Dim Line As String
        Dim Word As String
        Dim Words() As String
        Dim Start As Long
        Dim ReturnString As String

        Line = GetLine(ConsoleText.Text, CurrentLineIndex)
        Dim AbstractionDepth As Long: AbstractionDepth = 0
        Dim LeftNoStringValue As Long:  LeftNoStringValue = InStr(1, Line, Chr(34))
        Dim RightNoStringValue As Long: RightNoStringValue = InStr(LeftNoStringValue + 1, Line, Chr(34))
        '|---------------|      &    |----|     Dont know why this should be needed, but just in case
        'Project.Module.Fu"dfdf.sfd "nction
        If LeftNoStringValue <> 0 Then Line = Left(Line, LeftNoStringValue) & Right(Line, Len(Line) - RightNoStringValue)
        AbstractionDepth = Len(Line)-Len(Replace(Line,".",""))

        Words = Split(Line, ".")
        If Ubound(Words) = -1 Then Word = Line
        Word = Words(Ubound(Words))
        Select Case KeyCode
            Case vbKeyLeft
                Close_IntelliSenseList
                Exit Sub
            Case vbKeyTab, vbKeyRight
                If IntellisenseList.ListCount > 0 Then
                    ReturnString = IntelliSenseList.List(IntelliSense_Index, 0)
                    Start = InStr(1, Ucase(ReturnString), Ucase(Word))
                    If Start = 0 Then Start = 1
                    PrintConsole Mid(ReturnString, Start + Len(Word), Len(ReturnString) - Len(Word))
                    Close_IntelliSenseList
                    Exit Sub
                End If
            Case vbKeyUp
                Intellisense_Index = Intellisense_Index - 1
            Case vbKeyDown
                Intellisense_Index = Intellisense_Index + 1
        End Select
        If Intellisense_Index > IntelliSenseList.ListCount - 1 Then
            Intellisense_Index = 0
        ElseIf Intellisense_Index < 0 Then
            Intellisense_Index = IntelliSenseList.ListCount - 1
        Else

        End If
        If IntelliSenseList.ListCount > 0 Then IntelliSenseList.ListIndex = IntelliSense_Index

    End Sub
'