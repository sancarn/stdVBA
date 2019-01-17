VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "STD_Debug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pFilePath As String
Private pWarnings As New Collection
Private pErrors As New Collection
Private pRAWUse As Boolean
Private pRAW As New Collection


'*********************
'* WARNINGS PROPERTY *
'*********************
Public Property Get warnings() As Collection
    Set warnings = pWarnings
End Property
Public Property Let warnings(warnings As Collection)
    Set pWarnings = warnings
End Property

'*******************
'* ERRORS PROPERTY *
'*******************
Public Property Get errors() As Collection
    Set errors = pErrors
End Property
Public Property Let errors(warnings As Collection)
    Set pErrors = errors
End Property

'****************
'* LOG PROPERTY *
'****************
Public Property Get FilePath() As String
    FilePath = pFilePath
End Property
Public Property Let FilePath(FilePath As String)
    pFilePath = FilePath
End Property

'******************
'* RAW PROPERTIES *
'******************
Public Property Get RAWUse() As Boolean
    RAWUse = pRAWUse
End Property
Public Property Let RAWUse(RAWUse As Boolean)
    pRAWUse = RAWUse
End Property

Public Property Get RAW() As Collection
    Set RAW = pRAW
End Property
Public Property Let RAW(RAW As Variant)
    Set pRAW = RAW
End Property



Public Sub printt(var As Variant, Optional delimeter As String = ",")
    'This is a better version of debug.print()
    'This version tries to print objects and arrays as well as all other types of common variable types.
    Debug.Print getStrings(var, delimeter)
End Sub
Public Sub log(var As Variant, Optional delimeter As String = ",", Optional ByVal append As Boolean = True)
    'This is a better version of debug.print()
    'This version tries to print objects and arrays as well as all other types of common variable types.
    'The output of debugg.log() is a text file instead of the immediate window.
    'The file path for the text file is specified in debugg.filePath
    
    'if no file path given throw an error
    If Not (pFilePath Like "*.*") Then Err.Raise -1, "clsDebugg", "No filePath path has been provided"
    
    'Get the free file number
    Dim n As Long
    n = FreeFile
    
    'if you want to append the data do so, otherwise open file to output.
    If append Then
        Open pFilePath For Append As #n
    Else
        Open pFilePath For Output As #n
    End If
    
    Print #n, getStrings(var, delimeter)
    
    Close #n
End Sub
Public Sub report(var As Variant, Optional delimeter As String = ",", Optional ByVal append As Boolean = True)
    'debugg.Report() executes log AND printt in a speed efficient manner.
    
    'Get the free file number
    Dim n As Long
    n = FreeFile
    
    Dim str As String
    str = getStrings(var, delimeter)
    
    Debug.Print str
    
    'if you want to append the data do so, otherwise open file to output.
    If append Then
        Open pFilePath For Append As #n
    Else
        Open pFilePath For Output As #n
    End If
    
    Print #n, str
    
    Close #n
End Sub

Public Sub createReport(Optional ByVal macroName As String = "Macro", Optional FileName As String = "errorLog", Optional header As String = "")
    'Get current store pRAWUse state and then set it to false.
    Dim curState As Boolean: curState = pRAWUse
    pRAWUse = False
    
    'Output to filePath
    pFilePath = ThisWorkbook.path & "\" & FileName & Replace(Replace(Date & Time, ":", ""), "/", "") & ".txt"
    
    'Define a seperator for each section of the log file.
    Const reportSeperator = vbCrLf & "--------------------------------------------------" & vbCrLf
    
    'Print introduction to log file
    report macroName & " processing..."
    report ThisWorkbook.Name & "          " & Time & " " & Date
    If header <> "" Then report header
    report reportSeperator
    
    'Report RAW log
    report pRAW
    report reportSeperator
    
    'report errors
    report "Errors:" & vbCrLf
    report pErrors
    report reportSeperator
    
    'report warnings
    report "Warnings:" & vbCrLf
    report pWarnings
    report reportSeperator
    
    'report end of report
    report "END OF REPORT"
    report reportSeperator
    
    'reset pRAW back to previous state
    pRAWUse = curState
End Sub


Private Function getStrings(ByVal var As Variant, Optional ByVal delimeter As String = ",") As String
    Dim str As String, i As Long, vData As Variant
    
    'explore Var, try to find
    If IsObject(var) Then
        'If var is an object we can treat it like a collection.
        
        'Loop through collection and build string
        For Each vData In var
            str = str & vData & vbCrLf
        Next vData
        
        'Print string
        getStrings = str
        
    ElseIf IsArray(var) Then
        'if array is 1 dimensional
        If UBound(var, 2) = 1 Then
            'grab all data into string -- Can we just use join() here?
            For i = LBound(var) To UBound(var)
                str = str & var(i) & vbCrLf
            Next i
            
            'Remove last vbCrLf from string -- Question: Does a vbCrLf count as 2 characters or just 1? ==> 2
            str = Left(str, Len(str) - Len(vbCrLf))
            getStrings = str
            
            
        'if array is 2 dimensional then print as csv (or delimeted with delimeter specified)
        ElseIf UBound(var, 2) > 1 And UBound(var, 3) = 1 Then
            'grab all data into string - note Join()
            For i = LBound(var) To UBound(var)
                str = str & Join(var(i), delimeter) & vbCrLf
            Next i
            
            'Remove last vbCrLf from string
            str = Left(str, Len(str) - Len(vbCrLf))
            getStrings = str
        Else
            Err.Raise -1, "Debugg", "3 Dimensional arrays are not supported by this sub-routine."
        End If
    Else
        'If not an array or an object then it must be a regular variable.
        'Thus print regular variable
        getStrings = var
    End If
    
    If pRAWUse Then pRAW.Add getStrings
End Function

'******************************************
'* EXAMPLE OF HOW TO USE THE DEBUGG CLASS *
'******************************************

Private Sub debuggExample()
    Dim o As Object
    Dim str() As String
    Dim debugg As clsDebugg
    Set debugg = New clsDebugg
    
    'Add errors
    debugg.errors.Add "Error 1"
    debugg.errors.Add "Error 2"
    debugg.errors.Add "Error 3"
    debugg.errors.Add "Error 4"
    debugg.errors.Add "Error 5"
    
    'Add warnings
    debugg.warnings.Add "Warning 1"
    debugg.warnings.Add "Warning 2"
    debugg.warnings.Add "Warning 3"
    debugg.warnings.Add "Warning 4"
    
    'Output to filePath
    debugg.FilePath = "C:\Users\jwa\Desktop\Test\logs\abc.txt"
    debugg.report "Errors:" & vbCrLf
    debugg.report debugg.errors
    debugg.report "------------------------------------------------"
    debugg.report "Warnings:" & vbCrLf
    debugg.report debugg.warnings
    debugg.report "------------------------------------------------"
End Sub







'*****************************************************************************************
'*****************************************************************************************
'**                                                                                     **
'**                                  .|'''.|   .|'''.|                                  **
'**                                  ||..  '   ||..  '                                  **
'**                                   ''|||.    ''|||.                                  **
'**                                 .     '|| .     '||                                 **
'**                                 |'....|'  |'....|'                                  **
'**                                                                                     **
'*****************************************************************************************
'*****************************************************************************************
'* Font: kban   Reflection: no   Adjustment: left   Stretch: no      Width: 80  Text: SS *
'*                         http://www.network-science.de/ascii/                          *
'*****************************************************************************************

        'Public Sub printt(var As Variant, Optional delimeter As String = ",")
        '    'This is a better version of debug.print()
        '    'This version tries to print objects and arrays as well as all other types of common variable type.
        '
        '    'explore Var, try to find
        '    If IsObject(var) Then
        '        'If var is an object we can treat it like a collection.
        '        Dim vData As Variant, str As String
        '
        '        'Loop through collection and build string
        '        For Each vData In var
        '            str = str & vData & vbCrLf
        '        Next vData
        '
        '        'Print string
        '        Debug.Print str
        '
        '    ElseIf IsArray(var) Then
        '        'if array is 1 dimensional
        '        If UBound(var, 2) = 1 Then
        '            Dim str As String, i As Long
        '            'grab all data into string
        '            For i = LBound(var) To UBound(var)
        '                str = str & var(i) & vbCrLf
        '            Next i
        '
        '            'Remove last vbCrLf from string -- Question: Does a vbCrLf count as 2 characters or just 1?
        '            str = Replace(str & vbCrLf, vbCrLf & vbCrLf, "")
        '            Debug.Print str
        '        'if array is 2 dimensional then print as csv (or delimeted with delimeter specified)
        '        ElseIf UBound(var, 2) > 1 And UBound(var, 3) = 1 Then
        '            Dim i As Long, str As String
        '            'grab all data into string - note Join()
        '            For i = LBound(var) To UBound(var)
        '                str = str & Join(var(i), delimeter) & vbCrLf
        '            Next i
        '
        '            'Remove last vbCrLf from string -- Question: Does a vbCrLf count as 2 characters or just 1?
        '            str = Replace(str & vbCrLf, vbCrLf & vbCrLf, "")
        '            Debug.Print str
        '        Else
        '            Debug.Print "3 Dimensional arrays are not supported by this sub-routine"
        '        End If
        '    Else
        '        'If not an array or an object then it must be a regular variable.
        '        'Thus print regular variable
        '        Debug.Print var
        '    End If
        'End Sub
