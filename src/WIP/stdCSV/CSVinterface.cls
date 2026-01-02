VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSVinterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' CONSTANTS:
Private Const ASCW_CHR_BACKSLASH As Long = 92
Private Const CHR_APOSTROPHE As String = "'"
Private Const CHR_BACKSLASH As String = "\"
Private Const CHR_DOUBLE_QUOTES As String = """"
Private Const CHR_TILDE As String = "~"
Private Const CHR_CARET As String = "^"
Private Const CHR_LSQRB As String = "{"
Private Const CHR_RSQRB As String = "}"
'////////////////////////////////////////////////////////////////////////////////////////////
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' PROPERTIES VARIABLES:
Private config As CSVparserConfig '--------------Parser configuration object
Private CSVstream As CSVTextStream '-------------Enables file Stream
Private P_CSV_DATA As CSVArrayList '-------------Holds the CSV data for current instance
Private P_CSV_HEADER As CSVArrayList '-----------Holds the CSV header
Private P_ERROR_DESC As String '-----------------Holds the Error Description
Private P_ERROR_NUMBER As Long '-----------------Holds the Error number
Private P_ERROR_SOURCE As String '---------------Holds the Error source
Private P_SUCCESSFUL_EXPORT As Boolean
Private P_SUCCESSFUL_IMPORT As Boolean
Private P_VARYING_LENGTHS As Boolean '-----------Indicates if the read CSV has varying number _
                                                 of fields per records
Private P_VECTORS_MAX_BOUND As Long '------------Maximum index of vectors on jagged array
Private P_VECTORS_REGULAR_BOUND As Long '--------Regular upper index of vectors on jagged array
'////////////////////////////////////////////////////////////////////////////////////////////
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' VARIABLES:
' @Common
Private ASCIIcharw As Long '---------------AscW() applied to the current char
Private dTTargets() As Variant '-----------Template's links to fields
Private dTTemplate() As Variant '----------Template
Private EscapeAscW As Long '---------------AscW() applied to the quote char
Private EscapeSequence As String '---------Classic or unix escape sequence
Private EscapeSequence_Delim As String '---Classic or unix escape sequence in fields delimiter
Private FDAscW As Long '-------------------AscW() applied to the field delimiter char
Private FileHandled As Integer '-----------Pointer to the actual CSV file
Private IgnoreEmptyLines As Boolean '------Skip empty lines
Private index As Long '--------------------Pointer to the actual record or chunk of record
Private LenCurrentIndex As Long '----------Length of the actual record or chunk of record
Private MaxIndex As Long '-----------------Maximum number of records or chunks of records
Private OpenedToken As Boolean '-----------Multiline field indicator
Private OverStringPointer As Long '--------Pointer over actual record index string
Private QuoteChar As String '--------------Quote char
Private RecordDelimiter As String '--------Record delimiter char
Private RequestedFieldsArray() As Long '---List of the indexes of the fields to be requested
Private StreamEnd As Boolean '-------------End of file indicator, when using streams
Private StreamWithQuotes As Boolean '------Indicates whether the stream holds a quote char
Private TargetsDefined As Boolean '--------Indicates whether the DTT targets are defined
Private TemplateDefined As Boolean '-------Indicates whether the DTT is defined
Private tmpCSV() As String '---------------Stream or string split using record delimiter
Private UnixEscpedSep As Boolean '---------Unix escaped field delimiter or record delimiter
Private UnixEscapeSeq As String '----------Current unix escape sequence
Private UnixEscapeSeqDel As String '-------Current unix escape sequence for fields delimiter
Private UnixEscapeSeqRec As String '-------Current unix escape sequence for records delimiter
Private UseUnixEscapeSeq As Boolean '------Determines when to use or not to use the unix escape sequence
'@------------------------------------------------------------------------------------------------
' SEQUENTIAL READ VARIABLES
Private static_ASCIIcharw As Long
Private static_commToken As Long
Private static_CSVstream As CSVTextStream
Private static_dTTargets() As Variant
Private static_dTTemplate() As Variant
Private static_dynamicType As Boolean
Private static_EndOF As Boolean
Private static_EndOFStreamBuffer As Boolean
Private static_EscapeAscW As Long
Private static_EscapeChrCount As Long
Private static_EscapedBRS As Boolean
Private static_EscapedWFS As Boolean
Private static_EscapeSequence As String
Private static_EscapeSequence_Delim As String
Private static_FDAscW As Long
Private static_FieldDelimiter As String
Private static_FieldID As Long
Private static_Headers As Boolean
Private static_IgnoreCommentLines As Boolean
Private static_IgnoreEmptyLines As Boolean
Private static_IgnoreLines As Boolean
Private static_ImportSwitch As Boolean
Private static_Index As Long, static_MaxIndex As Long
Private static_IsHeader As Boolean
Private static_IsWellEscapedField As Boolean
Private static_LenCurrentIndex As Long
Private static_MaxReqIndex As Long
Private static_NotEvenEscapeChrNumber As Boolean
Private static_OpenedToken As Boolean
Private static_OpenTokenStreamIdx As Long
Private static_OpenTokenArrIdx As Long
Private static_StreamsCounter As Long
Private static_OpenTokenStartPos As Long
Private static_outputList As CSVArrayList
Private static_OverStringPointer As Long, static_hpointer As Long
Private static_QuoteChar As String
Private static_RecordDelimiter As String
Private static_RecordEndReached As Boolean
Private static_RecordsCount As Long
Private static_RecordToken As CSVArrayList
Private static_RequestedFieldsArray() As Long
Private static_SearchBeginningMark As Long
Private static_SeqReadStarted As Boolean
Private static_SplittedToken As Boolean
Private static_StreamEnd As Boolean
Private static_StreamWithQuotes As Boolean
Private static_TargetsDefined As Boolean
Private static_TemplateDefined As Boolean
Private static_tmpCSV() As String
Private static_tmpRequested() As Variant
Private static_tmpToken As CSVArrayList
Private static_TokenBeginningPos As Long
Private static_TokenEndingPos As Long
Private static_TokenEndReached As Boolean
Private static_UnixEscpedSep As Boolean
Private static_UnixEscapeSeq As String
Private static_UseUnixEscapeSeq As Boolean
Private static_fieldsBound As Long
Private static_fieldsMaxBound As Long
'////////////////////////////////////////////////////////////////////////////////////////////
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' Events:
Public Event AfterExportRecord()
Public Event AfterImportRecord()
Public Event AfterLoadStream()
Public Event AfterSort()
Public Event AfterWriteStream()
'////////////////////////////////////////////////////////////////////////////////////////////
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' PROPERTIES:
'#
Public Property Get count() As Long
    count = P_CSV_DATA.count
End Property
Public Property Get errDescription() As String
    errDescription = P_ERROR_DESC
End Property
Public Property Get errNumber() As Long
    errNumber = P_ERROR_NUMBER
End Property
Public Property Get errSource() As String
    errSource = P_ERROR_SOURCE
End Property
Public Property Get exportSuccess() As Boolean
    exportSuccess = P_SUCCESSFUL_EXPORT
End Property
Public Property Get csvHeader() As CSVArrayList
    Set csvHeader = P_CSV_HEADER
End Property
Public Property Get importSuccess() As Boolean
    importSuccess = P_SUCCESSFUL_IMPORT
End Property
Public Property Get item(ParamArray index() As Variant) As Variant
    Select Case UBound(index)
        Case 0
            item = P_CSV_DATA(CLng(index(0)))
        Case Is >= 1
            item = P_CSV_DATA(CLng(index(0)))(CLng(index(1)))
    End Select
End Property
Public Property Get items() As CSVArrayList
    Set items = P_CSV_DATA
End Property
Public Property Set items(values As CSVArrayList)
    Set P_CSV_DATA = values
End Property
Public Property Get parseConfig() As CSVparserConfig
    Set parseConfig = config
End Property
Public Property Set parseConfig(configuration As CSVparserConfig)
    Set config = configuration
End Property
Public Property Get uniformLengthRecords() As Boolean
    uniformLengthRecords = Not P_VARYING_LENGTHS
End Property
Public Property Get fieldsBound() As Long
    fieldsBound = P_VECTORS_REGULAR_BOUND
End Property
Public Property Get fieldsMaxBound() As Long
    fieldsMaxBound = P_VECTORS_MAX_BOUND
End Property
'////////////////////////////////////////////////////////////////////////////////////////////
'#

Private Sub Class_Initialize()
    ResetToDefault
End Sub
Private Sub Class_Terminate()
    Set P_CSV_DATA = Nothing
    Set P_CSV_HEADER = Nothing
End Sub
Public Sub ClearData()
    Set P_CSV_DATA = New CSVArrayList
    Set P_CSV_HEADER = New CSVArrayList
End Sub
Public Sub CloseSeqReader()
    static_SeqReadStarted = False
    Erase static_tmpCSV
    Set static_CSVstream = Nothing
    Set static_RecordToken = Nothing
    Set static_tmpToken = Nothing
    Set static_outputList = Nothing
End Sub
''' <summary>
''' Run a left, right outer or inner join on the provided data tables.
'''' 1) Use a string such as "{1-2,5,ID};{1-6}" as a predicate of the columns to
'''' indicate the join of columns 1 to 2, 5 and ID of leftTable with
'''' the columns 1 to 6 of rightTable.
'''' 2) Use a string such as "{*};{1-3}" to indicate the union of ALL '''' columns of leftTable.
'''' columns of leftTable with columns 1 to 3 of rightTable.
''' 3) The predicate must use the dot syntax [t1.#][t1.fieldName] to.
'''' indicate the fields of the table, where t1 refers to the leftTable.
''' 4) The matchKeys predicate must be given as ''''#/$;#/$'''.
''' </summary>
''' <param name="nType">Join nature.</param>.
''' <param name="leftTable">First table.</param>.
''' <param name="rightTable">Second table.</param> ''' <param name="rightTable">Second table.
''' <param name="columns">String that specifies the structure of the rows returned.</param>
''' <param name="matchKeys">Priority and preference keys to be matched.</param>
''' <param name="predicate">Condition that must be met when selecting rows.</param>
''' <param name="headers">Indicates if the tables have headers.</param>
''' All indexes must be given IN BASE 1.
Public Function tJoin(ByVal nType As JoinType, _
                    ByRef leftTable As CSVArrayList, _
                    ByRef rightTable As CSVArrayList, _
                    Columns As String, _
                    matchKeys As String, _
                    Optional predicate As String = vbNullString, _
                    Optional Headers As Boolean = True) As CSVArrayList

    Dim JoinHelper As CSVArrayList

    Select Case nType
        Case JoinType.JT_LeftJoin
            Set tJoin = JoinHelper.LeftJoin(leftTable, rightTable, _
                                            Columns, matchKeys, predicate, _
                                            Headers)
        Case JoinType.JT_RightJoin
            Set tJoin = JoinHelper.RightJoin(leftTable, rightTable, _
                                            Columns, matchKeys, predicate, _
                                            Headers)
        Case Else
            Set tJoin = JoinHelper.InnerJoin(leftTable, rightTable, _
                                            Columns, matchKeys, predicate, _
                                            Headers)
    End Select
End Function
Private Function JoinKeys(ByRef tList As CSVArrayList, ByRef keys As Variant, index As Long) As String
    If IsArray(keys) Then
        Dim i As Long
        Dim tLLB As Long
        Dim tLUB As Long
        Dim tmpResult() As String

        tLLB = LBound(keys)
        tLUB = UBound(keys)
        ReDim tmpResult(tLLB To tLUB)
        For i = tLLB To tLUB
            tmpResult(i) = tList(index)(Abs(keys(i)) - 1)
        Next i
        JoinKeys = Join$(tmpResult, "-")
    Else
        JoinKeys = tList(index)(Abs(keys) - 1)
    End If
End Function
Public Function CSVsubsetSplit(filePath As String, _
                                Optional subsetColumns As Variant = 1, _
                                Optional Headers As Boolean = True, _
                                Optional repeatHeaders As Boolean = True, _
                                Optional streamSize As Long = 20, _
                                Optional oConfig As CSVparserConfig = Nothing) As Collection
    Dim CreatedFiles As Collection
    Dim CSVhead As Variant
    Dim CSVreader As CSVinterface
    Dim CSVstream As CSVTextStream
    Dim CSVwriter As CSVinterface
    Dim curItems As CSVArrayList
    Dim ExportSubSet As CSVArrayList
    Dim FExtension As String
    Dim fileCreatedFlag As Boolean
    Dim FileNameAndExtension() As String
    Dim fileOnPathFlag As Boolean
    Dim FName As String
    Dim fRootPath As String
    Dim i As Long
    Dim OutputPath As String
    Dim readerConf As CSVparserConfig
    Dim tmpVar As Variant
    Dim writterConf As CSVparserConfig

    Set CreatedFiles = New Collection
    Set CSVreader = New CSVinterface
    If Not oConfig Is Nothing Then
        Set CSVreader.parseConfig = oConfig.CopyConfig
    Else
        Set CSVreader.parseConfig = Me.parseConfig.CopyConfig
    End If
    Set CSVwriter = New CSVinterface
    Set readerConf = CSVreader.parseConfig
    Set ExportSubSet = New CSVArrayList
    Set CSVstream = New CSVTextStream

    On Error GoTo CSVsubsetSplit_ErrHandler
    If filePath <> vbNullString Then
        '@----------------------------------------------------
        ' Check file
        If FileExists(filePath) Then
            If FileLen(filePath) = 0 Then
                GoTo CSVsubsetSplit_EmptyCSVfile
            End If
        Else
            GoTo CSVsubsetSplit_MissingCSVfile
        End If
        fRootPath = GetFilePath(filePath)
        FileNameAndExtension() = GetFileNameAndExtension(filePath)
        FName = FileNameAndExtension(0)
        FExtension = FileNameAndExtension(1)
        fileCreatedFlag = False
        DoEvents
        '@----------------------------------------------------
        'Sequential reading from file
        With CSVstream
            .endStreamOnLineBreak = True
            .unifiedLFOutput = readerConf.multiEndOfLineCSV
            .OpenStream filePath
            .utf8EncodedFile = readerConf.utf8EncodedFile
            .bufferSize = streamSize
        End With
        '@----------------------------------------------------
        'Parser config
        readerConf.path = filePath
        With readerConf
            Set .dialect = CSVreader.SniffDelimiters(readerConf)
        End With
        Set CSVreader.parseConfig = readerConf.CopyConfig
        '@----------------------------------------------------
        'Parse string
        With readerConf
            .Headers = Headers
            .delimitersGuessing = True
            If Headers Then
                .endingRecord = 1
                CSVhead = CSVreader.ImportFromCSV(readerConf)(0)
                .endingRecord = 0
            End If
        End With
        OutputPath = fRootPath & FName & "-WorkDir\"
        '@----------------------------------------------------
        'Check directory
        If LenB(Dir(OutputPath, vbDirectory)) = 0 Then
            MkDir OutputPath
        End If
        Set writterConf = readerConf.CopyConfig
        Do While Not CSVstream.atEndOfStream
            CSVstream.ReadText 'Read next CSV data chunk
            RaiseEvent AfterLoadStream
            If Headers Then
                If fileCreatedFlag Then
                    readerConf.Headers = False
                End If
            End If
            Set curItems = CSVreader.ImportFromCSVString(CSVstream.bufferString, readerConf).items
            If Not fileCreatedFlag Then
                If Headers Then
                    i = 1
                Else
                    i = 0
                End If
            Else
                i = 0
            End If
            curItems.Sort i + 1, sortingKeys:=subsetColumns
            RaiseEvent AfterSort
            tmpVar = JoinKeys(curItems, subsetColumns, i)
            writterConf.path = ReplaceIllegalCharacters(OutputPath & CStr(tmpVar) & "." & FExtension, "_")
            If Headers Then
                '@----------------------------------------------------
                'Check directory
                fileOnPathFlag = FileExists(writterConf.path)
                If Not fileOnPathFlag Then
                    If repeatHeaders Then
                        ExportSubSet.Add CSVhead
                        RaiseEvent AfterImportRecord
                    End If
                End If
            End If
            For i = i To curItems.count - 1
                If tmpVar = JoinKeys(curItems, subsetColumns, i) Then
                    ExportSubSet.Add curItems(i) 'Append data
                    RaiseEvent AfterImportRecord
                Else
                    On Error Resume Next
                    CreatedFiles.Add writterConf.path, CStr(tmpVar)
                    ExportSubSet.ShrinkBuffer
                    CSVwriter.ExportToCSV ExportSubSet, writterConf, False, False
                    ExportSubSet.Reinitialize ExportSubSet.count
                    tmpVar = JoinKeys(curItems, subsetColumns, i)
                    writterConf.path = OutputPath & CStr(tmpVar) & "." & FExtension
                    If Headers Then
                        fileOnPathFlag = FileExists(writterConf.path)
                        If Not fileOnPathFlag Then
                            If repeatHeaders Then
                                ExportSubSet.Add CSVhead
                            End If
                        End If
                    End If
                    ExportSubSet.Add curItems(i) 'Append data
                    RaiseEvent AfterImportRecord
                End If
            Next i
            '@----------------------------------------------------
            'Export remnant data
            CreatedFiles.Add writterConf.path, CStr(tmpVar)
            ExportSubSet.ShrinkBuffer
            CSVwriter.ExportToCSV ExportSubSet, writterConf, False, False
            fileCreatedFlag = True
            ExportSubSet.Reinitialize ExportSubSet.count
        Loop
        Set CSVsubsetSplit = CreatedFiles
        Set CreatedFiles = Nothing
        Set CSVreader = Nothing
        Set CSVstream = Nothing
        Set CSVwriter = Nothing
        Set curItems = Nothing
        Set ExportSubSet = Nothing
        Set readerConf = Nothing
        Set writterConf = Nothing
        tmpVar = Null
        ResetAlerts
    Else
        GoTo CSVsubsetSplit_MissingCSVfile
    End If
    Exit Function
CSVsubsetSplit_ErrHandler:
    Set CreatedFiles = Nothing
    Set CSVreader = Nothing
    Set CSVstream = Nothing
    Set CSVwriter = Nothing
    Set curItems = Nothing
    Set ExportSubSet = Nothing
    Set readerConf = Nothing
    Set writterConf = Nothing
    Exit Function
CSVsubsetSplit_EmptyCSVfile:
    P_ERROR_DESC = "[CSV file subset]: The specified CSV file is empty. No subset can be processed."
    P_ERROR_NUMBER = vbObjectError + 9013
    P_ERROR_SOURCE = "CSVinterface"
    GoTo CSVsubsetSplit_ErrHandler
CSVsubsetSplit_MissingCSVfile:
    P_ERROR_DESC = "[CSV file subset]: The given path name is an empty string or the specified CSV file does not exist in the supplied path."
    P_ERROR_NUMBER = vbObjectError + 9014
    P_ERROR_SOURCE = "CSVinterface"
    GoTo CSVsubsetSplit_ErrHandler
End Function
Public Function Dedupe(keys As String) As CSVArrayList
    On Error GoTo ErrHandler_Dedupe

    Set Dedupe = P_CSV_DATA.Dedupe(keys, parseConfig.Headers)
    Exit Function
ErrHandler_Dedupe:
    Set Dedupe = Nothing
    P_ERROR_DESC = "[CSV Dedupe]: " & err.Description
    P_ERROR_NUMBER = err.Number
    P_ERROR_SOURCE = "CSVinterface"
End Function
Public Sub DumpToArray(OutPutArray() As Variant, _
                        Optional ByRef DataSource As CSVArrayList = Nothing)
    If DataSource Is Nothing Then
        Set DataSource = P_CSV_DATA
    End If
    If Not DataSource Is Nothing Then
        Dim tmpOutputArray() As Variant
        tmpOutputArray = DataSource.items
        DataSource.JaggedToTwoDimArray tmpOutputArray, OutPutArray
    End If
End Sub
Public Sub DumpToJaggedArray(OutPutArray() As Variant, _
                                Optional ByRef DataSource As CSVArrayList = Nothing)
    If DataSource Is Nothing Then
        Set DataSource = P_CSV_DATA
    End If
    If Not DataSource Is Nothing Then
        OutPutArray() = DataSource.items
    End If
End Sub
Public Function DumpToSheet(Optional WBookName As String, _
                        Optional SheetName As String, _
                        Optional rngName As String = "A1", _
                        Optional ByRef DataSource As CSVArrayList = Nothing, _
                        Optional BlockAutoFormat As Boolean = True, _
                        Optional DrawCellBorders As Boolean = False) As String
    On Error Resume Next
    Dim colNumber As Long

    If DataSource Is Nothing Then
        Set DataSource = P_CSV_DATA
        colNumber = P_VECTORS_MAX_BOUND
    Else
        colNumber = UBound(DataSource(0))
    End If
    If Not DataSource Is Nothing Then
        Dim WBook As Workbook
        Dim outputSheet As Worksheet
        Dim OutputRange As Range

        EnableOptimization
        '@------------------------------------------------------
        'Set the target Workbook
        If WBookName = vbNullString Then
            Set WBook = ThisWorkbook
        ElseIf Not IsWorkbookOpen(WBookName) Then
            Set WBook = Workbooks.Add
        Else
            Set WBook = Workbooks(WBookName)
        End If
        '@------------------------------------------------------
        'Set the target Worksheet
        If IsSheetInWorkbook(SheetName, WBook) Then
            Set outputSheet = WBook.Sheets(SheetName)
        Else
            Set outputSheet = WBook.Sheets.Add
            outputSheet.name = SheetName
        End If
        '@------------------------------------------------------
        'Dump the data
        Dim tmpOutputArray() As Variant
        Dim UB1 As Long
        Dim OutPutArr() As Variant

        UB1 = DataSource.count - 1
        tmpOutputArray() = DataSource.items
        DataSource.JaggedToTwoDimArray tmpOutputArray, OutPutArr

        '@------------------------------------------------------
        'Set the target Range
        Set OutputRange = outputSheet.Range(rngName) _
                            .Resize _
                            ( _
                            UB1 + 1, _
                            colNumber + 1 _
                            )
        With OutputRange
            If BlockAutoFormat Then
                .NumberFormat = "@"
            End If
            .value2 = OutPutArr
            If DrawCellBorders Then
                'Left Borders
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).ColorIndex = 0
                .Borders(xlEdgeLeft).TintAndShade = 0
                .Borders(xlEdgeLeft).Weight = xlThin
                'Top Borders
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).ColorIndex = 0
                .Borders(xlEdgeTop).TintAndShade = 0
                .Borders(xlEdgeTop).Weight = xlThin
                'Bottom Borders
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).ColorIndex = 0
                .Borders(xlEdgeBottom).TintAndShade = 0
                .Borders(xlEdgeBottom).Weight = xlThin
                'Right Borders
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).ColorIndex = 0
                .Borders(xlEdgeRight).TintAndShade = 0
                .Borders(xlEdgeRight).Weight = xlThin
                'Inside Vertical Borders
                .Borders(xlInsideVertical).LineStyle = xlContinuous
                .Borders(xlInsideVertical).ColorIndex = 0
                .Borders(xlInsideVertical).TintAndShade = 0
                .Borders(xlInsideVertical).Weight = xlThin
                'Inside Horizontal Borders
                .Borders(xlInsideHorizontal).LineStyle = xlContinuous
                .Borders(xlInsideHorizontal).ColorIndex = 0
                .Borders(xlInsideHorizontal).TintAndShade = 0
                .Borders(xlInsideHorizontal).Weight = xlThin
            End If
        End With
        Erase tmpOutputArray
        Erase OutPutArr
        EnableOptimization False
        DumpToSheet = OutputRange.Address
    End If
End Function
Private Sub EnableOptimization(Optional Optimize As Boolean = True)
    If Optimize Then
        '@------------------------------------------------------
        'Optimize resource consumption
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
    Else
        '@------------------------------------------------------
        'Turn Off optimization
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
    End If
End Sub
Private Sub EscapeField(ByRef FieldString As String, _
                        ByRef EscapeCharacter As String, _
                        ByRef EscapeSeq As String, _
                        ByRef FDeliter As String, _
                        ByRef RDelimiter As String, _
                        Optional EnforcedQuotation As Boolean = False)

    Dim FieldDelimiterInField As Boolean
    Dim MultiLineField As Boolean
    Dim QuotesInField As Boolean
    Dim tmpFld() As String

    FieldDelimiterInField = InStrB(1, FieldString, FDeliter)
    MultiLineField = InStrB(1, FieldString, RDelimiter)
    QuotesInField = InStrB(1, FieldString, EscapeCharacter)
    If FieldDelimiterInField Or MultiLineField Or QuotesInField Then
        ReDim tmpFld(0 To 2)
        tmpFld(0) = EscapeCharacter
        tmpFld(2) = EscapeCharacter
        If QuotesInField Then
            tmpFld(1) = Join$(Split(FieldString, EscapeCharacter), EscapeSeq)
        Else
            tmpFld(1) = FieldString
        End If
        FieldString = Join$(tmpFld, vbNullString)
    Else
        If EnforcedQuotation Then
            ReDim tmpFld(0 To 2)
            tmpFld(0) = EscapeCharacter
            tmpFld(1) = FieldString
            tmpFld(2) = EscapeCharacter
            FieldString = Join$(tmpFld, vbNullString)
        End If
    End If
End Sub
Public Sub ExportToCSV(ByRef csvArray As Variant, _
                        Optional pConfig As CSVparserConfig = Nothing, _
                        Optional PassControlToOS As Boolean = True, _
                        Optional enableDelimitersSniffing As Boolean = True, _
                        Optional EnforcedQuotation As Boolean = False)
    On Error GoTo ErrHandler_ExportToCSV
    If PassControlToOS Then DoEvents 'Pass the control to the Operative System
    If IsArray(csvArray) Or LCase(TypeName(csvArray)) = "csvarraylist" Then
        Dim curConfig As CSVparserConfig
        '@----------------------------------------------------------------------------
        'Open a conection to the file and seek its end
        If pConfig Is Nothing Then
            Set curConfig = config
        Else
            Set curConfig = pConfig
        End If
        Dim CSVstream As CSVTextStream
        Dim EOLB As EndLineChar
        Set CSVstream = New CSVTextStream
        If FileExists(curConfig.path) Then
            If enableDelimitersSniffing Then
                Set curConfig.dialect = SniffDelimiters(curConfig)
            End If
            With CSVstream
                .utf8EncodedFile = curConfig.utf8EncodedFile
                .OpenStream curConfig.path
            End With
            If CSVstream.streamLength > 0 Then
                Select Case curConfig.dialect.recordsDelimiter
                    Case vbCrLf
                        CSVstream.SeekPointer CSVstream.streamLength - 1
                    Case Else
                        CSVstream.SeekPointer CSVstream.streamLength
                End Select
                '@----------------------------------------------------------------------------
                'Check line break at EOF
                CSVstream.ReadText
                RaiseEvent AfterLoadStream
                If CSVstream.bufferString <> curConfig.dialect.recordsDelimiter Then
                    Select Case CSVstream.bufferString
                        Case vbCrLf, vbCr, vbLf
                            'Not needed to write endline char
                        Case Else
                            Select Case curConfig.dialect.recordsDelimiter
                                Case vbCrLf
                                    EOLB = EndLineChar.CRLF
                                Case vbCr
                                    EOLB = EndLineChar.CR
                                Case Else
                                    EOLB = EndLineChar.LF
                            End Select
                            CSVstream.WriteBlankLines 1, EOLB
                    End Select
                End If
            End If
        Else
            CSVstream.OpenStream curConfig.path
        End If
        '@----------------------------------------------------------------------------
        'Export
        CSVstream.WriteText JoinRecordsFields(csvArray, curConfig, EnforcedQuotation)
        RaiseEvent AfterWriteStream
        Set CSVstream = Nothing
        If P_ERROR_NUMBER = 0& Then
            Close #FileHandled
            P_SUCCESSFUL_EXPORT = True
            P_ERROR_DESC = vbNullString
            P_ERROR_SOURCE = vbNullString
        Else
            GoTo ErrHandler_ExportToCSV
        End If
    Else
        P_SUCCESSFUL_EXPORT = False
        P_ERROR_DESC = "[CSV file Export]: The passed argument isn't an array or a CSVArrayList object."
        P_ERROR_NUMBER = vbObjectError + 9012
        P_ERROR_SOURCE = "CSVinterface"
    End If
    Exit Sub
ErrHandler_ExportToCSV:
    Close #FileHandled
    P_SUCCESSFUL_EXPORT = False
    P_ERROR_DESC = "[CSV file Export]: " & P_ERROR_DESC
End Sub
Private Function FileExists(ByVal filePath As String) As Boolean
    FileExists = CBool(LenB(Dir(filePath, vbHidden + vbNormal + vbSystem + vbReadOnly + vbArchive)))
End Function
''' <summary>
''' Filter the target CSV file using expression evaluation.
''' </summary>
''' <param name="Pattern">A valid string expression to evaluate when filtering records.</param>
''' <param name="FilePath">Full file path, including file extension.</param>
''' <param name="ExcludeFirstRecord">Used if the user wants to exclude the headers.</param>
''' <param name="Exclude"> Exclude records matching "Pattern" if set to "True".</param>
Public Function Filter(Pattern As String, Optional filePath As String = vbNullString, _
                            Optional ExcludeFirstRecord As Boolean = True, _
                            Optional Exclude As Boolean = False) As CSVArrayList

    If filePath <> vbNullString Then 'Filter CSV file
        Dim CSVparser As CSVinterface
        Dim StreamReader As CSVTextStream
        Dim StreamsCounter As Long

        Set CSVparser = New CSVinterface
        Set StreamReader = New CSVTextStream
        Set Filter = New CSVArrayList

        With CSVparser.parseConfig
            .delimitersGuessing = True
        End With
        With StreamReader
            .endStreamOnLineBreak = True
            .OpenStream filePath
            Do While Not .atEndOfStream
                .ReadText
                RaiseEvent AfterLoadStream
                StreamsCounter = StreamsCounter + 1
                CSVparser.ImportFromCSVString .bufferString, CSVparser.parseConfig
                Filter.Concat2 CSVparser.items.Filter(Pattern, 1 + Abs(ExcludeFirstRecord And StreamsCounter = 1), Exclude)
            Loop
        End With
    Else
        Set CSVparser = Me
        If CSVparser.importSuccess Then
            ExcludeFirstRecord = (CSVparser.parseConfig.Headers = True)
            Set Filter = P_CSV_DATA.Filter(Pattern, 1 + Abs(ExcludeFirstRecord), Exclude)
        End If
    End If
End Function

Public Function GetDataFromCSV(csvPathAndFilename As String) As String

    On Error GoTo ErrHandler_GetDataFromCSV
    FileHandled = FreeFile
    Open csvPathAndFilename For Binary As #FileHandled
    GetDataFromCSV = Space$(LOF(FileHandled))
    Get #FileHandled, , GetDataFromCSV
    Close #FileHandled
ErrHandler_GetDataFromCSV:
    P_ERROR_DESC = "[Get Data From CSV]: " & err.Description
    P_ERROR_NUMBER = err.Number
    P_ERROR_SOURCE = err.Source
End Function
Private Function GetFileNameAndExtension(path As String) As String()
    Dim tmpResult() As String
    tmpResult() = Split(Mid$(path, InStrRev(path, "\") + 1), ".")
    If UBound(tmpResult) > 1 Then
        Dim i As Single
        Dim tmpName As String

        tmpName = tmpResult(0)
        For i = 1 To UBound(tmpResult) - 1
            tmpName = tmpName & "." & tmpResult(i)
        Next i
        tmpResult(0) = tmpName
        tmpResult(1) = tmpResult(i)
        ReDim Preserve tmpResult(0 To 1)
    End If
    GetFileNameAndExtension = tmpResult
End Function
Private Function GetHeaderIndexes(ByRef tmpFilters() As Variant, _
                                ByRef header As CSVArrayList) As Long()
    Dim tmpResult() As Long
    Dim headerFieldsCount As Long
    Dim tmpColl As Collection
    Dim trqidx As Long, arridx As Long
    Dim matchFound As Boolean

    If Not header Is Nothing Then
        headerFieldsCount = UBound(header(0))
        Set tmpColl = New Collection
        For trqidx = LBound(tmpFilters) To UBound(tmpFilters)
            If IsNumeric(tmpFilters(trqidx)) Then
                On Error Resume Next
                tmpColl.Add CStr(CLng(tmpFilters(trqidx))), CStr(CLng(tmpFilters(trqidx)))
            Else
                arridx = 0
                Do
                    matchFound = (CStr(tmpFilters(trqidx)) = CStr(header(0)(arridx)))
                    If matchFound Then
                        On Error Resume Next
                        tmpColl.Add CStr(arridx + 1), CStr(arridx + 1)
                    End If
                    arridx = arridx + 1
                Loop While Not matchFound And arridx <= headerFieldsCount
            End If
        Next trqidx
        If tmpColl.count >= 1 Then
            ReDim tmpResult(0 To tmpColl.count - 1)
            For arridx = 0 To tmpColl.count - 1
                tmpResult(arridx) = CLng(tmpColl.item(arridx + 1))
            Next arridx
        End If
        err.Clear
        GetHeaderIndexes = tmpResult
    End If
End Function
Private Function GetFilePath(path As String)
    GetFilePath = Mid$(path, 1, InStrRev(path, "\"))
End Function
Public Function GetRecord() As CSVArrayList

    On Error GoTo GetRecord_Error_Handler
    If Not static_SeqReadStarted Then Exit Function
    '@----------------------------------------------------------------------------
    'Get CSV record using streams
    If Not static_ImportSwitch Then
        '@----------------------------------------------------------------------------
        'Start variables
        static_outputList.Reinitialize 1
        static_IsHeader = (static_Headers And static_RecordsCount = 0)
        If static_IgnoreLines Then
            '@----------------------------------------------------------------------------
            'Skip lines if needed
            StreamSkipUnwantedLines static_Index, static_CSVstream, static_tmpCSV, _
                                    static_MaxIndex, static_RecordDelimiter, _
                                    static_commToken, static_StreamWithQuotes, _
                                    static_StreamsCounter, static_IgnoreCommentLines, static_IgnoreEmptyLines
        End If
        static_EndOFStreamBuffer = (static_Index > static_MaxIndex)
        If static_EndOFStreamBuffer Then 'The file has no significant data
            Set GetRecord = Nothing
            Exit Function
        End If
        static_LenCurrentIndex = LenB(static_tmpCSV(static_Index))
        '@----------------------------------------------------------------------------
        'Restart pointers
        static_RecordEndReached = False
        static_FieldID = 0
        Do
            '@----------------------------------------------------------------------------
            'Process the current Token
            static_TokenEndReached = False
            static_TokenBeginningPos = static_OverStringPointer
            If static_StreamWithQuotes Then
                Select Case static_OverStringPointer
                    Case Is < static_LenCurrentIndex
                        static_ASCIIcharw = AscW(MidB$(static_tmpCSV(static_Index), static_OverStringPointer, 2))
                        '@----------------------------------------------------------------------------
                        'Ignore spaces & tabs
                        Select Case static_FDAscW
                            Case Is <> 9
                                Do While static_ASCIIcharw = 32 Or static_ASCIIcharw = 9
                                    static_OverStringPointer = static_OverStringPointer + 2&
                                    If static_OverStringPointer > static_LenCurrentIndex Then Exit Do
                                    static_ASCIIcharw = AscW(MidB$(static_tmpCSV(static_Index), static_OverStringPointer, 2))
                                Loop
                            Case Else
                                Do While static_ASCIIcharw = 32
                                    static_OverStringPointer = static_OverStringPointer + 2&
                                    If static_OverStringPointer > static_LenCurrentIndex Then Exit Do
                                    static_ASCIIcharw = AscW(MidB$(static_tmpCSV(static_Index), static_OverStringPointer, 2))
                                Loop
                        End Select
                        static_OpenedToken = (static_ASCIIcharw = static_EscapeAscW)
                    Case Else 'Avoid over runs
                        static_OpenedToken = False
                End Select
            Else
                static_OpenedToken = False
            End If
            If static_OpenedToken Then
                If static_OpenTokenStreamIdx = 0 Then
                    static_OpenTokenStreamIdx = static_StreamsCounter
                End If
                If static_OpenTokenArrIdx = 0 Then
                    static_OpenTokenArrIdx = static_Index
                End If
                '@----------------------------------------------------------------------------
                'Try to escape the Token
                static_EscapeChrCount = static_EscapeChrCount + 1&
                static_TokenBeginningPos = static_OverStringPointer + 2&
                static_SearchBeginningMark = static_TokenBeginningPos
                static_hpointer = InStrB(static_SearchBeginningMark, _
                                            static_tmpCSV(static_Index), static_QuoteChar) 'EOFld marker
                If static_hpointer Then static_EscapeChrCount = static_EscapeChrCount + 1&
                Do While (Not static_TokenEndReached And Not static_EndOF)
                    '@----------------------------------------------------------------------------
                    'Buffering if needed
                    If Not static_SplittedToken Then static_SplittedToken = (static_hpointer = 0)
                    Do
                        Do While static_hpointer = 0& 'Quote char may be on subsequent static_Index
                            Select Case static_SearchBeginningMark
                                Case 1 'Store the full Token from the current static_Index
                                    static_tmpToken.Add static_tmpCSV(static_Index)
                                Case Else 'Store the Right most Token portion from the current static_Index
                                    static_tmpToken.Add RightB$(static_tmpCSV(static_Index), static_LenCurrentIndex - static_SearchBeginningMark + 1)
                            End Select
                            If static_OpenTokenStartPos = 0 Then
                                static_OpenTokenStartPos = static_SearchBeginningMark
                            End If
                            static_SplittedToken = True
                            static_Index = static_Index + 1&
                            static_EndOFStreamBuffer = (static_Index > static_MaxIndex)
                            If Not static_EndOFStreamBuffer Then
                                static_LenCurrentIndex = LenB(static_tmpCSV(static_Index))
                                static_SearchBeginningMark = 1
                                static_hpointer = InStrB(static_SearchBeginningMark, static_tmpCSV(static_Index), static_QuoteChar) 'Find quote char
                                If static_hpointer Then static_EscapeChrCount = static_EscapeChrCount + 1&
                            Else
                                If Not static_StreamEnd Then
                                    '@----------------------------------------------------------------------------
                                    'Load new stream
                                    With static_CSVstream
                                        .ReadText
                                        static_StreamsCounter = static_StreamsCounter + 1
                                        RaiseEvent AfterLoadStream
                                        static_StreamWithQuotes = InStrB(1, .bufferString, static_QuoteChar)
                                        static_StreamEnd = .atEndOfStream
                                        If static_UseUnixEscapeSeq Then
                                            static_tmpCSV() = Split(UnixToStandardEscapeSeq(.bufferString, _
                                                                    static_UnixEscapeSeq, static_EscapeSequence), static_RecordDelimiter)
                                        Else
                                            static_tmpCSV() = Split(.bufferString, static_RecordDelimiter)
                                        End If
                                    End With
                                    static_Index = 0
                                    static_MaxIndex = UBound(static_tmpCSV)
                                    If Not static_StreamEnd Then
                                        If static_tmpCSV(static_MaxIndex) = vbNullString Then 'Streams lefts empty value when splited
                                            static_MaxIndex = static_MaxIndex - 1
                                        End If
                                    Else
                                        If static_IgnoreEmptyLines Then
                                            If static_tmpCSV(static_MaxIndex) = vbNullString Then 'Skip empty line
                                                static_MaxIndex = static_MaxIndex - 1
                                            End If
                                        End If
                                    End If
                                    static_EndOFStreamBuffer = False
                                    static_LenCurrentIndex = LenB(static_tmpCSV(static_Index))
                                    static_SearchBeginningMark = 1
                                    static_hpointer = InStrB(static_SearchBeginningMark, static_tmpCSV(static_Index), static_QuoteChar) 'Find quote char
                                    If static_hpointer Then static_EscapeChrCount = static_EscapeChrCount + 1&
                                Else
                                    static_hpointer = 0&
                                End If
                            End If
                            static_EndOF = (static_StreamEnd And static_EndOFStreamBuffer)
                            If static_hpointer = 0& And static_EndOF Then 'Missing quote char
                                P_ERROR_DESC = P_ERROR_DESC & "Missing some quote char. The parser retrieved the record in recovery mode. " & _
                                               "[Review the record #" & static_RecordsCount + 1 & ", field #" & static_RecordToken.count + 1 _
                                                 & " on the source CSV file]."
                                static_RecordEndReached = True
                                static_OverStringPointer = 1&
                                static_StreamsCounter = 0
                                With static_CSVstream
                                    .RestartPointer
                                    Do While static_StreamsCounter < static_OpenTokenStreamIdx
                                        .ReadText
                                        static_StreamsCounter = static_StreamsCounter + 1
                                        RaiseEvent AfterLoadStream
                                        static_StreamWithQuotes = InStrB(1, .bufferString, static_QuoteChar)
                                        static_StreamEnd = .atEndOfStream
                                    Loop
                                    If static_UseUnixEscapeSeq Then
                                        static_tmpCSV() = Split(UnixToStandardEscapeSeq(.bufferString, _
                                                                static_UnixEscapeSeq, static_EscapeSequence), static_RecordDelimiter)
                                    Else
                                        static_tmpCSV() = Split(.bufferString, static_RecordDelimiter)
                                    End If
                                End With
                                static_tmpToken.Clear
                                static_MaxIndex = UBound(static_tmpCSV)
                                static_LenCurrentIndex = LenB(static_tmpCSV(static_OpenTokenArrIdx))
                                If static_OpenTokenStartPos = 1 Then 'Save the entire record exclusively in recovery mode
                                    static_RecordToken.Add MidB$(static_tmpCSV(static_OpenTokenArrIdx), 3)
                                Else 'Save a portion of the record in recovery mode
                                    static_RecordToken.Add RightB$(static_tmpCSV(static_OpenTokenArrIdx), _
                                                            static_LenCurrentIndex - static_OpenTokenStartPos + 1)
                                End If
                                static_Index = static_OpenTokenArrIdx + 1
                                GoTo Restore_From_Missing_EscapeChar
                            End If
                        Loop
                        static_OverStringPointer = static_hpointer + 2&
                        static_NotEvenEscapeChrNumber = (static_EscapeChrCount And 1)
                        If static_OverStringPointer < static_LenCurrentIndex Then
                            static_ASCIIcharw = AscW(MidB$(static_tmpCSV(static_Index), static_OverStringPointer, 2))
                            '@----------------------------------------------------------------------------
                            'Ignore spaces & tabs
                            Select Case static_FDAscW
                                Case Is <> 9
                                    Do While static_ASCIIcharw = 32 Or static_ASCIIcharw = 9
                                        static_OverStringPointer = static_OverStringPointer + 2&
                                        If static_OverStringPointer > static_LenCurrentIndex Then Exit Do
                                        static_ASCIIcharw = AscW(MidB$(static_tmpCSV(static_Index), static_OverStringPointer, 2))
                                    Loop
                                Case Else
                                    Do While static_ASCIIcharw = 32
                                        static_OverStringPointer = static_OverStringPointer + 2&
                                        If static_OverStringPointer > static_LenCurrentIndex Then Exit Do
                                        static_ASCIIcharw = AscW(MidB$(static_tmpCSV(static_Index), static_OverStringPointer, 2))
                                    Loop
                            End Select
                            '@----------------------------------------------------------------------------
                            'Check if we are at the end of the field
                            If static_UseUnixEscapeSeq Then
                                static_EscapedWFS = (static_ASCIIcharw = static_FDAscW) And (Not static_NotEvenEscapeChrNumber) And _
                                                    (Not (AscW(MidB$(static_tmpCSV(static_Index), static_OverStringPointer - 2, 2)) _
                                                                                                = ASCW_CHR_BACKSLASH))
                            Else
                                static_EscapedWFS = (static_ASCIIcharw = static_FDAscW) And (Not static_NotEvenEscapeChrNumber)
                            End If
                        Else
                            static_EscapedWFS = False
                            static_EscapedBRS = Not static_NotEvenEscapeChrNumber
                        End If
                        static_IsWellEscapedField = (static_EscapedWFS Or static_EscapedBRS)
                        '@----------------------------------------------------------------------------
                        'Advance the search
                        If Not static_IsWellEscapedField Then
                            static_hpointer = InStrB(static_hpointer + 2&, static_tmpCSV(static_Index), static_QuoteChar) 'EOFld marker
                            If static_hpointer Then static_EscapeChrCount = static_EscapeChrCount + 1&
                        End If
                    Loop While Not static_IsWellEscapedField
                    If static_EscapedWFS Then 'Quote char succeeded by fields separator char
                        static_TokenEndingPos = static_hpointer - 1&
                        static_TokenEndReached = True
                        static_EscapeChrCount = 0&
                        '@----------------------------------------------------------------------------
                        'Save the Token
                        Select Case static_SplittedToken
                            Case False 'Just save Token
                                If RequestedField2(static_FieldID + 1&, static_MaxReqIndex) Then  'Filter fields
                                    Select Case static_dynamicType
                                        Case False
                                            static_RecordToken.Add UnescapeField(MidB$(static_tmpCSV(static_Index), static_TokenBeginningPos, _
                                                                                        static_TokenEndingPos - static_TokenBeginningPos + 1), _
                                                                                static_QuoteChar, static_EscapeSequence)
                                        Case Else
                                            static_RecordToken.Add TypeData(UnescapeField(MidB$(static_tmpCSV(static_Index), static_TokenBeginningPos, _
                                                                                            static_TokenEndingPos - static_TokenBeginningPos + 1), _
                                                                                static_QuoteChar, static_EscapeSequence), _
                                                                    static_FieldID, static_IsHeader)
                                    End Select
                                End If
                            Case Else 'The Token is stored on static_tmpToken object
                                static_tmpToken.Add LeftB$(static_tmpCSV(static_Index), static_TokenEndingPos)
                                If RequestedField2(static_FieldID + 1&, static_MaxReqIndex) Then 'Filter fields
                                    static_RecordToken.Add UnescapeField(Join$(static_tmpToken.items, static_RecordDelimiter), _
                                                                    static_QuoteChar, static_EscapeSequence) 'Save whole Token
                                    static_tmpToken.Clear 'Clear static_tmpToken content
                                End If
                        End Select
                        static_SplittedToken = False 'Restart Token status
                        static_OverStringPointer = static_OverStringPointer + 2& 'Advance over the field separator char
                        static_RecordEndReached = False
                    Else
                        If static_OverStringPointer >= static_LenCurrentIndex Then  'Quote char succeeded
                                                                                    'by records separator char
                            static_TokenEndingPos = static_hpointer - 1&
                            static_TokenEndReached = True
                            static_EscapeChrCount = 0&
                            static_RecordEndReached = True 'EOR
                            '@----------------------------------------------------------------------------
                            'Save the Token
                            Select Case static_SplittedToken
                                Case False 'Just save Token
                                    If RequestedField2(static_FieldID + 1&, static_MaxReqIndex) Then 'Filter fields
                                        Select Case static_dynamicType
                                            Case False
                                                static_RecordToken.Add UnescapeField(MidB$(static_tmpCSV(static_Index), static_TokenBeginningPos, _
                                                                                                static_TokenEndingPos - static_TokenBeginningPos + 1), _
                                                                                        static_QuoteChar, static_EscapeSequence)
                                            Case Else
                                                static_RecordToken.Add TypeData(UnescapeField(MidB$(static_tmpCSV(static_Index), static_TokenBeginningPos, _
                                                                                                static_TokenEndingPos - static_TokenBeginningPos + 1), _
                                                                                            static_QuoteChar, static_EscapeSequence), _
                                                                                static_FieldID, static_IsHeader)
                                        End Select
                                    End If
                                Case Else 'The Token is stored on static_tmpToken object
                                    static_tmpToken.Add LeftB$(static_tmpCSV(static_Index), static_OverStringPointer - 3)
                                    If RequestedField2(static_FieldID + 1&, static_MaxReqIndex) Then 'Filter fields
                                        static_RecordToken.Add UnescapeField(Join$(static_tmpToken.items, static_RecordDelimiter), _
                                                                            static_QuoteChar, static_EscapeSequence) 'Save whole Token
                                        static_tmpToken.Clear 'Clear static_tmpToken content
                                    End If
                            End Select
                            static_SplittedToken = False 'Restart Token status
                            static_Index = static_Index + 1&
                            static_EndOFStreamBuffer = (static_Index > static_MaxIndex)
                            static_EndOF = (static_StreamEnd And static_EndOFStreamBuffer)
                            If static_EndOFStreamBuffer And Not static_EndOF Then
                                '@----------------------------------------------
                                'Advance stream
                                With static_CSVstream
                                    .ReadText
                                    static_StreamsCounter = static_StreamsCounter + 1
                                    RaiseEvent AfterLoadStream
                                    static_StreamWithQuotes = InStrB(1, .bufferString, static_QuoteChar)
                                    static_StreamEnd = .atEndOfStream
                                    If static_UseUnixEscapeSeq Then
                                        static_tmpCSV() = Split(UnixToStandardEscapeSeq(.bufferString, _
                                                                static_UnixEscapeSeq, static_EscapeSequence), static_RecordDelimiter)
                                    Else
                                        static_tmpCSV() = Split(.bufferString, static_RecordDelimiter)
                                    End If
                                End With
                                static_Index = 0
                                static_MaxIndex = UBound(static_tmpCSV)
                                If Not static_StreamEnd Then
                                    If static_tmpCSV(static_MaxIndex) = vbNullString Then 'Streams lefts empty value when splited
                                        static_MaxIndex = static_MaxIndex - 1
                                    End If
                                Else
                                    If static_IgnoreEmptyLines Then
                                        If static_tmpCSV(static_MaxIndex) = vbNullString Then 'Skip empty line
                                            static_MaxIndex = static_MaxIndex - 1
                                        End If
                                    End If
                                End If
                                static_EndOFStreamBuffer = False
                            End If
                            static_OverStringPointer = 1&
                        Else
                            If static_SplittedToken Then 'Store the Left most Token portion from the current static_Index
                                static_tmpToken.Add LeftB$(static_tmpCSV(static_Index), static_OverStringPointer - 1)
                            End If
                            static_SearchBeginningMark = static_OverStringPointer
                            static_hpointer = InStrB(static_SearchBeginningMark, static_tmpCSV(static_Index), static_QuoteChar) 'Find quote char
                            If static_hpointer Then static_EscapeChrCount = static_EscapeChrCount + 1&
                            static_TokenEndReached = False
                        End If
                    End If
                Loop
Restore_From_Missing_EscapeChar:
                static_OpenTokenStreamIdx = 0
                static_OpenTokenStartPos = 0
            Else
                static_OpenTokenArrIdx = 0
                '@----------------------------------------------------------------------------
                'Try to mark the Starting and Ending position for the field
                static_hpointer = InStrB(static_OverStringPointer, static_tmpCSV(static_Index), static_FieldDelimiter)
                If static_UseUnixEscapeSeq Then
                    Do While static_hpointer
                        If static_hpointer > 2 Then
                            static_UnixEscpedSep = (AscW(MidB$(static_tmpCSV(static_Index), static_hpointer - 2, 2)) = ASCW_CHR_BACKSLASH)
                        Else
                            static_UnixEscpedSep = False
                        End If
                        If static_UnixEscpedSep Then
                            static_hpointer = InStrB(static_hpointer + 1, static_tmpCSV(static_Index), static_FieldDelimiter)
                        Else
                            Exit Do
                        End If
                    Loop
                End If
                If static_hpointer Then 'Found fields delimiter char
                    static_SplittedToken = False
                    static_TokenEndReached = True
                    static_TokenEndingPos = static_hpointer - 1&
                    static_OverStringPointer = static_hpointer + 2&
                    static_SearchBeginningMark = static_OverStringPointer
                    static_RecordEndReached = False
                    '@----------------------------------------------------------------------------
                    'Save the Token
                    If RequestedField2(static_FieldID + 1&, static_MaxReqIndex) Then 'Filter fields
                        Select Case static_dynamicType
                            Case False
                                If static_UseUnixEscapeSeq Then
                                    static_RecordToken.Add UnescapeField(UnescapeField(MidB$(static_tmpCSV(static_Index), static_TokenBeginningPos, _
                                                                            static_TokenEndingPos - static_TokenBeginningPos + 1), _
                                                                        static_FieldDelimiter, static_EscapeSequence_Delim), _
                                                                    static_QuoteChar, static_EscapeSequence)
                                Else
                                    static_RecordToken.Add MidB$(static_tmpCSV(static_Index), static_TokenBeginningPos, _
                                                                static_TokenEndingPos - static_TokenBeginningPos + 1)
                                End If
                            Case Else
                                If static_UseUnixEscapeSeq Then
                                    static_RecordToken.Add TypeData(UnescapeField(UnescapeField(MidB$(static_tmpCSV(static_Index), static_TokenBeginningPos, _
                                                                                    static_TokenEndingPos - static_TokenBeginningPos + 1), _
                                                                            static_FieldDelimiter, static_EscapeSequence_Delim), _
                                                                        static_QuoteChar, static_EscapeSequence), _
                                                                    static_FieldID, static_IsHeader)
                                Else
                                    static_RecordToken.Add TypeData(MidB$(static_tmpCSV(static_Index), static_TokenBeginningPos, _
                                                                        static_TokenEndingPos - static_TokenBeginningPos + 1), _
                                                                static_FieldID, static_IsHeader)
                                End If
                        End Select
                    End If
                Else 'Missing fields delimiter char
                    If static_UseUnixEscapeSeq Then
                        Do While Not static_EndOF And (AscW(RightB$(static_tmpCSV(static_Index), 2)) = ASCW_CHR_BACKSLASH) 'Unix escaped Record delimiter
                            static_SplittedToken = True
                            Select Case static_SearchBeginningMark
                                Case 1 'Store the full Token from the current Index
                                    static_tmpToken.Add MidB$(static_tmpCSV(static_Index), 1, static_LenCurrentIndex - 1)
                                Case Else 'Store the Right most Token portion from the current Index
                                    static_tmpToken.Add MidB$(static_tmpCSV(static_Index), static_TokenBeginningPos, static_LenCurrentIndex - static_TokenBeginningPos - 1)
                            End Select
                            static_tmpCSV(static_Index) = vbNullString
                            static_Index = static_Index + 1&
                            static_OverStringPointer = 1&
                            static_EndOFStreamBuffer = (static_Index > static_MaxIndex)
                            If Not static_EndOFStreamBuffer Then
                                static_LenCurrentIndex = LenB(static_tmpCSV(static_Index))
                                static_SearchBeginningMark = 1
                                '@----------------------------------------------------------------------------
                                'Try to mark the Starting and Ending position for the field
                                static_hpointer = InStrB(static_OverStringPointer, static_tmpCSV(static_Index), static_FieldDelimiter)
                                Do While static_hpointer
                                    If static_hpointer > 2 Then
                                        static_UnixEscpedSep = (AscW(MidB$(static_tmpCSV(static_Index), static_hpointer - 2, 2)) = ASCW_CHR_BACKSLASH)
                                    Else
                                        static_UnixEscpedSep = False
                                    End If
                                    If static_UnixEscpedSep Then
                                        static_hpointer = InStrB(static_hpointer + 1, static_tmpCSV(static_Index), static_FieldDelimiter)
                                    Else
                                        Exit Do
                                    End If
                                Loop
                            Else
                                If Not static_StreamEnd Then
                                    '@----------------------------------------------------------------------------
                                    'Load new stream
                                    With static_CSVstream
                                        .ReadText
                                        static_StreamsCounter = static_StreamsCounter + 1
                                        RaiseEvent AfterLoadStream
                                        static_StreamWithQuotes = InStrB(1, .bufferString, static_QuoteChar)
                                        static_StreamEnd = .atEndOfStream
                                        If static_UseUnixEscapeSeq Then
                                            static_tmpCSV() = Split(UnixToStandardEscapeSeq(.bufferString, _
                                                                    static_UnixEscapeSeq, static_EscapeSequence), static_RecordDelimiter)
                                        Else
                                            static_tmpCSV() = Split(.bufferString, static_RecordDelimiter)
                                        End If
                                    End With
                                    static_Index = 0
                                    static_MaxIndex = UBound(static_tmpCSV)
                                    If Not static_StreamEnd Then
                                        If static_tmpCSV(static_MaxIndex) = vbNullString Then 'Streams lefts empty value when splited
                                            static_MaxIndex = static_MaxIndex - 1
                                        End If
                                    Else
                                        If static_IgnoreEmptyLines Then
                                            If static_tmpCSV(static_MaxIndex) = vbNullString Then 'Skip empty line
                                                static_MaxIndex = static_MaxIndex - 1
                                            End If
                                        End If
                                    End If
                                    static_EndOFStreamBuffer = False
                                    static_LenCurrentIndex = LenB(static_tmpCSV(static_Index))
                                    static_SearchBeginningMark = 1
                                    static_OverStringPointer = static_SearchBeginningMark
                                    '@----------------------------------------------------------------------------
                                    'Try to mark the Starting and Ending position for the field
                                    static_hpointer = InStrB(static_OverStringPointer, static_tmpCSV(static_Index), static_FieldDelimiter)
                                    Do While static_hpointer
                                        If static_hpointer > 2 Then
                                            static_UnixEscpedSep = (AscW(MidB$(static_tmpCSV(static_Index), static_hpointer - 2, 2)) = ASCW_CHR_BACKSLASH)
                                        Else
                                            static_UnixEscpedSep = False
                                        End If
                                        If static_UnixEscpedSep Then
                                            static_hpointer = InStrB(static_hpointer + 1, static_tmpCSV(static_Index), static_FieldDelimiter)
                                        Else
                                            Exit Do
                                        End If
                                    Loop
                                Else
                                    static_hpointer = 0&
                                End If
                            End If
                            static_EndOF = (static_StreamEnd And static_EndOFStreamBuffer)
                            If static_hpointer Then 'Unescaped fields delimiter found
                                static_TokenEndReached = True
                                static_TokenEndingPos = static_hpointer - 1&
                                static_OverStringPointer = static_hpointer + 2&
                                static_SearchBeginningMark = static_OverStringPointer
                                static_RecordEndReached = False
                                static_tmpToken.Add LeftB$(static_tmpCSV(static_Index), static_TokenEndingPos) 'Store the Left most Token portion from the current static_Index
                                Exit Do
                            Else
                                If Not static_UnixEscpedSep Then 'End of record
                                    Select Case static_SearchBeginningMark
                                        Case 1 'Store the full Token from the current Index
                                            static_tmpToken.Add MidB$(static_tmpCSV(static_Index), 1, static_LenCurrentIndex)
                                        Case Else 'Store the Right most Token portion from the current Index
                                            static_tmpToken.Add MidB$(static_tmpCSV(static_Index), static_TokenBeginningPos, _
                                                                static_LenCurrentIndex - static_TokenBeginningPos)
                                    End Select
                                End If
                            End If
                        Loop
                    Else
                        static_hpointer = 0&
                    End If
                    If static_hpointer = 0 Then
                        static_TokenEndReached = True
                        static_RecordEndReached = True
                    End If
                    '@----------------------------------------------------------------------------
                    'Save the Token
                    Select Case static_SplittedToken
                        Case False 'Just save Token
                            If RequestedField2(static_FieldID + 1&, static_MaxReqIndex) Then 'Filter fields
                                Select Case static_dynamicType
                                    Case False
                                        If static_UseUnixEscapeSeq Then
                                            static_RecordToken.Add UnescapeField(UnescapeField(RightB$(static_tmpCSV(static_Index), static_LenCurrentIndex _
                                                                                    - static_TokenBeginningPos + 1), _
                                                                                static_FieldDelimiter, static_EscapeSequence_Delim), _
                                                                        static_QuoteChar, static_EscapeSequence)
                                        Else
                                            static_RecordToken.Add RightB$(static_tmpCSV(static_Index), static_LenCurrentIndex _
                                                                    - static_TokenBeginningPos + 1)
                                        End If
                                    Case Else
                                        If static_UseUnixEscapeSeq Then
                                            static_RecordToken.Add TypeData(UnescapeField(UnescapeField(RightB$(static_tmpCSV(static_Index), static_LenCurrentIndex _
                                                                                        - static_TokenBeginningPos + 1), _
                                                                                static_FieldDelimiter, static_EscapeSequence_Delim), _
                                                                            static_QuoteChar, static_EscapeSequence), _
                                                                        static_FieldID, static_IsHeader)
                                        Else
                                            static_RecordToken.Add TypeData(RightB$(static_tmpCSV(static_Index), static_LenCurrentIndex _
                                                                                - static_TokenBeginningPos + 1), _
                                                                        static_FieldID, static_IsHeader)
                                        End If
                                End Select
                            End If
                        Case Else
                            If RequestedField2(static_FieldID + 1&, static_MaxReqIndex) Then 'Filter fields
                                If static_UseUnixEscapeSeq Then
                                    static_RecordToken.Add UnescapeField(UnescapeField(Join$(static_tmpToken.items, static_RecordDelimiter), _
                                                                            static_FieldDelimiter, static_EscapeSequence_Delim), _
                                                                        static_QuoteChar, static_EscapeSequence) 'Save whole Token
                                Else
                                    static_RecordToken.Add Join$(static_tmpToken.items, static_RecordDelimiter)
                                End If
                            End If
                            static_tmpToken.Clear 'Clear tmpToken content
                            static_SplittedToken = False 'Restart Token status
                    End Select
                    If static_RecordEndReached Then
                        static_Index = static_Index + 1&
                        static_EndOFStreamBuffer = (static_Index > static_MaxIndex)
                        static_EndOF = (static_StreamEnd And static_EndOFStreamBuffer)
                        If static_EndOFStreamBuffer And Not static_EndOF Then
                            '@----------------------------------------------
                            'Advance stream
                            With static_CSVstream
                                .ReadText
                                static_StreamsCounter = static_StreamsCounter + 1
                                RaiseEvent AfterLoadStream
                                static_StreamWithQuotes = InStrB(1, .bufferString, static_QuoteChar)
                                static_StreamEnd = .atEndOfStream
                                If static_UseUnixEscapeSeq Then
                                    static_tmpCSV() = Split(UnixToStandardEscapeSeq(.bufferString, static_UnixEscapeSeq, _
                                                static_EscapeSequence), static_RecordDelimiter)
                                Else
                                    static_tmpCSV() = Split(.bufferString, static_RecordDelimiter)
                                End If
                            End With
                            static_Index = 0
                            static_MaxIndex = UBound(static_tmpCSV)
                            If Not static_StreamEnd Then
                                If static_tmpCSV(static_MaxIndex) = vbNullString Then 'Streams lefts empty value when splited
                                    static_MaxIndex = static_MaxIndex - 1
                                End If
                            Else
                                If static_IgnoreEmptyLines Then
                                    If static_tmpCSV(static_MaxIndex) = vbNullString Then 'Skip empty line
                                        static_MaxIndex = static_MaxIndex - 1
                                    End If
                                End If
                            End If
                            static_EndOFStreamBuffer = False
                        End If
                        static_OverStringPointer = 1&
                    End If
                End If
            End If
            static_FieldID = static_FieldID + 1&
        Loop While (Not static_RecordEndReached And Not static_EndOF)
        Select Case static_fieldsBound
            Case Is > 0
                If static_RecordToken.count - 1 > static_fieldsMaxBound Then
                    static_fieldsMaxBound = static_RecordToken.count - 1
                    static_RecordToken.ShrinkBuffer
                End If
            Case Else
                Select Case static_FieldID
                    Case Is > 0
                        static_fieldsBound = static_RecordToken.count - 1
                        static_RecordToken.ShrinkBuffer
                    Case Else
                        static_fieldsBound = static_FieldID
                End Select
                static_fieldsMaxBound = static_fieldsBound
        End Select
        static_RecordsCount = static_RecordsCount + 1&
        static_EscapedBRS = False
        '@----------------------------------------------------------------------------
        'Save record
        static_outputList.Add static_RecordToken.items
        static_RecordToken.Reinitialize static_fieldsBound + 1
        P_VARYING_LENGTHS = (static_fieldsBound <> static_fieldsMaxBound)
        If static_IgnoreLines Then
            '@----------------------------------------------------------------------------
            'Skip lines if needed
            StreamSkipUnwantedLines static_Index, static_CSVstream, static_tmpCSV, _
                                    static_MaxIndex, static_RecordDelimiter, _
                                    static_commToken, static_StreamWithQuotes, _
                                    static_StreamsCounter, static_IgnoreCommentLines, static_IgnoreEmptyLines
        End If
        static_EndOFStreamBuffer = (static_Index > static_MaxIndex)
        static_EndOF = (static_StreamEnd And static_EndOFStreamBuffer)
        If static_EndOFStreamBuffer And Not static_EndOF Then
            '@----------------------------------------------
            'Advance stream
            With static_CSVstream
                .ReadText
                static_StreamsCounter = static_StreamsCounter + 1
                RaiseEvent AfterLoadStream
                static_StreamWithQuotes = InStrB(1, .bufferString, static_QuoteChar)
                static_StreamEnd = .atEndOfStream
                If static_UseUnixEscapeSeq Then
                    static_tmpCSV() = Split(UnixToStandardEscapeSeq(.bufferString, _
                                            static_UnixEscapeSeq, static_EscapeSequence), static_RecordDelimiter)
                Else
                    static_tmpCSV() = Split(.bufferString, static_RecordDelimiter)
                End If
            End With
            static_Index = 0
            static_MaxIndex = UBound(static_tmpCSV)
            If Not static_StreamEnd Then
                If static_tmpCSV(static_MaxIndex) = vbNullString Then 'Streams lefts empty value when splited
                    static_MaxIndex = static_MaxIndex - 1
                End If
            Else
                If static_IgnoreEmptyLines Then
                    If static_tmpCSV(static_MaxIndex) = vbNullString Then 'Skip empty line
                        static_MaxIndex = static_MaxIndex - 1
                    End If
                End If
            End If
            static_EndOFStreamBuffer = False
        End If
        'static_ImportSwitch = static_EndOF And static_RecordEndReached
        static_ImportSwitch = static_EndOF And static_RecordEndReached
        P_VECTORS_REGULAR_BOUND = static_fieldsBound
        P_VECTORS_MAX_BOUND = static_fieldsMaxBound
        '@----------------------------------------------------------------------------
        'Check if the array need to be shrinked at the output stage
        static_outputList.ShrinkBuffer
        Set GetRecord = static_outputList
        P_SUCCESSFUL_IMPORT = True
    Else
        Set GetRecord = Nothing
        CloseSeqReader
    End If
    Exit Function
GetRecord_Error_Handler:
    Set GetRecord = Nothing
    P_SUCCESSFUL_IMPORT = False
    P_ERROR_NUMBER = err.Number
    P_ERROR_DESC = err.Description & " [Review the record #" & static_RecordsCount + 1 _
                                  & ", field #" & static_RecordToken.count + 1 _
                                   & " on the source CSV file]."
    P_ERROR_SOURCE = err.Source
    Exit Function
GetRecord_MissingEscapeChar:
    err.Raise Number:=vbObjectError + 9006, _
                    Source:="CSVinterface Class [GetRecord]", _
                    Description:="Missing some quote char. Check the data and try again."
    Resume GetRecord_Error_Handler
End Function
Public Function ImportFromCSV(configObj As CSVparserConfig, _
                            ParamArray FilterColumns() As Variant) As CSVinterface
    DoEvents 'Pass the control to the Operative System
    With configObj
        If FileExists(.path) Then
            If FileLen(.path) = 0 Then
                GoTo ImportFromCSV_EmptyFile
            End If
            If UBound(.dTypingTemplate) >= UBound(.dTypingLinks) Then
                Dim tmpRequested() As Variant
                Dim curConfig As CSVparserConfig
                Dim EmptyParam() As Variant
                ReDim EmptyParam(-1 To -1) 'Emule empty ParamArray

                If .delimitersGuessing Then
                    Set configObj.dialect = SniffDelimiters(configObj)
                End If
                tmpRequested() = FilterColumns
                If .Headers Then
                    '@--------------------------------------------------------------------------------
                    'Read the whole header
                    Set P_CSV_HEADER = New CSVArrayList
                    Set curConfig = .CopyConfig
                    curConfig.startingRecord = 1
                    curConfig.endingRecord = 1
                    StreamParseCSV curConfig, P_CSV_HEADER, EmptyParam
                End If
                If P_ERROR_NUMBER <> 0& Then GoTo ErrHandler_ImportFromCSV
                '@--------------------------------------------------------------------------------
                'Read the CSV data
                Set P_CSV_DATA = New CSVArrayList
                StreamParseCSV configObj, P_CSV_DATA, tmpRequested
                If P_ERROR_NUMBER = 0& Then
                    'Save pointers
                    P_SUCCESSFUL_IMPORT = True
                    P_ERROR_DESC = vbNullString
                    P_ERROR_SOURCE = vbNullString
                Else
                    GoTo ErrHandler_ImportFromCSV
                End If
            Else
                P_SUCCESSFUL_IMPORT = False
                P_ERROR_DESC = "[CSV file Import]: the config object has an invalid Dynamic Typing Template (DTT). " _
                                & "The number of Dynamic Typing Links (DTL) must be less or equal than the number of " _
                                & "Dynamic Typing Targets Fields (DTTF) defined."
                P_ERROR_NUMBER = vbObjectError + 9007
                P_ERROR_SOURCE = "CSVinterface"
            End If
        Else
            P_SUCCESSFUL_IMPORT = False
            P_ERROR_DESC = "[CSV file Import]: the config object is not linked to a CSV file. " _
                                & "Ensure set the path property to valid CSV before import data."
            P_ERROR_NUMBER = vbObjectError + 9008
            P_ERROR_SOURCE = "CSVinterface"
        End If
    End With
    Set ImportFromCSV = Me
    Exit Function
ErrHandler_ImportFromCSV:
    P_SUCCESSFUL_IMPORT = False
    P_ERROR_DESC = "[ImportFromCSV]: " & P_ERROR_DESC
    Exit Function
ImportFromCSV_EmptyFile:
    Set P_CSV_DATA = Nothing
    err.Raise Number:=vbObjectError + 9011, _
                    Source:="CSVinterface Class", _
                    Description:="The specified source CSV is empty. Please check and try again."
    Resume ErrHandler_ImportFromCSV
End Function
Public Function ImportFromCSVString(ByRef CSVstring As String, _
                                    configObj As CSVparserConfig, _
                                    ParamArray FilterColumns() As Variant) As CSVinterface
    With configObj
        If UBound(.dTypingTemplate) >= UBound(.dTypingLinks) Then
            If LenB(CSVstring) = 0 Then
                GoTo ImportFromCSV_EmptyString
            End If
            Dim tmpRequested() As Variant
            Dim curConfig As CSVparserConfig
            Dim EmptyParam() As Variant
            ReDim EmptyParam(-1 To -1) 'Emule empty ParamArray

            If .delimitersGuessing Then
                Set configObj.dialect = SniffDelimiters(configObj, CSVstring)
            End If
            tmpRequested() = FilterColumns
            If .Headers Then
                '@--------------------------------------------------------------------------------
                'Read the header
                Set P_CSV_HEADER = New CSVArrayList
                Set curConfig = .CopyConfig
                curConfig.startingRecord = 1
                curConfig.endingRecord = 1
                ParseCSVstring CSVstring, curConfig, P_CSV_HEADER, EmptyParam
            End If
            If P_ERROR_NUMBER <> 0& Then GoTo ErrHandler_ImportFromString
            '@--------------------------------------------------------------------------------
            'Read the CSV data
            Set P_CSV_DATA = New CSVArrayList
            ParseCSVstring CSVstring, configObj, P_CSV_DATA, tmpRequested
            '@----------------------------------------------------------------------------
            If P_ERROR_NUMBER = 0& Then
                'Save pointers
                P_SUCCESSFUL_IMPORT = True
                P_ERROR_DESC = vbNullString
                P_ERROR_SOURCE = vbNullString
            Else
                GoTo ErrHandler_ImportFromString
            End If
        Else
            P_SUCCESSFUL_IMPORT = False
            P_ERROR_DESC = "[CSV String Import]: the config object has an invalid Dynamic Typing Template (DTT). " _
                            & "The number of Dynamic Typing Links (DTL) must be less or equal than the number of " _
                            & "Dynamic Typing Targets Fields(DTTF) defined."
            P_ERROR_NUMBER = vbObjectError + 9007
            P_ERROR_SOURCE = "CSVinterface"
        End If
    End With
    Set ImportFromCSVString = Me
    Exit Function
ErrHandler_ImportFromString:
    P_SUCCESSFUL_IMPORT = False
    P_ERROR_DESC = "[ImportFromCSVstring]: " & P_ERROR_DESC
    Exit Function
ImportFromCSV_EmptyString:
    Set P_CSV_DATA = Nothing
    err.Raise Number:=vbObjectError + 9011, _
                    Source:="CSVinterface Class", _
                    Description:="The specified source String is empty. Please check and try again."
    Resume ErrHandler_ImportFromString
End Function
Public Function InsertField(aIndex As Long, _
                            Optional FieldName As String = vbNullString, _
                            Optional Formula As String = vbNullString) As CSVinterface
    On Error GoTo ErrHandler_InsertField
    If P_SUCCESSFUL_IMPORT And Not P_VARYING_LENGTHS Then
        Select Case aIndex
            Case 0 To P_VECTORS_REGULAR_BOUND + 1
                P_CSV_DATA.InsertField aIndex, FieldName, parseConfig.Headers, Formula
            Case Else
                GoTo OutOfBounds_InsertField
        End Select
    Else
        P_ERROR_DESC = "[CSV Field Insert]: Cannot insert a field in the current instance." _
                        & " This is because there is no imported data or the records do not " _
                        & "have the same number of fields."
        P_ERROR_NUMBER = vbObjectError + 9015
        P_ERROR_SOURCE = "CSVinterface"
    End If
    P_VECTORS_REGULAR_BOUND = P_VECTORS_REGULAR_BOUND + 1
    P_VECTORS_MAX_BOUND = P_VECTORS_MAX_BOUND + 1
    Set InsertField = Me
    Exit Function
ErrHandler_InsertField:
    P_ERROR_DESC = "[CSV Field Insert]: " & err.Description
    P_ERROR_NUMBER = err.Number
    P_ERROR_SOURCE = "CSVinterface"
    Exit Function
OutOfBounds_InsertField:
    err.Raise Number:=vbObjectError + 9016, _
                    Source:="CSVinterface Class", _
                    Description:="The specified index is out of bounds. Please check and try again."
    Resume ErrHandler_InsertField
End Function
Public Function InsertRecord(aIndex As Long) As CSVinterface
    On Error GoTo ErrHandler_InsertRecord
    If P_SUCCESSFUL_IMPORT And Not P_VARYING_LENGTHS Then
        Select Case aIndex
            Case 0 To P_CSV_DATA.count
                Dim cpRecord() As Variant
                Dim cIndex As Long

                cIndex = aIndex
                If aIndex = 0 Then
                    If parseConfig.Headers Then 'Insert on top only if there is no headers
                        cIndex = 1
                    End If
                End If
                ReDim cpRecord(0 To P_VECTORS_REGULAR_BOUND)
                With P_CSV_DATA
                    .Insert cIndex, cpRecord
                End With
            Case Else
                GoTo OutOfBounds_InsertRecord
        End Select
    Else
        P_ERROR_DESC = "[CSV Record Insert]: Cannot insert the record in the current instance." _
                        & " This is because there is no imported data or the records do not " _
                        & "have the same number of fields."
        P_ERROR_NUMBER = vbObjectError + 9018
        P_ERROR_SOURCE = "CSVinterface"
    End If

    Set InsertRecord = Me
    Exit Function
ErrHandler_InsertRecord:
    P_ERROR_DESC = "[CSV Record Insert]: " & err.Description
    P_ERROR_NUMBER = err.Number
    P_ERROR_SOURCE = "CSVinterface"
    Exit Function
OutOfBounds_InsertRecord:
    err.Raise Number:=vbObjectError + 9016, _
                    Source:="CSVinterface Class", _
                    Description:="The specified index is out of bounds. Please check and try again."
    Resume ErrHandler_InsertRecord
End Function
Private Function IsSheetInWorkbook(SheetName As String, WBook As Workbook) As Boolean
    With WBook
        On Error Resume Next
        IsSheetInWorkbook = (.Sheets(SheetName).name = SheetName)
        On Error GoTo 0
    End With
End Function
Private Function IsWorkbookOpen(WBookName As String) As Boolean
    Dim WBook As Workbook, BookMatching As Boolean

    On Error Resume Next
    For Each WBook In Workbooks
        BookMatching = (WBook.name = WBookName)
        If BookMatching Then Exit For
    Next
    IsWorkbookOpen = BookMatching
End Function
Private Function JoinRecordsFields(ByRef RecordsArray As Variant, _
                                    pConfig As CSVparserConfig, _
                                    Optional EnforcedQuotation As Boolean = False) As String
    Dim Buffer() As String
    Dim CoerceChr As String
    Dim ConcatenatedArray() As String
    Dim QuoteT As QuoteTokens
    Dim FldDelimiter As String
    Dim iLCounter As Long, jLCounter As Long
    Dim JaggedRecordsArray() As Variant
    Dim LB1 As Long, UB1 As Long
    Dim LB2 As Long, UB2 As Long
    Dim recDelimiter As String
    Dim ArrayHelper As CSVArrayList

    On Error GoTo JoinRecordsFields_Error
    Set ArrayHelper = New CSVArrayList
    '@----------------------------------------------------------------------------
    'Expand configuration object
    QuoteT = pConfig.dialect.quoteToken
    FldDelimiter = pConfig.dialect.fieldsDelimiter
    recDelimiter = pConfig.dialect.recordsDelimiter
    '@----------------------------------------------------------------------------
    'Set coerce char
    Select Case QuoteT
        Case 1
            CoerceChr = CHR_APOSTROPHE
        Case 2
            CoerceChr = CHR_DOUBLE_QUOTES
        Case Else
            CoerceChr = CHR_TILDE
    End Select
    '@----------------------------------------------------------------------------
    'Set the escape sequences
    UseUnixEscapeSeq = (pConfig.dialect.escapeMode = unix)
    UnixEscapeSeq = CHR_BACKSLASH & CoerceChr
    UnixEscapeSeqDel = CHR_BACKSLASH & FldDelimiter
    UnixEscapeSeqRec = CHR_BACKSLASH & recDelimiter
    EscapeSequence = CoerceChr & CoerceChr
    Dim tmpBuffer As Variant
    If LCase(TypeName(RecordsArray)) = "csvarraylist" Then
        JaggedRecordsArray() = RecordsArray.items
        '@----------------------------------------------------------------------------
        'Set array sizes
        LB1 = 0
        UB1 = RecordsArray.count - 1
        ReDim ConcatenatedArray(LB1 To UB1)
        '@----------------------------------------------------------------------------
        'Concatenate fields
        For iLCounter = LB1 To UB1
            tmpBuffer = JaggedRecordsArray(iLCounter)
            LB2 = LBound(tmpBuffer)
            UB2 = UBound(tmpBuffer)
            ReDim Buffer(LB2 To UB2)
            For jLCounter = LB2 To UB2
                Buffer(jLCounter) = tmpBuffer(jLCounter)
                If UseUnixEscapeSeq Then
                    UnixEscapeField Buffer(jLCounter), CoerceChr, FldDelimiter, _
                                    recDelimiter
                Else
                    EscapeField Buffer(jLCounter), CoerceChr, EscapeSequence, _
                                FldDelimiter, recDelimiter, EnforcedQuotation
                End If
            Next jLCounter
            ConcatenatedArray(iLCounter) = Join$(Buffer, FldDelimiter)
            RaiseEvent AfterExportRecord
        Next iLCounter
        '@----------------------------------------------------------------------------
        'Concatenate records
        JoinRecordsFields = Join$(ConcatenatedArray, recDelimiter)
        '@----------------------------------------------------------------------------
        'Free the memory
        Erase ConcatenatedArray
        Erase Buffer
    ElseIf ArrayHelper.MultiDimensional(RecordsArray) Then
        LB1 = LBound(RecordsArray, 1)
        LB2 = LBound(RecordsArray, 2)
        UB1 = UBound(RecordsArray, 1)
        UB2 = UBound(RecordsArray, 2)
        '@----------------------------------------------------------------------------
        'Set array sizes
        ReDim ConcatenatedArray(LB1 To UB1)
        ReDim Buffer(LB2 To UB2)
        '@----------------------------------------------------------------------------
        'Concatenate fields
        For iLCounter = LB1 To UB1
            For jLCounter = LB2 To UB2
                Buffer(jLCounter) = RecordsArray(iLCounter, jLCounter)
                If UseUnixEscapeSeq Then
                    UnixEscapeField Buffer(jLCounter), CoerceChr, FldDelimiter, _
                                    recDelimiter
                Else
                    EscapeField Buffer(jLCounter), CoerceChr, EscapeSequence, _
                                FldDelimiter, recDelimiter, EnforcedQuotation
                End If
            Next jLCounter
            ConcatenatedArray(iLCounter) = Join$(Buffer, FldDelimiter)
            RaiseEvent AfterExportRecord
        Next iLCounter
        '@----------------------------------------------------------------------------
        'Concatenate records
        JoinRecordsFields = Join$(ConcatenatedArray, recDelimiter)
        '@----------------------------------------------------------------------------
        'Free the memory
        Erase ConcatenatedArray
        Erase Buffer
    Else
        '@----------------------------------------------------------------------------
        'Check for jagged arrays
        LB1 = LBound(RecordsArray, 1)
        UB1 = UBound(RecordsArray, 1)
        If Not ArrayHelper.isJaggedArray(RecordsArray) Then
            '@----------------------------------------------------------------------------
            'Set array sizes
            ReDim ConcatenatedArray(LB1 To UB1)
            ReDim Buffer(0)
            '@----------------------------------------------------------------------------
            'Concatenate fields
            For iLCounter = LB1 To UB1
                Buffer(0) = RecordsArray(iLCounter)
                If UseUnixEscapeSeq Then
                    UnixEscapeField Buffer(jLCounter), CoerceChr, FldDelimiter, _
                                    recDelimiter
                Else
                    EscapeField Buffer(0), CoerceChr, EscapeSequence, _
                                FldDelimiter, recDelimiter, EnforcedQuotation
                End If
                ConcatenatedArray(iLCounter) = Buffer(0)
                RaiseEvent AfterExportRecord
            Next iLCounter
            '@----------------------------------------------------------------------------
            'Concatenate records
            JoinRecordsFields = Join$(ConcatenatedArray, recDelimiter)
            '@----------------------------------------------------------------------------
            'Free the memory
            Erase ConcatenatedArray
            Erase Buffer
        Else
            '@----------------------------------------------------------------------------
            'Set array sizes
            ReDim ConcatenatedArray(LB1 To UB1)
            '@----------------------------------------------------------------------------
            'Concatenate fields
            For iLCounter = LB1 To UB1
                tmpBuffer = RecordsArray(iLCounter)
                LB2 = LBound(tmpBuffer)
                UB2 = UBound(tmpBuffer)
                ReDim Buffer(LB2 To UB2)
                For jLCounter = LB2 To UB2
                    Buffer(jLCounter) = tmpBuffer(jLCounter)
                    If UseUnixEscapeSeq Then
                        UnixEscapeField Buffer(jLCounter), CoerceChr, FldDelimiter, _
                                        recDelimiter
                    Else
                        EscapeField Buffer(jLCounter), CoerceChr, EscapeSequence, _
                                    FldDelimiter, recDelimiter, EnforcedQuotation
                    End If
                Next jLCounter
                ConcatenatedArray(iLCounter) = Join$(Buffer, FldDelimiter)
                RaiseEvent AfterExportRecord
            Next iLCounter
            '@----------------------------------------------------------------------------
            'Concatenate records
            JoinRecordsFields = Join$(ConcatenatedArray, recDelimiter)
            '@----------------------------------------------------------------------------
            'Free the memory
            Erase ConcatenatedArray
            Erase Buffer
        End If
    End If
    Exit Function
JoinRecordsFields_Error:
    JoinRecordsFields = vbNullString
    P_ERROR_NUMBER = err.Number
    P_ERROR_DESC = err.Description
    P_ERROR_SOURCE = err.Source
End Function
Private Sub LoadNewStream()
    '@----------------------------------------------------------------------------
    'Load new stream
    With CSVstream
        .ReadText
        RaiseEvent AfterLoadStream
        StreamWithQuotes = InStrB(1, .bufferString, QuoteChar)
        StreamEnd = .atEndOfStream
        If UseUnixEscapeSeq Then
            tmpCSV() = Split(UnixToStandardEscapeSeq(.bufferString, UnixEscapeSeq, _
                                EscapeSequence), RecordDelimiter)
        Else
            tmpCSV() = Split(.bufferString, RecordDelimiter)
        End If
    End With
    index = 0
    MaxIndex = UBound(tmpCSV)
    If Not StreamEnd Then
        If tmpCSV(MaxIndex) = vbNullString Then 'Streams lefts empty value when splited
            MaxIndex = MaxIndex - 1
        End If
    Else
        If IgnoreEmptyLines Then
            If tmpCSV(MaxIndex) = vbNullString Then 'Skip empty line
                MaxIndex = MaxIndex - 1
            End If
        End If
    End If
End Sub
Private Function MaxIndexVal(ByRef values As Variant) As Long
    Dim tmpVal As Double, idx0 As Long, idx As Long

    idx0 = LBound(values)
    idx = idx0
    tmpVal = CDbl(values(idx0))
    For idx0 = idx0 + 1 To UBound(values)
        If CDbl(values(idx0)) > tmpVal Then
            idx = idx0
            tmpVal = CDbl(values(idx0))
        End If
    Next idx0
    MaxIndexVal = idx
End Function
Public Function MergeFields(indexes As String, CharToMergeWith As String) As CSVinterface
    P_CSV_DATA.MergeFields indexes, CharToMergeWith
    Set MergeFields = Me
End Function
''' <summary>
''' Returns a copy of the current instance.
''' </summary>
Public Function Clone() As Object
    Set Clone = Me
End Function
Private Sub ClearDirectory(path As String)
    On Error Resume Next
    Kill path & "*.*"  ' delete all files in the folder
End Sub
''' <summary>
''' Sorts an external CSV file on disk.
''' </summary>
''' <param name="filePath">CSV file path, including file extension.</param>
''' <param name="sortingKeys">The columns/keys for the logical comparisons.</param>
''' <param name="headers">Indicates if the CSV file has a header row.</param>
''' <param name="streamSize">Single textstream factor size.</param>
''' <param name="sortAlgorithm">Algorithm used to sort the data.</param>
''' <param name="ExportationBunchSize">The amount of items to export in a single operation.</param>
Public Function SortOnDisk(ByVal filePath As String, _
                                Optional ByRef sortingKeys As Variant = 1, _
                                Optional Headers As Boolean = True, _
                                Optional streamSize As Single = 20, _
                                Optional ByVal sortAlgorithm As SortingAlgorithms = SortingAlgorithms.SA_Quicksort, _
                                Optional ByVal ExportationBunchSize As Long = 10000) As String

    On Error GoTo Sort_Error_Handler:

    Dim CSV As CSVinterface
    Dim iCounter As Long
    Dim inputTextStream As CSVTextStream
    Dim inputTextStreams() As CSVTextStream
    Dim streamsConsumedItems() As Long
    Dim streamsItemsCount() As Long
    Dim tmpOutputPath As String
    Dim OutPutFpath As String
    Dim OutputPath As String
    Dim sampledItems() As CSVinterface
    Dim sortedFiles As CSVArrayList
    Dim FirstStream As Boolean
    Dim fileHeader As Object
    Dim SortedCounter As Long
    Dim exportConfig As CSVparserConfig
    Dim sortedItems As CSVArrayList
    Dim ComparisonList As CSVArrayList
    Dim ItemIndex As Long
    Dim importConfig As CSVparserConfig

    tmpOutputPath = GetFilePath(filePath) & "tmp-sorted" & Application.PathSeparator
    If LenB(Dir(tmpOutputPath, vbDirectory)) > 0 Then 'Delete previous works
        ClearDirectory tmpOutputPath
    Else
        MkDir tmpOutputPath
    End If
    '@--------------------------------------------------------------------------------
    'Sort blocks of data with streamSize MB
    Set inputTextStream = New CSVTextStream
    FirstStream = True
    Set CSV = New CSVinterface
    With CSV.parseConfig
        .Headers = Headers
        .endingRecord = 1
    End With
    With inputTextStream
        .bufferSize = streamSize
        .endStreamOnLineBreak = True
        .OpenStream filePath
        .ReadText
        'Set the dialect
        Set CSV.parseConfig.dialect = CSV.SniffDelimiters(CSV.parseConfig, .bufferString)
        '@--------------------------------------------------------------------------------
        'Get header
        Set fileHeader = CSV.ImportFromCSVString(.bufferString, CSV.parseConfig).items.Clone
        '@--------------------------------------------------------------------------------
        'Consume first stream
        CSV.parseConfig.startingRecord = 2
        CSV.parseConfig.endingRecord = 0
        CSV.ImportFromCSVString .bufferString, CSV.parseConfig
        CSV.parseConfig.startingRecord = 1
        CSV.parseConfig.Headers = False
        '@--------------------------------------------------------------------------------
        'Sort and export
        Set sortedFiles = New CSVArrayList
        Set exportConfig = New CSVparserConfig
        CSV.Sort sortingKeys:=sortingKeys, sortAlgorithm:=sortAlgorithm
        SortedCounter = SortedCounter + 1
        Set exportConfig.dialect = CSV.parseConfig.dialect
        exportConfig.path = tmpOutputPath & "tmp_" & SortedCounter & ".csv"
        CSV.ExportToCSV CSV.items, exportConfig, enableDelimitersSniffing:=False
        sortedFiles.Add exportConfig.path
        Do While Not .atEndOfStream
            .ReadText
            CSV.ImportFromCSVString .bufferString, CSV.parseConfig
            CSV.Sort sortingKeys:=sortingKeys, sortAlgorithm:=sortAlgorithm
            SortedCounter = SortedCounter + 1
            exportConfig.path = tmpOutputPath & "tmp_" & SortedCounter & ".csv"
            CSV.ExportToCSV CSV.items, exportConfig, enableDelimitersSniffing:=False
            sortedFiles.Add exportConfig.path
        Loop
        .CloseStream
    End With
    '@--------------------------------------------------------------------------------
    'Configure output
    OutputPath = GetFilePath(filePath)
    OutPutFpath = OutputPath & GetFileNameAndExtension(filePath)(0) & "-sorted.csv"
    If FileExists(OutPutFpath) Then 'Delete prev sorted file
        Kill OutPutFpath
    End If
    'Point to target file
    exportConfig.path = OutPutFpath
    'Export header
    If Headers Then
        CSV.ExportToCSV fileHeader, exportConfig, enableDelimitersSniffing:=False
    End If
    ReDim inputTextStreams(0 To sortedFiles.count - 1)
    ReDim streamsItemsCount(0 To sortedFiles.count - 1)
    ReDim sampledItems(0 To sortedFiles.count - 1)
    Set ComparisonList = New CSVArrayList
    'Configure import options
    Set importConfig = exportConfig.CopyConfig
    With importConfig
        .path = vbNullString
        .Headers = False
    End With
    'Configure readers and parsers
    ComparisonList.indexing = True
    For iCounter = 0 To sortedFiles.count - 1
        Set inputTextStreams(iCounter) = New CSVTextStream
        Set sampledItems(iCounter) = New CSVinterface
        With inputTextStreams(iCounter)
            .endStreamOnLineBreak = True
            .OpenStream sortedFiles(iCounter)
            .ReadText
        End With
        With sampledItems(iCounter)
            Set .parseConfig = importConfig
            .ImportFromCSVString inputTextStreams(iCounter).bufferString, .parseConfig 'Import and store
            streamsItemsCount(iCounter) = .count
            'Append the first record to comparison list
            ComparisonList.AddIndexedItem CStr(iCounter), .items(0)
        End With
    Next iCounter
    ReDim streamsConsumedItems(0 To sortedFiles.count - 1)
    '@--------------------------------------------------------------------------------
    'Merge sorted data
    Set sortedItems = New CSVArrayList
    Do
        With ComparisonList
            .Sort sortingKeys:=sortingKeys
            'Save to sorted items
            sortedItems.Add .item(0)
            ItemIndex = CLng(.itemKey(0)) 'Worked record key
            'Export if needed
            If sortedItems.count = ExportationBunchSize Then
                CSV.ExportToCSV sortedItems, exportConfig, enableDelimitersSniffing:=False
                sortedItems.Clear
            End If
            .Swap 0, .count - 1, True
            .RemoveAt .count - 1  'Remove worked item
            'Update the count of consumed items
            streamsConsumedItems(ItemIndex) = streamsConsumedItems(ItemIndex) + 1
            'Add new item from the target list, if there are any remaining items
            If streamsConsumedItems(ItemIndex) = streamsItemsCount(ItemIndex) Then
                If Not inputTextStreams(ItemIndex).atEndOfStream Then
                    'Read new data block
                    inputTextStreams(ItemIndex).ReadText
                    sampledItems(ItemIndex).ImportFromCSVString _
                                            inputTextStreams(ItemIndex).bufferString, _
                                            sampledItems(ItemIndex).parseConfig 'Import and store
                    streamsItemsCount(ItemIndex) = sampledItems(ItemIndex).count
                    streamsConsumedItems(ItemIndex) = 0
                Else
                    inputTextStreams(ItemIndex).CloseStream 'Done with the file
                    Set sampledItems(ItemIndex) = Nothing
                End If
            End If
            If Not sampledItems(ItemIndex) Is Nothing Then
                .AddIndexedItem CStr(ItemIndex), sampledItems(ItemIndex)(streamsConsumedItems(ItemIndex)) 'Next item
            End If
        End With
    Loop While ComparisonList.count
    If sortedItems.count > 0 Then 'Export remaining data
        CSV.ExportToCSV sortedItems, exportConfig, enableDelimitersSniffing:=False
        sortedItems.Clear
    End If
    ClearDirectory tmpOutputPath  ' delete all files in the folder
    RmDir tmpOutputPath           ' delete folder
    Set inputTextStream = Nothing
    Erase inputTextStreams
    Erase sampledItems
    Erase streamsConsumedItems
    Erase streamsItemsCount
    SortOnDisk = OutPutFpath
    Exit Function
Sort_Error_Handler:
    P_ERROR_NUMBER = err.Number
    P_ERROR_DESC = "[BigCSVsort] " & err.Description
    P_ERROR_SOURCE = err.Source
End Function

Public Sub OpenSeqReader(configObj As CSVparserConfig, _
                            ParamArray FilterColumns() As Variant)
    On Error GoTo OpenSeqReader_Error_Handler
    With configObj
        If FileExists(.path) Then
            If FileLen(.path) = 0 Then
                GoTo OpenSeqReader_EmptyFile
            End If
            If UBound(.dTypingTemplate) >= UBound(.dTypingLinks) Then
                Dim tmpConf As CSVparserConfig
                Dim EmptyParam() As Variant
                ReDim EmptyParam(-1 To -1) 'Emule empty ParamArray

                If .delimitersGuessing Then
                    Set configObj.dialect = SniffDelimiters(configObj)
                End If
                static_tmpRequested() = FilterColumns
                If .Headers Then
                    '@--------------------------------------------------------------------------------
                    'Read the whole header
                    Set P_CSV_HEADER = New CSVArrayList
                    Set tmpConf = configObj.CopyConfig
                    tmpConf.startingRecord = 1
                    tmpConf.endingRecord = 1
                    StreamParseCSV tmpConf, P_CSV_HEADER, EmptyParam
                    Set tmpConf = Nothing
                End If
                If P_ERROR_NUMBER <> 0& Then GoTo OpenSeqReader_Error_Handler
            Else
                P_SUCCESSFUL_IMPORT = False
                P_ERROR_DESC = "the config object has an invalid Dynamic Typing Template (DTT). " _
                                & "The number of Dynamic Typing Links (DTL) must be less or equal than the number of " _
                                & "Dynamic Typing Targets Fields(DTTF) defined."
                P_ERROR_NUMBER = vbObjectError + 9007
                P_ERROR_SOURCE = "CSVinterface"
                GoTo OpenSeqReader_Error_Handler
            End If
        Else
            P_SUCCESSFUL_IMPORT = False
            P_ERROR_DESC = "The config object is not linked to a CSV file. " _
                                & "Ensure set the path property to valid CSV before import data."
            P_ERROR_NUMBER = vbObjectError + 9008
            P_ERROR_SOURCE = "CSVinterface"
            GoTo OpenSeqReader_Error_Handler
        End If
    End With
    '@----------------------------------------------------------------------------
    'Start variables
    static_SearchBeginningMark = 1&
    static_EscapeChrCount = 0&
    static_RecordsCount = 0&
    static_fieldsBound = 0&
    static_fieldsMaxBound = 0&
    static_StreamsCounter = 0
    Set static_tmpToken = New CSVArrayList
    Set static_CSVstream = New CSVTextStream
    With configObj
        static_CSVstream.bufferSize = .bufferSize
        static_CSVstream.endStreamOnLineBreak = True
        static_CSVstream.linebreakMatchingBehavior = EndLineMatchingBehavior.OnlyBackwardSense
        static_CSVstream.unifiedLFOutput = .multiEndOfLineCSV
        static_CSVstream.utf8EncodedFile = .utf8EncodedFile
        static_CSVstream.OpenStream .path
        static_commToken = AscW(.commentsToken)
        static_dynamicType = .dynamicTyping
        static_dTTemplate() = .dTypingTemplate
        static_dTTargets() = .dTypingLinks
        static_Headers = .Headers
        static_IgnoreCommentLines = .skipCommentLines
        static_IgnoreEmptyLines = .skipEmptyLines
        static_IgnoreLines = (static_IgnoreCommentLines Or static_IgnoreEmptyLines)
        static_TemplateDefined = .dTTemplateDefined
        static_TargetsDefined = .dTTemplateLinksDefined
        static_FieldDelimiter = .dialect.fieldsDelimiter
        static_RecordDelimiter = .dialect.recordsDelimiter
        static_UseUnixEscapeSeq = (.dialect.escapeMode = unix)
    End With
    static_TokenEndReached = False
    static_RecordEndReached = False
    If configObj.Headers Then
        Select Case UBound(static_tmpRequested)
            Case -1
                ReDim static_RequestedFieldsArray(0)
                static_RequestedFieldsArray(0) = -1
            Case Else
                static_RequestedFieldsArray() = GetHeaderIndexes(static_tmpRequested, P_CSV_HEADER)
        End Select
    Else
        ReDim static_RequestedFieldsArray(0)
        static_RequestedFieldsArray(0) = -1
    End If
    '@----------------------------------------------------------------------------
    'Set variables
    static_OverStringPointer = static_SearchBeginningMark
    static_MaxReqIndex = UBound(static_RequestedFieldsArray)
    '@----------------------------------------------------------------------------
    'Set Coerce Type to fields
    Select Case configObj.dialect.quoteToken
        Case 1
            static_QuoteChar = CHR_APOSTROPHE
        Case 2
            static_QuoteChar = CHR_DOUBLE_QUOTES
        Case Else
            static_QuoteChar = CHR_TILDE
    End Select
    '@----------------------------------------------------------------------------
    'Set the escape sequences
    static_UnixEscapeSeq = CHR_BACKSLASH & static_QuoteChar
    static_EscapeSequence = static_QuoteChar & static_QuoteChar
    static_EscapeSequence_Delim = CHR_BACKSLASH & static_FieldDelimiter
    '@----------------------------------------------------------------------------
    static_EscapeAscW = AscW(static_QuoteChar)
    static_FDAscW = AscW(static_FieldDelimiter)
    Set static_RecordToken = New CSVArrayList
    Set static_outputList = New CSVArrayList
    static_Index = 0
    '@----------------------------------------------------------------------------
    'Get CSV stream and populate a temp array
    With static_CSVstream
        .ReadText
        RaiseEvent AfterLoadStream
        static_StreamWithQuotes = InStrB(1, .bufferString, static_QuoteChar)
        static_StreamEnd = .atEndOfStream
    End With
    '@----------------------------------------------------------------------------
    'Ensuring the stream is sufficient enough to hold at least one CSV line
    With static_CSVstream
        Do While Not static_StreamEnd And InStrB(1, .bufferString, static_RecordDelimiter) = 0
            configObj.bufferSize = 2 * configObj.bufferSize
            .bufferSize = configObj.bufferSize
            .RestartPointer
            .ReadText
            RaiseEvent AfterLoadStream
            static_StreamWithQuotes = InStrB(1, .bufferString, static_QuoteChar)
            static_StreamEnd = .atEndOfStream
        Loop
        static_StreamsCounter = static_StreamsCounter + 1
        If static_UseUnixEscapeSeq Then
            static_tmpCSV() = Split(UnixToStandardEscapeSeq(.bufferString, _
                                    static_UnixEscapeSeq, static_EscapeSequence), static_RecordDelimiter)
        Else
            static_tmpCSV() = Split(.bufferString, static_RecordDelimiter)
        End If
    End With
    static_MaxIndex = UBound(static_tmpCSV)
    If Not static_StreamEnd Then
        If static_tmpCSV(static_MaxIndex) = vbNullString Then 'Streams lefts empty value when splited
            static_MaxIndex = static_MaxIndex - 1
        End If
    Else
        If static_IgnoreEmptyLines Then
            If static_tmpCSV(static_MaxIndex) = vbNullString Then 'Skip empty line
                static_MaxIndex = static_MaxIndex - 1
            End If
        End If
    End If
    static_EndOFStreamBuffer = (static_Index > static_MaxIndex)
    static_SeqReadStarted = True
    Exit Sub
OpenSeqReader_Error_Handler:
    static_SeqReadStarted = False
    P_SUCCESSFUL_IMPORT = False
    P_ERROR_DESC = "[OpenSeqReader]: " & err.Description
    P_ERROR_NUMBER = err.Number
    Exit Sub
OpenSeqReader_EmptyFile:
    Set P_CSV_DATA = Nothing
    err.Raise Number:=vbObjectError + 9011, _
                    Source:="CSVinterface Class", _
                    Description:="The specified source CSV is empty. Please check and try again."
    Resume OpenSeqReader_Error_Handler
End Sub
Private Sub ParseCSVstring(ByRef CSVtext As String, _
                            configObj As CSVparserConfig, _
                            outputList As CSVArrayList, _
                            ByRef FilterColumns() As Variant)
    Dim brCounter As Long
    Dim commToken As Long
    Dim dynamicType As Boolean
    Dim EndOfString As Boolean
    Dim EndRecord As Long
    Dim EscapeChrCount As Long
    Dim EscapedBRS As Boolean
    Dim EscapedWFS As Boolean
    Dim FieldID As Long
    Dim fieldDelimiter As String
    Dim hpointer As Long
    Dim IgnoreCommentLines As Boolean
    Dim IgnoreLines As Boolean
    Dim ImportDepth As Long
    Dim ImportSwitch As Boolean
    Dim IncludeRecord As Boolean
    Dim IsHeader As Boolean
    Dim IsWellEscapedField As Boolean
    Dim KnownEnding As Boolean
    Dim MaxReqIndex As Long
    Dim NotEvenEscapeChrNumber As Boolean
    Dim RecordEndReached As Boolean
    Dim RecordsCount As Long
    Dim RecordToken As CSVArrayList
    Dim SavedRecords As Long
    Dim SearchBeginningMark As Long
    Dim SplittedToken As Boolean
    Dim StartRecord As Long
    Dim tmpRequested() As Variant
    Dim tmpToken As CSVArrayList
    Dim TokenBeginningPos As Long
    Dim TokenEndingPos As Long
    Dim TokenEndReached As Boolean
    Dim OpenTokenArrIdx As Long
    Dim OpenTokenStartPos As Long

    On Error GoTo ParseCSVstring_Error_Handler
    '@----------------------------------------------------------------------------
    'Start variables
    SearchBeginningMark = 1&
    SearchBeginningMark = 1&
    EscapeChrCount = 0&
    RecordsCount = 0&
    SavedRecords = 0&
    P_VECTORS_REGULAR_BOUND = 0&
    P_VECTORS_MAX_BOUND = 0&
    Set tmpToken = New CSVArrayList
    TokenEndReached = False
    RecordEndReached = False
    With configObj
        commToken = AscW(.commentsToken)
        dynamicType = .dynamicTyping
        dTTemplate() = .dTypingTemplate
        dTTargets() = .dTypingLinks
        IgnoreCommentLines = .skipCommentLines
        IgnoreEmptyLines = .skipEmptyLines
        IgnoreLines = (IgnoreCommentLines Or IgnoreEmptyLines)
        TemplateDefined = .dTTemplateDefined
        TargetsDefined = .dTTemplateLinksDefined
        fieldDelimiter = .dialect.fieldsDelimiter
        RecordDelimiter = .dialect.recordsDelimiter
        EndRecord = .endingRecord
        StartRecord = .startingRecord
        UseUnixEscapeSeq = (.dialect.escapeMode = unix)
    End With
    tmpRequested() = FilterColumns
    Select Case UBound(tmpRequested)
        Case -1
            ReDim RequestedFieldsArray(0)
            RequestedFieldsArray(0) = -1
        Case Else
            RequestedFieldsArray() = GetHeaderIndexes(tmpRequested, P_CSV_HEADER)
    End Select

    '@----------------------------------------------------------------------------
    'Set variables
    OverStringPointer = SearchBeginningMark
    MaxReqIndex = UBound(RequestedFieldsArray)
    '@----------------------------------------------------------------------------
    'Check if user want Import all records
    KnownEnding = True
    Select Case EndRecord
        Case 0&
            If StartRecord >= 1 Then
                KnownEnding = False
            Else
                ImportDepth = 1&
            End If
        Case Is < StartRecord
            ImportDepth = 1&
        Case Else
            ImportDepth = EndRecord - StartRecord + 1&
    End Select
    If StartRecord = 1& Then
        If configObj.headersOmission And configObj.Headers Then
            If ImportDepth > 1& Or EndRecord = 0& Then
                StartRecord = 2&
                ImportDepth = ImportDepth - 1&
            End If
        End If
    End If
    '@----------------------------------------------------------------------------
    'Set Coerce Type to fields
    Select Case configObj.dialect.quoteToken
        Case 1
            QuoteChar = CHR_APOSTROPHE
        Case 2
            QuoteChar = CHR_DOUBLE_QUOTES
        Case Else
            QuoteChar = CHR_TILDE
    End Select
    '@----------------------------------------------------------------------------
    'Set the escape sequences
    UnixEscapeSeq = CHR_BACKSLASH & QuoteChar
    EscapeSequence = QuoteChar & QuoteChar
    EscapeSequence_Delim = CHR_BACKSLASH & fieldDelimiter
    '@----------------------------------------------------------------------------
    EscapeAscW = AscW(QuoteChar)
    FDAscW = AscW(fieldDelimiter)
    Set RecordToken = New CSVArrayList
    index = 0
    '@----------------------------------------------------------------------------
    'Populate a temp array
    StreamWithQuotes = InStrB(1, CSVtext, QuoteChar)
    If UseUnixEscapeSeq Then
        tmpCSV() = Split(UnixToStandardEscapeSeq(CSVtext, UnixEscapeSeq, EscapeSequence), RecordDelimiter)
    Else
        tmpCSV() = Split(CSVtext, RecordDelimiter)
    End If
    MaxIndex = UBound(tmpCSV)
    If MaxIndex = -1 Then outputList.Add2 vbNullString: outputList.ShrinkBuffer: Exit Sub
    '@----------------------------------------------------------------------------
    'Skip empty lines at end of string
    If IgnoreEmptyLines Then
        brCounter = 0
        Do While LenB(tmpCSV(MaxIndex - brCounter)) = 0
            brCounter = brCounter + 1
        Loop
        MaxIndex = MaxIndex - brCounter
    End If
    If IgnoreLines Then
        '@----------------------------------------------------------------------------
        'Skip lines if needed
        SkipUnwantedLines index, MaxIndex, tmpCSV, commToken, _
                            IgnoreCommentLines, IgnoreEmptyLines
    End If
    EndOfString = (index > MaxIndex)
    If EndOfString Then 'The file has no significant data
        GoTo ParseCSVstring_NoSignificantData
    End If
    '@----------------------------------------------------------------------------
    'Loop through CSV
    Do
        IsHeader = (StartRecord = 1 And RecordsCount = 0)
        LenCurrentIndex = LenB(tmpCSV(index))
        '@----------------------------------------------------------------------------
        'Restart pointers
        RecordEndReached = False
        FieldID = 0
        IncludeRecord = (RecordsCount + 1& >= StartRecord)
        Do
            '@----------------------------------------------------------------------------
            'Process the current Token
            TokenEndReached = False
            TokenBeginningPos = OverStringPointer
            If StreamWithQuotes Then
                Select Case OverStringPointer
                    Case Is < LenCurrentIndex
                        ASCIIcharw = AscW(MidB$(tmpCSV(index), OverStringPointer, 2))
                        '@----------------------------------------------------------------------------
                        'Ignore spaces & tabs
                        SkipSpacesAndTabs
                        OpenedToken = (ASCIIcharw = EscapeAscW)
                    Case Else 'Avoid over runs
                        OpenedToken = False
                End Select
            Else
                OpenedToken = False
            End If
            If OpenedToken Then
                If OpenTokenArrIdx = 0 Then
                    OpenTokenArrIdx = index
                End If
                '@----------------------------------------------------------------------------
                'Try to escape the Token
                EscapeChrCount = EscapeChrCount + 1&
                TokenBeginningPos = OverStringPointer + 2&
                SearchBeginningMark = TokenBeginningPos
                hpointer = InStrB(SearchBeginningMark, tmpCSV(index), QuoteChar) 'EOFld marker
                If hpointer Then EscapeChrCount = EscapeChrCount + 1&
                Do While (Not TokenEndReached And Not EndOfString)
                    '@----------------------------------------------------------------------------
                    'Buffering if needed
                    If Not SplittedToken Then SplittedToken = (hpointer = 0)
                    Do
                        Do While hpointer = 0& And Not EndOfString 'Quote char may be on subsequent Index
                            If IncludeRecord Then
                                Select Case SearchBeginningMark
                                    Case 1 'Store the full Token from the current Index
                                        tmpToken.Add tmpCSV(index)
                                    Case Else 'Store the Right most Token portion from the current Index
                                        tmpToken.Add RightB$(tmpCSV(index), LenCurrentIndex - SearchBeginningMark + 1)
                                End Select
                                If OpenTokenStartPos = 0 Then
                                    OpenTokenStartPos = SearchBeginningMark
                                End If
                            End If
                            SplittedToken = True
                            'tmpCSV(index) = vbNullString
                            index = index + 1&
                            EndOfString = (index > MaxIndex)
                            If Not EndOfString Then
                                LenCurrentIndex = LenB(tmpCSV(index))
                                SearchBeginningMark = 1
                                hpointer = InStrB(SearchBeginningMark, tmpCSV(index), QuoteChar) 'Find quote char
                                If hpointer Then EscapeChrCount = EscapeChrCount + 1&
                            Else
                                hpointer = 0&
                            End If
'                            If hpointer = 0& And EndOfString Then GoTo ParseCSVstring_MissingEscapeChar 'Missing quote char
                            If hpointer = 0& And EndOfString Then 'Missing quote char
                                P_ERROR_DESC = P_ERROR_DESC & "Missing some quote char. The parser retrieved the record in recovery mode. " & _
                                               "[Review the record #" & RecordsCount + 1 & ", field #" & RecordToken.count + 1 _
                                                 & " on the source CSV file]."
                                RecordEndReached = True
                                OverStringPointer = 1&
                                tmpToken.Clear
                                LenCurrentIndex = LenB(tmpCSV(OpenTokenArrIdx))
                                If OpenTokenStartPos = 1 Then 'Save the entire record exclusively in recovery mode
                                    RecordToken.Add MidB$(tmpCSV(OpenTokenArrIdx), 3)
                                Else 'Save a portion of the record in recovery mode
                                    RecordToken.Add RightB$(tmpCSV(OpenTokenArrIdx), LenCurrentIndex - OpenTokenStartPos + 1)
                                End If
                                index = OpenTokenArrIdx + 1
                                GoTo Restore_From_Missing_EscapeChar
                            End If
                        Loop
                        OverStringPointer = hpointer + 2&
                        NotEvenEscapeChrNumber = (EscapeChrCount And 1)
                        If OverStringPointer < LenCurrentIndex Then
                            ASCIIcharw = AscW(MidB$(tmpCSV(index), OverStringPointer, 2))
                            '@----------------------------------------------------------------------------
                            'Ignore spaces & tabs
                            SkipSpacesAndTabs
                            '@----------------------------------------------------------------------------
                            'Check if we are at the end of the field
                            If UseUnixEscapeSeq Then
                                EscapedWFS = (ASCIIcharw = FDAscW) And (Not NotEvenEscapeChrNumber) And _
                                                (Not (AscW(MidB$(tmpCSV(index), OverStringPointer - 2, 2)) = ASCW_CHR_BACKSLASH))
                            Else
                                EscapedWFS = (ASCIIcharw = FDAscW) And (Not NotEvenEscapeChrNumber)
                            End If
                        Else
                            EscapedWFS = False
                            EscapedBRS = Not NotEvenEscapeChrNumber
                        End If
                        IsWellEscapedField = (EscapedWFS Or EscapedBRS)
                        '@----------------------------------------------------------------------------
                        'Advance the search
                        If Not IsWellEscapedField Then
                            hpointer = InStrB(hpointer + 2&, tmpCSV(index), QuoteChar) 'EOFld marker
                            If hpointer Then EscapeChrCount = EscapeChrCount + 1&
                        End If
                    Loop While Not IsWellEscapedField
                    If EscapedWFS Then 'Quote char succeeded by fields separator char
                        TokenEndingPos = hpointer - 1&
                        TokenEndReached = True
                        EscapeChrCount = 0&
                        If IncludeRecord Then
                            '@----------------------------------------------------------------------------
                            'Save the Token
                            Select Case SplittedToken
                                Case False 'Just save Token
                                    If RequestedField(FieldID + 1&, MaxReqIndex) Then 'Filter fields
                                        Select Case dynamicType
                                            Case False
                                                RecordToken.Add UnescapeField(MidB$(tmpCSV(index), TokenBeginningPos, _
                                                                                TokenEndingPos - TokenBeginningPos + 1), _
                                                                            QuoteChar, EscapeSequence)
                                            Case Else
                                                RecordToken.Add TypeData(UnescapeField(MidB$(tmpCSV(index), TokenBeginningPos, _
                                                                                        TokenEndingPos - TokenBeginningPos + 1), _
                                                                                    QuoteChar, EscapeSequence), _
                                                                        FieldID, IsHeader)
                                        End Select
                                    End If
                                Case Else 'The Token is stored on tmpToken object
                                    tmpToken.Add LeftB$(tmpCSV(index), TokenEndingPos)
                                    If RequestedField(FieldID + 1&, MaxReqIndex) Then 'Filter fields
                                        RecordToken.Add UnescapeField(Join$(tmpToken.items, RecordDelimiter), _
                                                                    QuoteChar, EscapeSequence) 'Save whole Token
                                    End If
                                    tmpToken.Clear 'Clear tmpToken content
                            End Select
                        End If
                        SplittedToken = False 'Restart Token status
                        OverStringPointer = OverStringPointer + 2& 'Advance over the field separator char
                        RecordEndReached = False
                    Else
                        If EscapedBRS Then  'Quote char succeeded
                                            'by records separator char
                            TokenEndingPos = hpointer - 1&
                            TokenEndReached = True
                            EscapeChrCount = 0&
                            RecordEndReached = True 'EOR
                            If IncludeRecord Then
                                '@----------------------------------------------------------------------------
                                'Save the Token
                                Select Case SplittedToken
                                    Case False 'Just save Token
                                        If RequestedField(FieldID + 1&, MaxReqIndex) Then 'Filter fields
                                            Select Case dynamicType
                                                Case False
                                                    RecordToken.Add UnescapeField(MidB$(tmpCSV(index), TokenBeginningPos, _
                                                                                        TokenEndingPos - TokenBeginningPos + 1), _
                                                                                QuoteChar, EscapeSequence)
                                                Case Else
                                                    RecordToken.Add TypeData(UnescapeField(MidB$(tmpCSV(index), TokenBeginningPos, _
                                                                                                TokenEndingPos - TokenBeginningPos + 1), _
                                                                                        QuoteChar, EscapeSequence), _
                                                                                FieldID, IsHeader)
                                            End Select
                                        End If
                                    Case Else 'The Token is stored on tmpToken array
                                        tmpToken.Add LeftB$(tmpCSV(index), OverStringPointer - 3)
                                        If RequestedField(FieldID + 1&, MaxReqIndex) Then 'Filter fields
                                            RecordToken.Add UnescapeField(Join$(tmpToken.items, RecordDelimiter), _
                                                                        fieldDelimiter, EscapeSequence_Delim) 'Save whole Token
                                        End If
                                        tmpToken.Clear 'Clear tmpToken content
                                End Select
                            End If
                            SplittedToken = False 'Restart Token status
                            'tmpCSV(index) = vbNullString
                            index = index + 1&
                            EndOfString = (index > MaxIndex)
                            OverStringPointer = 1&
                        Else
                            If IncludeRecord Then
                                If SplittedToken Then 'Store the Left most Token portion from the current Index
                                    tmpToken.Add LeftB$(tmpCSV(index), OverStringPointer - 1)
                                End If
                            End If
                            SearchBeginningMark = OverStringPointer
                            hpointer = InStrB(SearchBeginningMark, tmpCSV(index), QuoteChar) 'Find quote char
                            If hpointer Then EscapeChrCount = EscapeChrCount + 1&
                            TokenEndReached = False
                        End If
                    End If
                Loop
Restore_From_Missing_EscapeChar:
                OpenTokenStartPos = 0
            Else
                OpenTokenArrIdx = 0
                '@----------------------------------------------------------------------------
                'Try to mark the Starting and Ending position for the field
                hpointer = InStrB(OverStringPointer, tmpCSV(index), fieldDelimiter)
                If UseUnixEscapeSeq Then
                    Do While hpointer
                        If hpointer > 2 Then
                            UnixEscpedSep = (AscW(MidB$(tmpCSV(index), hpointer - 2, 2)) = ASCW_CHR_BACKSLASH)
                        Else
                            UnixEscpedSep = False
                        End If
                        If UnixEscpedSep Then
                            hpointer = InStrB(hpointer + 1, tmpCSV(index), fieldDelimiter)
                        Else
                            Exit Do
                        End If
                    Loop
                End If
                If hpointer Then 'Found fields delimiter char
                    SplittedToken = False
                    TokenEndReached = True
                    TokenEndingPos = hpointer - 1&
                    OverStringPointer = hpointer + 2&
                    SearchBeginningMark = OverStringPointer
                    RecordEndReached = False
                    If IncludeRecord Then
                        '@----------------------------------------------------------------------------
                        'Save the Token
                        If RequestedField(FieldID + 1&, MaxReqIndex) Then 'Filter fields
                            Select Case dynamicType
                                Case False
                                    If UseUnixEscapeSeq Then
                                        RecordToken.Add UnescapeField(UnescapeField(MidB$(tmpCSV(index), TokenBeginningPos, _
                                                                        TokenEndingPos - TokenBeginningPos + 1), _
                                                                    fieldDelimiter, EscapeSequence_Delim), _
                                                                QuoteChar, EscapeSequence)
                                    Else
                                        RecordToken.Add MidB$(tmpCSV(index), TokenBeginningPos, _
                                                            TokenEndingPos - TokenBeginningPos + 1)
                                    End If
                                Case Else
                                    If UseUnixEscapeSeq Then
                                        RecordToken.Add TypeData(UnescapeField(UnescapeField(MidB$(tmpCSV(index), TokenBeginningPos, _
                                                                                TokenEndingPos - TokenBeginningPos + 1), _
                                                                            fieldDelimiter, EscapeSequence_Delim), _
                                                                    QuoteChar, EscapeSequence), _
                                                            FieldID, IsHeader)
                                    Else
                                        RecordToken.Add TypeData(MidB$(tmpCSV(index), TokenBeginningPos, _
                                                                    TokenEndingPos - TokenBeginningPos + 1), _
                                                            FieldID, IsHeader)
                                    End If
                            End Select
                        End If
                    End If
                Else 'Missing fields delimiter char
                    If UseUnixEscapeSeq Then
                        Do While Not EndOfString And (AscW(RightB$(tmpCSV(index), 2)) = ASCW_CHR_BACKSLASH) 'Unix escaped Record delimiter
                            SplittedToken = True
                            If IncludeRecord Then
                                Select Case SearchBeginningMark
                                    Case 1 'Store the full Token from the current Index
                                        tmpToken.Add MidB$(tmpCSV(index), 1, LenCurrentIndex - 1)
                                    Case Else 'Store the Right most Token portion from the current Index
                                        tmpToken.Add MidB$(tmpCSV(index), TokenBeginningPos, LenCurrentIndex - TokenBeginningPos - 1)
                                End Select
                            End If
                            'tmpCSV(index) = vbNullString
                            index = index + 1&
                            OverStringPointer = 1&
                            EndOfString = (index > MaxIndex)
                            If Not EndOfString Then
                                LenCurrentIndex = LenB(tmpCSV(index))
                                SearchBeginningMark = 1
                                '@----------------------------------------------------------------------------
                                'Try to mark the Starting and Ending position for the field
                                hpointer = InStrB(OverStringPointer, tmpCSV(index), fieldDelimiter)
                                Do While hpointer
                                    If hpointer > 2 Then
                                        UnixEscpedSep = (AscW(MidB$(tmpCSV(index), hpointer - 2, 2)) = ASCW_CHR_BACKSLASH)
                                    Else
                                        UnixEscpedSep = False
                                    End If
                                    If UnixEscpedSep Then
                                        hpointer = InStrB(hpointer + 1, tmpCSV(index), fieldDelimiter)
                                    Else
                                        Exit Do
                                    End If
                                Loop
                            Else
                                hpointer = 0&
                            End If
                            If hpointer Then 'Unescaped fields delimiter found
                                TokenEndReached = True
                                TokenEndingPos = hpointer - 1&
                                OverStringPointer = hpointer + 2&
                                SearchBeginningMark = OverStringPointer
                                RecordEndReached = False
                                tmpToken.Add LeftB$(tmpCSV(index), TokenEndingPos) 'Store the Left most Token portion from the current Index
                                Exit Do
                            Else
                                If Not UnixEscpedSep Then 'End of record
                                    If IncludeRecord Then
                                        Select Case SearchBeginningMark
                                            Case 1 'Store the full Token from the current Index
                                                tmpToken.Add MidB$(tmpCSV(index), 1, LenCurrentIndex)
                                            Case Else 'Store the Right most Token portion from the current Index
                                                tmpToken.Add MidB$(tmpCSV(index), TokenBeginningPos, LenCurrentIndex - TokenBeginningPos)
                                        End Select
                                    End If
                                End If
                            End If
                        Loop
                    Else
                        hpointer = 0&
                    End If
                    If hpointer = 0 Then
                        TokenEndReached = True
                        RecordEndReached = True
                    End If
                    If IncludeRecord Then
                        '@----------------------------------------------------------------------------
                        'Save the Token
                        Select Case SplittedToken
                            Case False 'Just save Token
                                If RequestedField(FieldID + 1&, MaxReqIndex) Then 'Filter fields
                                    Select Case dynamicType
                                        Case False
                                            If UseUnixEscapeSeq Then
                                                RecordToken.Add UnescapeField(UnescapeField(RightB$(tmpCSV(index), LenCurrentIndex _
                                                                                - TokenBeginningPos + 1), _
                                                                            fieldDelimiter, EscapeSequence_Delim), _
                                                                        QuoteChar, EscapeSequence)
                                            Else
                                                RecordToken.Add RightB$(tmpCSV(index), LenCurrentIndex _
                                                                    - TokenBeginningPos + 1)
                                            End If
                                        Case Else
                                            If UseUnixEscapeSeq Then
                                                RecordToken.Add TypeData(UnescapeField(UnescapeField(RightB$(tmpCSV(index), LenCurrentIndex _
                                                                            - TokenBeginningPos + 1), _
                                                                            fieldDelimiter, EscapeSequence_Delim), _
                                                                        QuoteChar, EscapeSequence), _
                                                                    FieldID, IsHeader)
                                            Else
                                                RecordToken.Add TypeData(RightB$(tmpCSV(index), LenCurrentIndex _
                                                                            - TokenBeginningPos + 1), _
                                                                    FieldID, IsHeader)
                                            End If
                                    End Select
                                End If
                            Case Else 'The Token is stored on tmpToken array
                                If RequestedField(FieldID + 1&, MaxReqIndex) Then 'Filter fields
                                    If UseUnixEscapeSeq Then
                                        RecordToken.Add UnescapeField(UnescapeField(Join$(tmpToken.items, RecordDelimiter), _
                                                                        fieldDelimiter, EscapeSequence_Delim), _
                                                                    QuoteChar, EscapeSequence) 'Save whole Token
                                    Else
                                        RecordToken.Add Join$(tmpToken.items, RecordDelimiter)
                                    End If
                                End If
                                tmpToken.Clear 'Clear tmpToken content
                                SplittedToken = False 'Restart Token status
                        End Select
                    End If
                    If RecordEndReached Then
                        'tmpCSV(index) = vbNullString
                        index = index + 1&
                        EndOfString = (index > MaxIndex)
                        OverStringPointer = 1&
                    End If
                End If
            End If
            FieldID = FieldID + 1&
        Loop While (Not RecordEndReached And Not EndOfString)
        Select Case P_VECTORS_REGULAR_BOUND
            Case Is > 0
                If RecordToken.count - 1 > P_VECTORS_MAX_BOUND Then
                    P_VECTORS_MAX_BOUND = RecordToken.count - 1
                    RecordToken.ShrinkBuffer
                End If
            Case Else
                Select Case FieldID
                    Case Is > 0
                        P_VECTORS_REGULAR_BOUND = RecordToken.count - 1
                        RecordToken.ShrinkBuffer
                    Case Else
                        P_VECTORS_REGULAR_BOUND = FieldID
                End Select
                P_VECTORS_MAX_BOUND = P_VECTORS_REGULAR_BOUND
        End Select
        RecordsCount = RecordsCount + 1&
        EscapedBRS = False
        '@----------------------------------------------------------------------------
        'Start saving on given record
        If IncludeRecord Then
            outputList.Add RecordToken.items
            SavedRecords = SavedRecords + 1&
            RaiseEvent AfterImportRecord
            RecordToken.Reinitialize P_VECTORS_REGULAR_BOUND + 1
        End If
        P_VARYING_LENGTHS = (P_VECTORS_REGULAR_BOUND <> (FieldID - 1))
        If IgnoreLines Then
            '@----------------------------------------------------------------------------
            'Skip lines if needed
            SkipUnwantedLines index, MaxIndex, tmpCSV, commToken, _
                                IgnoreCommentLines, IgnoreEmptyLines
        End If
        EndOfString = (index > MaxIndex)
        Select Case KnownEnding
            Case False
                ImportSwitch = EndOfString
            Case Else
                ImportSwitch = (SavedRecords = ImportDepth) Or EndOfString
        End Select
    Loop While Not ImportSwitch
    '@----------------------------------------------------------------------------
    'Check if the array need to be shrinked at the output stage
    outputList.ShrinkBuffer
    Erase tmpCSV
    Set RecordToken = Nothing
    Set tmpToken = Nothing
    Exit Sub
ParseCSVstring_Error_Handler:
    Set outputList = Nothing
    Set tmpToken = Nothing
    Erase tmpCSV
    P_ERROR_NUMBER = err.Number
    P_ERROR_DESC = err.Description & " [Review the record #" & RecordsCount + 1 _
                                  & ", field #" & RecordToken.count + 1 _
                                   & " on the source CSV string]."
    Set RecordToken = Nothing
    P_ERROR_SOURCE = err.Source
    Exit Sub
ParseCSVstring_MissingEscapeChar:
    err.Raise Number:=vbObjectError + 9006, _
                    Source:="CSVinterface Class [ParseCSVstring]", _
                    Description:="Missing some quote char. Check the data and try again."
    Resume ParseCSVstring_Error_Handler
ParseCSVstring_NoSignificantData:
    err.Raise Number:=vbObjectError + 9010, _
                    Source:="CSVinterface Class [ParseCSVstring]", _
                    Description:="The String has no significant data. This can occur when " & _
                                    "the String has only empty or commented lines that can be omitted."
    Resume ParseCSVstring_Error_Handler
End Sub
Public Function RearrangeFields(FieldsOrder As String) As CSVinterface
    On Error GoTo ErrHandler_RearrangeFields
    If P_SUCCESSFUL_IMPORT And Not P_VARYING_LENGTHS Then
        P_CSV_DATA.RearrangeFields FieldsOrder
    Else
        P_ERROR_DESC = "[CSV Fields Rearrange]: Cannot rearrange the fields in the current instance." _
                        & " This is because there is no imported data or the records do not " _
                        & "have the same number of fields."
        P_ERROR_NUMBER = vbObjectError + 9020
        P_ERROR_SOURCE = "CSVinterface"
    End If
    Set RearrangeFields = Me
    Exit Function
ErrHandler_RearrangeFields:
    P_ERROR_DESC = "[CSV Fields Rearrange]: " & err.Description
    P_ERROR_NUMBER = err.Number
    P_ERROR_SOURCE = "CSVinterface"
    Exit Function
Undefined_RearrangeFields:
    err.Raise Number:=vbObjectError + 9021, _
                    Source:="CSVinterface Class", _
                    Description:="The order specified for the fields is incomplete. " _
                    & "Please enter all field indexes before continuing."
    Resume ErrHandler_RearrangeFields
End Function

Public Function RemoveField(aIndex As Long) As CSVinterface
    On Error GoTo ErrHandler_RemoveField
    If P_SUCCESSFUL_IMPORT And Not P_VARYING_LENGTHS Then
        Select Case aIndex
            Case 0 To P_VECTORS_REGULAR_BOUND
                P_CSV_DATA.RemoveField aIndex
            Case Else
                GoTo OutOfBounds_RemoveField
        End Select
    Else
        P_ERROR_DESC = "[CSV Field Remove]: Cannot remove the field in the current instance." _
                        & " This is because there is no imported data or the records do not " _
                        & "have the same number of fields."
        P_ERROR_NUMBER = vbObjectError + 9017
        P_ERROR_SOURCE = "CSVinterface"
    End If
    P_VECTORS_REGULAR_BOUND = P_VECTORS_REGULAR_BOUND - 1
    P_VECTORS_MAX_BOUND = P_VECTORS_MAX_BOUND - 1
    Set RemoveField = Me
    Exit Function
ErrHandler_RemoveField:
    P_ERROR_DESC = "[CSV Field Remove]: " & err.Description
    P_ERROR_NUMBER = err.Number
    P_ERROR_SOURCE = "CSVinterface"
    Exit Function
OutOfBounds_RemoveField:
    err.Raise Number:=vbObjectError + 9016, _
                    Source:="CSVinterface Class", _
                    Description:="The specified index is out of bounds. Please check and try again."
    Resume ErrHandler_RemoveField
End Function
Public Function RemoveRecords(aIndex As Long, Optional count As Long = 1) As CSVinterface
    On Error GoTo ErrHandler_RemoveRecords
    If P_SUCCESSFUL_IMPORT And Not P_VARYING_LENGTHS Then
        Select Case aIndex
            Case 0 To P_CSV_DATA.count - 1
                Dim cpRecord() As Variant

                ReDim cpRecord(0 To P_VECTORS_REGULAR_BOUND)
                With P_CSV_DATA
                    .RemoveRange aIndex, count
                End With
            Case Else
                GoTo OutOfBounds_RemoveRecords
        End Select
    Else
        P_ERROR_DESC = "[CSV Records Remove]: Cannot delete the records in the current instance." _
                        & " This is because there is no imported data or the records do not " _
                        & "have the same number of fields."
        P_ERROR_NUMBER = vbObjectError + 9019
        P_ERROR_SOURCE = "CSVinterface"
    End If
    Set RemoveRecords = Me
    Exit Function
ErrHandler_RemoveRecords:
    P_ERROR_DESC = "[CSV Records Remove]: " & err.Description
    P_ERROR_NUMBER = err.Number
    P_ERROR_SOURCE = "CSVinterface"
    Exit Function
OutOfBounds_RemoveRecords:
    err.Raise Number:=vbObjectError + 9016, _
                    Source:="CSVinterface Class", _
                    Description:="The specified index is out of bounds. Please check and try again."
    Resume ErrHandler_RemoveRecords
End Function
Private Function ReplaceIllegalCharacters(strIn As String, strChar As String) As String
    Dim strSpecialChars As String
    Dim i As Long
    strSpecialChars = "~""#%&*:<>?{|}/\[]" & chr(10) & chr(13)

    For i = 1 To Len(strSpecialChars)
        strIn = Replace(strIn, Mid$(strSpecialChars, i, 1), strChar)
    Next

    ReplaceIllegalCharacters = strIn
End Function
Private Function RequestedField(ByVal fieldIndex As Long, _
                                ByVal mxReq As Long) As Boolean
    Select Case RequestedFieldsArray(0)
        Case -1
            RequestedField = True
        Case Else
            Dim idx As Long
            Dim tmpBool As Boolean

            idx = 0
            Do
                tmpBool = (fieldIndex = RequestedFieldsArray(idx))
                idx = idx + 1
            Loop While Not tmpBool And idx <= mxReq
            RequestedField = tmpBool
    End Select
End Function
Private Function RequestedField2(ByVal fieldIndex As Long, _
                                ByVal mxReq As Long) As Boolean
    Select Case static_RequestedFieldsArray(0)
        Case -1
            RequestedField2 = True
        Case Else
            Dim idx As Long
            Dim tmpBool As Boolean

            idx = 0
            Do
                tmpBool = (fieldIndex = static_RequestedFieldsArray(idx))
                idx = idx + 1
            Loop While Not tmpBool And idx <= mxReq
            RequestedField2 = tmpBool
    End Select
End Function
Public Sub ResetToDefault()
    Set config = New CSVparserConfig
    Set P_CSV_DATA = New CSVArrayList
    P_ERROR_DESC = vbNullString
    P_ERROR_NUMBER = 0&
    P_ERROR_SOURCE = vbNullString
    P_SUCCESSFUL_EXPORT = False
    P_SUCCESSFUL_IMPORT = False
    P_VARYING_LENGTHS = False
End Sub
Private Sub ResetAlerts()
    P_ERROR_DESC = vbNullString
    P_ERROR_NUMBER = 0&
    P_ERROR_SOURCE = vbNullString
End Sub
Private Sub SkipSpacesAndTabs()
    '@----------------------------------------------------------------------------
    'Ignore spaces & tabs
    If FDAscW <> 9 And FDAscW <> 32 Then
        Do While ASCIIcharw = 32 Or ASCIIcharw = 9
            OverStringPointer = OverStringPointer + 2&
            If OverStringPointer > LenCurrentIndex Then Exit Do
            ASCIIcharw = AscW(MidB$(tmpCSV(index), OverStringPointer, 2))
        Loop
    Else
        If FDAscW = 9 Then
            Do While ASCIIcharw = 32
                OverStringPointer = OverStringPointer + 2&
                If OverStringPointer > LenCurrentIndex Then Exit Do
                ASCIIcharw = AscW(MidB$(tmpCSV(index), OverStringPointer, 2))
            Loop
        Else 'FDAscW = 32
            Do While ASCIIcharw = 9
                OverStringPointer = OverStringPointer + 2&
                If OverStringPointer > LenCurrentIndex Then Exit Do
                ASCIIcharw = AscW(MidB$(tmpCSV(index), OverStringPointer, 2))
            Loop
        End If
    End If
End Sub
Private Sub SkipUnwantedLines(ByRef idx As Long, _
                                ByRef MaxIdx As Long, _
                                ByRef arr() As String, _
                                ByVal CommentToken As Long, _
                                Optional SkipComments As Boolean = True, _
                                Optional skipEmptyLines As Boolean = True)
    Dim CurLength As Long
    Dim charCode As Long
    Dim UnwantedLine As Boolean

    '@----------------------------------------------------------------------------
    'Skip commented and empty lines if needed
    If idx <= MaxIdx Then
        Do
            CurLength = LenB(arr(idx))
            UnwantedLine = False
            If CurLength = 0 Then  'Empty line found
                If skipEmptyLines Then
                    UnwantedLine = True
                End If
            Else
                charCode = AscW(arr(idx))
                If charCode = CommentToken Then 'Commented line found
                    If SkipComments Then
                        UnwantedLine = True
                    End If
                End If
            End If
            If UnwantedLine Then
                idx = idx + 1&
            End If
        Loop While UnwantedLine And idx <= MaxIdx
    End If
End Sub
Public Function ShiftField(aIndex As Long, Shift As Long) As CSVinterface
On Error GoTo ErrHandler_ShiftField
    If P_SUCCESSFUL_IMPORT And Not P_VARYING_LENGTHS Then
        Select Case aIndex + Shift
            Case 0 To P_VECTORS_REGULAR_BOUND
               P_CSV_DATA.ShiftField aIndex, Shift
            Case Else
                GoTo OutOfBounds_ShiftField
        End Select
    Else
        P_ERROR_DESC = "[CSV Field Shift]: Cannot shift the field in the current instance." _
                        & " This is because there is no imported data or the records do not " _
                        & "have the same number of fields."
        P_ERROR_NUMBER = vbObjectError + 9024
        P_ERROR_SOURCE = "CSVinterface"
    End If
    P_VECTORS_REGULAR_BOUND = P_VECTORS_REGULAR_BOUND - 1
    P_VECTORS_MAX_BOUND = P_VECTORS_MAX_BOUND - 1
    Set ShiftField = Me
    Exit Function
ErrHandler_ShiftField:
    P_ERROR_DESC = "[CSV Field Shift]: " & err.Description
    P_ERROR_NUMBER = err.Number
    P_ERROR_SOURCE = "CSVinterface"
    Exit Function
OutOfBounds_ShiftField:
    err.Raise Number:=vbObjectError + 9016, _
                    Source:="CSVinterface Class", _
                    Description:="The specified index is out of bounds. Please check and try again."
    Resume ErrHandler_ShiftField
End Function
Public Function ShiftRecord(aIndex As Long, Shift As Long) As CSVinterface
On Error GoTo ErrHandler_ShiftRecord
    If P_SUCCESSFUL_IMPORT And Not P_VARYING_LENGTHS Then
        Select Case aIndex + Shift
            Case 0 To P_CSV_DATA.count - 1
                P_CSV_DATA.ShiftRecord aIndex, Shift
            Case Else
                GoTo OutOfBounds_ShiftRecord
        End Select
    Else
        P_ERROR_DESC = "[CSV Record Shift]: Cannot shift the record in the current instance." _
                        & " This is because there is no imported data or the records do not " _
                        & "have the same number of fields."
        P_ERROR_NUMBER = vbObjectError + 9025
        P_ERROR_SOURCE = "CSVinterface"
    End If
    P_VECTORS_REGULAR_BOUND = P_VECTORS_REGULAR_BOUND - 1
    P_VECTORS_MAX_BOUND = P_VECTORS_MAX_BOUND - 1
    Set ShiftRecord = Me
    Exit Function
ErrHandler_ShiftRecord:
    P_ERROR_DESC = "[CSV Record Shift]: " & err.Description
    P_ERROR_NUMBER = err.Number
    P_ERROR_SOURCE = "CSVinterface"
    Exit Function
OutOfBounds_ShiftRecord:
    err.Raise Number:=vbObjectError + 9016, _
                    Source:="CSVinterface Class", _
                    Description:="The specified index is out of bounds. Please check and try again."
    Resume ErrHandler_ShiftRecord
End Function
Public Function SniffDelimiters(ByRef confObject As CSVparserConfig, _
                            Optional ByRef CSVstring As String = vbNullString, _
                            Optional ByVal startingRecord As Long = 1, _
                            Optional Threshold As Long = 10) As CSVdialect

    If Not (confObject.path = vbNullString And CSVstring = vbNullString) Then
        If CSVstring = vbNullString Then
            Dim StreamObj As CSVTextStream

            Set StreamObj = New CSVTextStream
            With StreamObj
                .endStreamOnLineBreak = True
                .utf8EncodedFile = confObject.utf8EncodedFile
                .unifiedLFOutput = confObject.multiEndOfLineCSV
                .OpenStream confObject.path
                .ReadText
                RaiseEvent AfterLoadStream
                Set SniffDelimiters = SniffInString(confObject, .bufferString, True, startingRecord, Threshold)
                Set StreamObj = Nothing
            End With
        Else
            Set SniffDelimiters = SniffInString(confObject, CSVstring, startingRecord:=startingRecord, Threshold:=Threshold)
        End If
        ResetAlerts
    End If
End Function
Private Function SniffInString(ByRef confObject As CSVparserConfig, _
                            ByRef CSVstring As String, _
                            Optional Buffered As Boolean = False, _
                            Optional startingRecord As Long = 1, _
                            Optional Threshold As Long = 10) As CSVdialect

    Dim EmptyParam() As Variant
    Dim GuesserHelper As CSVSniffer
    Dim ImportedTable As CSVArrayList
    Dim i As Long, j As Long, k As Long
    Dim LinesEnds() As String
    Dim QuoteChar() As QuoteTokens
    Dim ScoreArray As CSVArrayList
    Dim tmpConfig As CSVparserConfig
    Dim tmpResult As String
    Dim TmpCSVstr As String
    Dim TmpDelimiters() As String

    ReDim LinesEnds(0 To 2)
    Set tmpConfig = confObject.CopyConfig
    '@--------------------------------------------------------------------------------
    'Field delimiters array
    TmpDelimiters() = tmpConfig.delimitersToGuess
    '@--------------------------------------------------------------------------------
    'LinesEnds array
    LinesEnds(0) = vbCrLf
    LinesEnds(1) = vbCr
    LinesEnds(2) = vbLf
    '@--------------------------------------------------------------------------------
    'Quotes array
    ReDim QuoteChar(0 To 2)
    QuoteChar(0) = DoubleQuotes
    QuoteChar(1) = Apostrophe
    QuoteChar(2) = Tilde
    '@--------------------------------------------------------------------------------
    'Parser config
    ReDim EmptyParam(-1 To -1) 'To emule empty ParamArray parameter
    Set SniffInString = New CSVdialect
    '@--------------------------------------------------------------------------------
    'String source
    If Not Buffered Then
        TmpCSVstr = MidB$(CSVstring, 1, 524288) 'Load 0.5 MB of data
    Else
        TmpCSVstr = CSVstring
    End If
    With tmpConfig
        '@--------------------------------------------------------------------------------
        'Guess dialect using NumberOfRecords samples records with skipping
        .startingRecord = startingRecord
        .endingRecord = startingRecord + Threshold - 1
        .skipCommentLines = True
        .skipEmptyLines = True
        Set GuesserHelper = New CSVSniffer
        Set ScoreArray = New CSVArrayList
        ScoreArray.indexing = True
        For i = LBound(TmpDelimiters) To UBound(TmpDelimiters)
                For j = LBound(LinesEnds) To UBound(LinesEnds)
                        For k = LBound(QuoteChar) To UBound(QuoteChar)
                            '@--------------------------------------------------------------------------------
                            'Set CSV dialect
                            .dialect.fieldsDelimiter = TmpDelimiters(i)
                            .dialect.recordsDelimiter = LinesEnds(j)
                            .dialect.quoteToken = QuoteChar(k)
                            Set ImportedTable = New CSVArrayList
                            ParseCSVstring TmpCSVstr, tmpConfig, ImportedTable, EmptyParam
                            '@--------------------------------------------------------------------------------
                            'Save results with keys
                            If InStrB(1, TmpCSVstr, GetQuoteChar(QuoteChar(k))) Then
                                ScoreArray.AddIndexedItem AppendIndexesToKey(DialectToString(.dialect), i, j, k), _
                                                            GuesserHelper.TableScore(ImportedTable, Threshold)
                            Else
                                ScoreArray.AddIndexedItem AppendIndexesToKey(DialectToString(.dialect) & CHR_CARET, i, j, k), _
                                                            GuesserHelper.TableScore(ImportedTable, Threshold)
                            End If
                        Next k
                Next j
        Next i
    End With
    With ScoreArray
        '@--------------------------------------------------------------------------------
        'Choose the maximum score
        tmpResult = RemoveIndexesFromKey(.keys()(MaxIndexVal(.indexedItems)))
        '@--------------------------------------------------------------------------------
        'Returns
        Set SniffInString = StringToDialect(tmpResult)
        If InStrB(1, TmpCSVstr, CHR_BACKSLASH) Then
            SniffInString.escapeMode = unix
        End If
        '@--------------------------------------------------------------------------------
        '2nd pass for unique record CSV file
        Set ImportedTable = New CSVArrayList
        Set tmpConfig.dialect = SniffInString
        ParseCSVstring TmpCSVstr, tmpConfig, ImportedTable, EmptyParam
        If ImportedTable.count = 1 Then
            If InStrB(1, TmpCSVstr, vbCrLf) Then
                SniffInString.recordsDelimiter = vbCrLf
            Else
                If InStrB(1, TmpCSVstr, vbCr) Then
                    SniffInString.recordsDelimiter = vbCr
                Else
                    If InStrB(1, TmpCSVstr, vbLf) Then
                        SniffInString.recordsDelimiter = vbLf
                    Else
                        SniffInString.recordsDelimiter = vbCrLf
                    End If
                End If
            End If
        End If
    End With
    Erase EmptyParam
    Erase LinesEnds
    Erase QuoteChar
    Erase TmpDelimiters
    Set ScoreArray = Nothing
End Function
Private Function DialectToString(ByRef dialectObj As CSVdialect) As String
    Dim tmpResult() As String
    ReDim tmpResult(0 To 2)

    With dialectObj
        tmpResult(0) = .fieldsDelimiter
        tmpResult(1) = .recordsDelimiter
        tmpResult(2) = CStr(.quoteToken)
    End With
    DialectToString = Join$(tmpResult, "ii")
End Function
Private Function AppendIndexesToKey(ByRef aKey As String, idx1 As Long, idx2 As Long, idx3 As Long) As String
    AppendIndexesToKey = CHR_LSQRB & idx1 & idx2 & idx3 & CHR_RSQRB & aKey
End Function
Private Function RemoveIndexesFromKey(ByRef aKey As Variant) As String
    RemoveIndexesFromKey = MidB(aKey, InStrB(1, aKey, CHR_RSQRB) + 2)
End Function
Private Function StringToDialect(ByRef dialectString As String) As CSVdialect
    Dim tmpArr() As String
    Dim idx As Long
    Dim tmpResult As CSVdialect

    tmpArr() = Split(dialectString, "ii")
    idx = LBound(tmpArr)
    Set tmpResult = New CSVdialect
    With tmpResult
        .fieldsDelimiter = tmpArr(idx)
        .recordsDelimiter = tmpArr(idx + 1)
        If InStrB(1, dialectString, CHR_CARET) Then
            .quoteToken = QuoteTokens.DoubleQuotes
        Else
            .quoteToken = CLng(tmpArr(idx + 2))
        End If
    End With
    Set StringToDialect = tmpResult
End Function
Private Function GetQuoteChar(ByRef QTokenCode As QuoteTokens) As String
    Select Case QTokenCode
        Case 1
            GetQuoteChar = CHR_APOSTROPHE
        Case 2
            GetQuoteChar = CHR_DOUBLE_QUOTES
        Case Else
            GetQuoteChar = CHR_TILDE
    End Select
End Function
Private Function GetQuoteToken(ByRef QTokenCode As Long) As QuoteTokens
    Select Case QTokenCode
        Case 1
            GetQuoteToken = QuoteTokens.Apostrophe
        Case 2
            GetQuoteToken = QuoteTokens.DoubleQuotes
        Case Else
            GetQuoteToken = QuoteTokens.Tilde
    End Select
End Function
Public Function Sort(Optional ByVal fromIndex As Long = -1, _
                        Optional ByVal toIndex As Long = -1, _
                        Optional ByVal sortingKeys As Variant = 1, _
                        Optional ByVal sortAlgorithm As SortingAlgorithms = SortingAlgorithms.SA_Quicksort) As CSVinterface

    On Error GoTo Sort_Error_Handler:

    If P_SUCCESSFUL_IMPORT Then
        If fromIndex = -1 Then
            If config.Headers And Not config.headersOmission Then
                fromIndex = 2
            Else
                fromIndex = 1
            End If
        End If
        If toIndex = -1 Then
            toIndex = P_CSV_DATA.count
        End If
        P_CSV_DATA.Sort fromIndex, toIndex, sortingKeys, sortAlgorithm
        RaiseEvent AfterSort
        Set Sort = Me
    End If
    Exit Function
Sort_Error_Handler:
    P_ERROR_NUMBER = err.Number
    P_ERROR_DESC = "[Sort] " & err.Description
    P_ERROR_SOURCE = err.Source
End Function
Public Function SortByField(Optional ByVal fromIndex As Long = -1, _
                            Optional ByVal toIndex As Long = -1, _
                            Optional ByVal SortingKey As Long = 1, _
                            Optional ByVal sortAlgorithm As SortingAlgorithms = SortingAlgorithms.SA_Quicksort) As CSVinterface

    On Error GoTo SortByField_Error_Handler:

    If P_SUCCESSFUL_IMPORT Then
        P_CSV_DATA.SortByField fromIndex, toIndex, SortingKey, sortAlgorithm
        RaiseEvent AfterSort
        Set SortByField = Me
    End If
    Exit Function
SortByField_Error_Handler:
    P_ERROR_NUMBER = err.Number
    P_ERROR_DESC = "[SortByField] " & err.Description
    P_ERROR_SOURCE = err.Source
End Function
Public Function SplitField(aIndex As Long, CharToSplitWith As String, _
                            Optional RowSplit As Boolean = False) As CSVinterface
    Dim FldDiff As Long

    On Error GoTo ErrHandler_SplitField
    If P_SUCCESSFUL_IMPORT Then
        If Not RowSplit Then
            If Not P_VARYING_LENGTHS Then
                Select Case aIndex
                    Case 0 To P_VECTORS_REGULAR_BOUND
                        P_CSV_DATA.SplitField aIndex, CharToSplitWith
                    Case Else
                        GoTo OutOfBounds_SplitField
                End Select
            End If
            FldDiff = UBound(P_CSV_DATA(0)) - P_VECTORS_REGULAR_BOUND
            P_VECTORS_REGULAR_BOUND = P_VECTORS_REGULAR_BOUND + FldDiff
            P_VECTORS_MAX_BOUND = P_VECTORS_MAX_BOUND + FldDiff
            Set SplitField = Me
        Else
            Select Case aIndex
                Case 0 To P_VECTORS_REGULAR_BOUND
                    P_CSV_DATA.SplitField aIndex, CharToSplitWith, RowSplit
                Case Else
                    GoTo OutOfBounds_SplitField
            End Select
        End If
    Else
        P_ERROR_DESC = "[CSV Field Split]: Cannot split the field in the current instance." _
                        & " This is because there is no imported data or the records do not " _
                        & "have the same number of fields."
        P_ERROR_NUMBER = vbObjectError + 9023
        P_ERROR_SOURCE = "CSVinterface"
    End If
    Exit Function
ErrHandler_SplitField:
    P_ERROR_DESC = "[CSV Field Split]: " & err.Description
    P_ERROR_NUMBER = err.Number
    P_ERROR_SOURCE = "CSVinterface"
    Exit Function
OutOfBounds_SplitField:
    err.Raise Number:=vbObjectError + 9016, _
                    Source:="CSVinterface Class", _
                    Description:="The specified index is out of bounds. Please check and try again."
    Resume ErrHandler_SplitField
End Function
Private Sub StreamParseCSV(configObj As CSVparserConfig, _
                            outputList As CSVArrayList, _
                            ByRef FilterColumns() As Variant)
    Dim commToken As Long
    Dim dynamicType As Boolean
    Dim EndOF As Boolean
    Dim EndOFStreamBuffer As Boolean
    Dim EndRecord As Long
    Dim EscapeChrCount As Long
    Dim EscapedBRS As Boolean
    Dim EscapedWFS As Boolean
    Dim fieldDelimiter As String
    Dim FieldID As Long
    Dim hpointer As Long
    Dim IgnoreCommentLines As Boolean
    Dim IgnoreLines As Boolean
    Dim ImportDepth As Long
    Dim ImportSwitch As Boolean
    Dim IncludeRecord As Boolean
    Dim IsHeader As Boolean
    Dim IsWellEscapedField As Boolean
    Dim KnownEnding As Boolean
    Dim MaxReqIndex As Long
    Dim NotEvenEscapeChrNumber As Boolean
    Dim OpenTokenArrIdx As Long
    Dim OpenTokenStartPos As Long
    Dim OpenTokenStreamIdx As Long
    Dim RecordEndReached As Boolean
    Dim RecordsCount As Long
    Dim RecordToken As CSVArrayList
    Dim SavedRecords As Long
    Dim SearchBeginningMark As Long
    Dim SplittedToken As Boolean
    Dim StartRecord As Long
    Dim StreamsCounter As Long
    Dim tmpRequested() As Variant
    Dim tmpToken As CSVArrayList
    Dim TokenBeginningPos As Long
    Dim TokenEndingPos As Long
    Dim TokenEndReached As Boolean
    
    On Error GoTo StreamParseCSV_Error_Handler
    '@----------------------------------------------------------------------------
    'Start variables
    SearchBeginningMark = 1&
    EscapeChrCount = 0&
    RecordsCount = 0&
    SavedRecords = 0&
    P_VECTORS_REGULAR_BOUND = 0&
    P_VECTORS_MAX_BOUND = 0&
    Set tmpToken = New CSVArrayList
    Set CSVstream = New CSVTextStream
    With configObj
        CSVstream.bufferSize = .bufferSize
        CSVstream.endStreamOnLineBreak = True
        CSVstream.linebreakMatchingBehavior = EndLineMatchingBehavior.OnlyBackwardSense
        CSVstream.unifiedLFOutput = .multiEndOfLineCSV
        CSVstream.utf8EncodedFile = .utf8EncodedFile
        CSVstream.autoDetectEncoding = .autoDetectEncoding
        CSVstream.OpenStream .path
        commToken = AscW(.commentsToken)
        dynamicType = .dynamicTyping
        dTTemplate() = .dTypingTemplate
        dTTargets() = .dTypingLinks
        IgnoreCommentLines = .skipCommentLines
        IgnoreEmptyLines = .skipEmptyLines
        IgnoreLines = (IgnoreCommentLines Or IgnoreEmptyLines)
        TemplateDefined = .dTTemplateDefined
        TargetsDefined = .dTTemplateLinksDefined
        fieldDelimiter = .dialect.fieldsDelimiter
        RecordDelimiter = .dialect.recordsDelimiter
        EndRecord = .endingRecord
        StartRecord = .startingRecord
        UseUnixEscapeSeq = (.dialect.escapeMode = unix)
    End With
    TokenEndReached = False
    RecordEndReached = False
    tmpRequested() = FilterColumns
    Select Case UBound(tmpRequested)
        Case -1
            ReDim RequestedFieldsArray(0)
            RequestedFieldsArray(0) = -1
        Case Else
            RequestedFieldsArray() = GetHeaderIndexes(tmpRequested, P_CSV_HEADER)
    End Select
    '@----------------------------------------------------------------------------
    'Set variables
    OverStringPointer = SearchBeginningMark
    MaxReqIndex = UBound(RequestedFieldsArray)
    '@----------------------------------------------------------------------------
    'Check if user want Import all records
    KnownEnding = True
    Select Case EndRecord
        Case 0&
            If StartRecord >= 1 Then
                KnownEnding = False
            Else
                ImportDepth = 1&
            End If
        Case Is < StartRecord
            ImportDepth = 1&
        Case Else
            ImportDepth = EndRecord - StartRecord + 1&
    End Select
    If StartRecord = 1& Then
        If configObj.headersOmission And configObj.Headers Then
            If ImportDepth > 1& Or EndRecord = 0& Then
                StartRecord = 2&
                ImportDepth = ImportDepth - 1&
            End If
        End If
    End If
    '@----------------------------------------------------------------------------
    'Set Coerce Type to fields
    Select Case configObj.dialect.quoteToken
        Case 1
            QuoteChar = CHR_APOSTROPHE
        Case 2
            QuoteChar = CHR_DOUBLE_QUOTES
        Case Else
            QuoteChar = CHR_TILDE
    End Select
    '@----------------------------------------------------------------------------
    'Set the escape sequences
    UnixEscapeSeq = CHR_BACKSLASH & QuoteChar
    EscapeSequence = QuoteChar & QuoteChar
    EscapeSequence_Delim = CHR_BACKSLASH & fieldDelimiter
    '@----------------------------------------------------------------------------
    EscapeAscW = AscW(QuoteChar)
    FDAscW = AscW(fieldDelimiter)
    Set RecordToken = New CSVArrayList
    index = 0
    '@----------------------------------------------------------------------------
    'Get CSV stream and populate a temp array
    With CSVstream
        .ReadText
        RaiseEvent AfterLoadStream
        StreamWithQuotes = InStrB(1, .bufferString, QuoteChar)
        StreamEnd = .atEndOfStream
    End With
    '@----------------------------------------------------------------------------
    'Ensuring the stream is sufficient enough to hold at least one CSV line
    With CSVstream
        Do While Not StreamEnd And InStrB(1, .bufferString, RecordDelimiter) = 0
            configObj.bufferSize = 2 * configObj.bufferSize
            .bufferSize = configObj.bufferSize
            .RestartPointer
            .ReadText
            RaiseEvent AfterLoadStream
            StreamWithQuotes = InStrB(1, .bufferString, QuoteChar)
            StreamEnd = .atEndOfStream
        Loop
        StreamsCounter = StreamsCounter + 1
        If UseUnixEscapeSeq Then
            tmpCSV() = Split(UnixToStandardEscapeSeq(.bufferString, UnixEscapeSeq, EscapeSequence), RecordDelimiter)
        Else
            tmpCSV() = Split(.bufferString, RecordDelimiter)
        End If
    End With
    MaxIndex = UBound(tmpCSV)
    If MaxIndex = -1 Then outputList.Add2 vbNullString: outputList.ShrinkBuffer: Exit Sub
    If Not StreamEnd Then
        If tmpCSV(MaxIndex) = vbNullString Then 'Streams lefts empty value when splited
            MaxIndex = MaxIndex - 1
        End If
    Else
        If IgnoreEmptyLines Then
            If tmpCSV(MaxIndex) = vbNullString Then 'Skip empty line
                MaxIndex = MaxIndex - 1
            End If
        End If
    End If
    If IgnoreLines Then
        '@----------------------------------------------------------------------------
        'Skip lines if needed
        StreamSkipUnwantedLines index, CSVstream, tmpCSV, MaxIndex, _
                                RecordDelimiter, commToken, StreamWithQuotes, _
                                StreamsCounter, IgnoreCommentLines, IgnoreEmptyLines
    End If
    EndOFStreamBuffer = (index > MaxIndex)
    If EndOFStreamBuffer Then 'The file has no significant data
        GoTo StreamParseCSV_NoSignificantData
    End If
    '@----------------------------------------------------------------------------
    'Loop through CSV using streams
    Do
        IsHeader = (StartRecord = 1 And RecordsCount = 0)
        LenCurrentIndex = LenB(tmpCSV(index))
        '@----------------------------------------------------------------------------
        'Restart pointers
        RecordEndReached = False
        FieldID = 0
        IncludeRecord = (RecordsCount + 1& >= StartRecord)
        Do
            '@----------------------------------------------------------------------------
            'Process the current Token
            TokenEndReached = False
            TokenBeginningPos = OverStringPointer
            If StreamWithQuotes Then
                Select Case OverStringPointer
                    Case Is < LenCurrentIndex
                        ASCIIcharw = AscW(MidB$(tmpCSV(index), OverStringPointer, 2))
                        '@----------------------------------------------------------------------------
                        'Ignore spaces & tabs
                        SkipSpacesAndTabs
                        OpenedToken = (ASCIIcharw = EscapeAscW)
                    Case Else 'Avoid over runs
                        OpenedToken = False
                End Select
            Else
                OpenedToken = False
            End If
            If OpenedToken Then
                If OpenTokenStreamIdx = 0 Then
                    OpenTokenStreamIdx = StreamsCounter
                End If
                If OpenTokenArrIdx = 0 Then
                    OpenTokenArrIdx = index
                End If
                '@----------------------------------------------------------------------------
                'Try to escape the Token
                EscapeChrCount = EscapeChrCount + 1&
                TokenBeginningPos = OverStringPointer + 2&
                SearchBeginningMark = TokenBeginningPos
                hpointer = InStrB(SearchBeginningMark, tmpCSV(index), QuoteChar) 'EOFld marker
                If hpointer Then EscapeChrCount = EscapeChrCount + 1&
                Do While (Not TokenEndReached And Not EndOF)
                    '@----------------------------------------------------------------------------
                    'Buffering if needed
                    If Not SplittedToken Then SplittedToken = (hpointer = 0)
                    Do
                        Do While hpointer = 0& 'Quote char may be on subsequent Index
                            If IncludeRecord Then
                                Select Case SearchBeginningMark
                                    Case 1 'Store the full Token from the current Index
                                        tmpToken.Add tmpCSV(index)
                                    Case Else 'Store the Right most Token portion from the current Index
                                        tmpToken.Add RightB$(tmpCSV(index), LenCurrentIndex - SearchBeginningMark + 1)
                                End Select
                                If OpenTokenStartPos = 0 Then
                                    OpenTokenStartPos = SearchBeginningMark
                                End If
                            End If
                            SplittedToken = True
                            index = index + 1&
                            EndOFStreamBuffer = (index > MaxIndex)
                            If Not EndOFStreamBuffer Then
                                LenCurrentIndex = LenB(tmpCSV(index))
                                SearchBeginningMark = 1
                                hpointer = InStrB(SearchBeginningMark, tmpCSV(index), QuoteChar) 'Find quote char
                                If hpointer Then EscapeChrCount = EscapeChrCount + 1&
                            Else
                                If Not StreamEnd Then
                                    '@----------------------------------------------------------------------------
                                    'Load new stream
                                    LoadNewStream
                                    StreamsCounter = StreamsCounter + 1
                                    EndOFStreamBuffer = False
                                    LenCurrentIndex = LenB(tmpCSV(index))
                                    SearchBeginningMark = 1
                                    hpointer = InStrB(SearchBeginningMark, tmpCSV(index), QuoteChar) 'Find quote char
                                    If hpointer Then EscapeChrCount = EscapeChrCount + 1&
                                Else
                                    hpointer = 0&
                                End If
                            End If
                            EndOF = (StreamEnd And EndOFStreamBuffer)
                            If hpointer = 0& And EndOF Then 'Missing quote char
                                P_ERROR_DESC = P_ERROR_DESC & "Missing some quote char. The parser retrieved the record in recovery mode. " & _
                                               "[Review the record #" & RecordsCount + 1 & ", field #" & RecordToken.count + 1 _
                                                 & " on the source CSV file]."
                                RecordEndReached = True
                                OverStringPointer = 1&
                                StreamsCounter = 0
                                With CSVstream
                                    .RestartPointer
                                    Do While StreamsCounter < OpenTokenStreamIdx
                                        LoadNewStream
                                        StreamsCounter = StreamsCounter + 1
                                        RaiseEvent AfterLoadStream
                                    Loop
                                End With
                                tmpToken.Clear
                                LenCurrentIndex = LenB(tmpCSV(OpenTokenArrIdx))
                                If OpenTokenStartPos = 1 Then 'Save the entire record exclusively in recovery mode
                                    RecordToken.Add MidB$(tmpCSV(OpenTokenArrIdx), 3)
                                Else 'Save a portion of the record in recovery mode
                                    RecordToken.Add RightB$(tmpCSV(OpenTokenArrIdx), LenCurrentIndex - OpenTokenStartPos + 1)
                                End If
                                index = OpenTokenArrIdx + 1
                                GoTo Restore_From_Missing_EscapeChar
                            End If
                        Loop
                        OverStringPointer = hpointer + 2&
                        NotEvenEscapeChrNumber = (EscapeChrCount And 1)
                        If OverStringPointer < LenCurrentIndex Then
                            ASCIIcharw = AscW(MidB$(tmpCSV(index), OverStringPointer, 2))
                            '@----------------------------------------------------------------------------
                            'Ignore spaces & tabs
                            SkipSpacesAndTabs
                            '@----------------------------------------------------------------------------
                            'Check if we are at the end of the field
                            If UseUnixEscapeSeq Then
                                EscapedWFS = (ASCIIcharw = FDAscW) And (Not NotEvenEscapeChrNumber) And _
                                                (Not (AscW(MidB$(tmpCSV(index), OverStringPointer - 2, 2)) = ASCW_CHR_BACKSLASH))
                            Else
                                EscapedWFS = (ASCIIcharw = FDAscW) And (Not NotEvenEscapeChrNumber)
                            End If
                        Else
                            EscapedWFS = False
                            EscapedBRS = Not NotEvenEscapeChrNumber
                        End If
                        IsWellEscapedField = (EscapedWFS Or EscapedBRS)
                        If Not IsWellEscapedField Then
                            hpointer = InStrB(hpointer + 2&, tmpCSV(index), QuoteChar) 'EOFld marker
                            If hpointer Then EscapeChrCount = EscapeChrCount + 1&
                        End If
                    Loop While Not IsWellEscapedField
                    If EscapedWFS Then 'Quote char succeeded by fields separator char
                        TokenEndingPos = hpointer - 1&
                        TokenEndReached = True
                        EscapeChrCount = 0&
                        If IncludeRecord Then
                            '@----------------------------------------------------------------------------
                            'Save the Token
                            Select Case SplittedToken
                                Case False 'Just save Token
                                    If RequestedField(FieldID + 1&, MaxReqIndex) Then 'Filter fields
                                        Select Case dynamicType
                                            Case False
                                                RecordToken.Add UnescapeField(MidB$(tmpCSV(index), TokenBeginningPos, _
                                                                                    TokenEndingPos - TokenBeginningPos + 1), _
                                                                            QuoteChar, EscapeSequence)
                                            Case Else
                                                RecordToken.Add TypeData(UnescapeField(MidB$(tmpCSV(index), TokenBeginningPos, _
                                                                                    TokenEndingPos - TokenBeginningPos + 1), _
                                                                            QuoteChar, EscapeSequence), _
                                                                        FieldID, IsHeader)
                                        End Select
                                    End If
                                Case Else 'The Token is stored on tmpToken object
                                    tmpToken.Add LeftB$(tmpCSV(index), TokenEndingPos)
                                    If RequestedField(FieldID + 1&, MaxReqIndex) Then 'Filter fields
                                        RecordToken.Add UnescapeField(Join$(tmpToken.items, RecordDelimiter), _
                                                                    QuoteChar, EscapeSequence) 'Save whole Token
                                    End If
                                    tmpToken.Clear 'Clear tmpToken content
                            End Select
                        End If
                        SplittedToken = False 'Restart Token status
                        OverStringPointer = OverStringPointer + 2& 'Advance over the field separator char
                        RecordEndReached = False
                    Else
                        If EscapedBRS Then  'Quote char succeeded
                                            'by records separator char
                            TokenEndingPos = hpointer - 1&
                            TokenEndReached = True
                            EscapeChrCount = 0&
                            RecordEndReached = True 'EOR
                            If IncludeRecord Then
                                '@----------------------------------------------------------------------------
                                'Save the Token
                                Select Case SplittedToken
                                    Case False 'Just save Token
                                        If RequestedField(FieldID + 1&, MaxReqIndex) Then 'Filter fields
                                            Select Case dynamicType
                                                Case False
                                                    RecordToken.Add UnescapeField(MidB$(tmpCSV(index), TokenBeginningPos, _
                                                                                    TokenEndingPos - TokenBeginningPos + 1), _
                                                                                QuoteChar, EscapeSequence)
                                                Case Else
                                                    RecordToken.Add TypeData(UnescapeField(MidB$(tmpCSV(index), TokenBeginningPos, _
                                                                                            TokenEndingPos - TokenBeginningPos + 1), _
                                                                                    QuoteChar, EscapeSequence), _
                                                                            FieldID, IsHeader)
                                            End Select
                                        End If
                                    Case Else 'The Token is stored on tmpToken array
                                        tmpToken.Add LeftB$(tmpCSV(index), OverStringPointer - 3)
                                        If RequestedField(FieldID + 1&, MaxReqIndex) Then 'Filter fields
                                            RecordToken.Add UnescapeField(Join$(tmpToken.items, RecordDelimiter), _
                                                                        QuoteChar, EscapeSequence) 'Save whole Token
                                        End If
                                        tmpToken.Clear 'Clear tmpToken content
                                End Select
                            End If
                            SplittedToken = False 'Restart Token status
                            index = index + 1&
                            EndOFStreamBuffer = (index > MaxIndex)
                            EndOF = (StreamEnd And EndOFStreamBuffer)
                            If EndOFStreamBuffer And Not EndOF Then
                                '@----------------------------------------------
                                'Advance stream
                                LoadNewStream
                                StreamsCounter = StreamsCounter + 1
                                EndOFStreamBuffer = False
                            End If
                            OverStringPointer = 1&
                        Else
                            If IncludeRecord Then
                                If SplittedToken Then 'Store the Left most Token portion from the current Index
                                    tmpToken.Add LeftB$(tmpCSV(index), OverStringPointer - 1)
                                End If
                            End If
                            SearchBeginningMark = OverStringPointer
                            hpointer = InStrB(SearchBeginningMark, tmpCSV(index), QuoteChar) 'Find quote char
                            If hpointer Then EscapeChrCount = EscapeChrCount + 1&
                            TokenEndReached = False
                        End If
                    End If
                Loop
Restore_From_Missing_EscapeChar:
                OpenTokenStreamIdx = 0
                OpenTokenStartPos = 0
            Else
                OpenTokenArrIdx = 0
                '@----------------------------------------------------------------------------
                'Try to mark the Starting and Ending position for the field
                hpointer = InStrB(OverStringPointer, tmpCSV(index), fieldDelimiter)
                If UseUnixEscapeSeq Then
                    Do While hpointer
                        If hpointer > 2 Then
                            UnixEscpedSep = (AscW(MidB$(tmpCSV(index), hpointer - 2, 2)) = ASCW_CHR_BACKSLASH)
                        Else
                            UnixEscpedSep = False
                        End If
                        If UnixEscpedSep Then
                            hpointer = InStrB(hpointer + 1, tmpCSV(index), fieldDelimiter)
                        Else
                            Exit Do
                        End If
                    Loop
                End If
                If hpointer Then 'Found fields delimiter char
                    SplittedToken = False
                    TokenEndReached = True
                    TokenEndingPos = hpointer - 1&
                    OverStringPointer = hpointer + 2&
                    SearchBeginningMark = OverStringPointer
                    RecordEndReached = False
                    If IncludeRecord Then
                        '@----------------------------------------------------------------------------
                        'Save the Token
                        If RequestedField(FieldID + 1&, MaxReqIndex) Then 'Filter fields
                            Select Case dynamicType
                                Case False
                                    If UseUnixEscapeSeq Then
                                        RecordToken.Add UnescapeField(UnescapeField(MidB$(tmpCSV(index), TokenBeginningPos, _
                                                                        TokenEndingPos - TokenBeginningPos + 1), _
                                                                fieldDelimiter, EscapeSequence_Delim), _
                                                                        QuoteChar, EscapeSequence)
                                    Else
                                        RecordToken.Add MidB$(tmpCSV(index), TokenBeginningPos, _
                                                                TokenEndingPos - TokenBeginningPos + 1)
                                    End If
                                Case Else
                                    If UseUnixEscapeSeq Then
                                        RecordToken.Add TypeData(UnescapeField(UnescapeField(MidB$(tmpCSV(index), TokenBeginningPos, _
                                                                            TokenEndingPos - TokenBeginningPos + 1), _
                                                                        fieldDelimiter, EscapeSequence_Delim), _
                                                                    QuoteChar, EscapeSequence), _
                                                            FieldID, IsHeader)
                                    Else
                                        RecordToken.Add TypeData(MidB$(tmpCSV(index), TokenBeginningPos, _
                                                                    TokenEndingPos - TokenBeginningPos + 1), _
                                                            FieldID, IsHeader)
                                    End If
                            End Select
                        End If
                    End If
                Else 'Missing fields delimiter char
                    If UseUnixEscapeSeq Then
                        Do While Not EndOF And (AscW(RightB$(tmpCSV(index), 2)) = ASCW_CHR_BACKSLASH) 'Unix escaped Record delimiter
                            SplittedToken = True
                            If IncludeRecord Then
                                Select Case SearchBeginningMark
                                    Case 1 'Store the full Token from the current Index
                                        tmpToken.Add MidB$(tmpCSV(index), 1, LenCurrentIndex - 1)
                                    Case Else 'Store the Right most Token portion from the current Index
                                        tmpToken.Add MidB$(tmpCSV(index), TokenBeginningPos, LenCurrentIndex - TokenBeginningPos - 1)
                                End Select
                            End If
                            tmpCSV(index) = vbNullString
                            index = index + 1&
                            OverStringPointer = 1&
                            EndOFStreamBuffer = (index > MaxIndex)
                            If Not EndOFStreamBuffer Then
                                LenCurrentIndex = LenB(tmpCSV(index))
                                SearchBeginningMark = 1
                                '@----------------------------------------------------------------------------
                                'Try to mark the Starting and Ending position for the field
                                hpointer = InStrB(OverStringPointer, tmpCSV(index), fieldDelimiter)
                                Do While hpointer
                                    If hpointer > 2 Then
                                        UnixEscpedSep = (AscW(MidB$(tmpCSV(index), hpointer - 2, 2)) = ASCW_CHR_BACKSLASH)
                                    Else
                                        UnixEscpedSep = False
                                    End If
                                    If UnixEscpedSep Then
                                        hpointer = InStrB(hpointer + 1, tmpCSV(index), fieldDelimiter)
                                    Else
                                        Exit Do
                                    End If
                                Loop
                            Else
                                If Not StreamEnd Then
                                    '@----------------------------------------------------------------------------
                                    'Load new stream
                                    LoadNewStream
                                    StreamsCounter = StreamsCounter + 1
                                    EndOFStreamBuffer = False
                                    LenCurrentIndex = LenB(tmpCSV(index))
                                    SearchBeginningMark = 1
                                    OverStringPointer = SearchBeginningMark
                                    '@----------------------------------------------------------------------------
                                    'Try to mark the Starting and Ending position for the field
                                    hpointer = InStrB(OverStringPointer, tmpCSV(index), fieldDelimiter)
                                    Do While hpointer
                                        If hpointer > 2 Then
                                            UnixEscpedSep = (AscW(MidB$(tmpCSV(index), hpointer - 2, 2)) = ASCW_CHR_BACKSLASH)
                                        Else
                                            UnixEscpedSep = False
                                        End If
                                        If UnixEscpedSep Then
                                            hpointer = InStrB(hpointer + 1, tmpCSV(index), fieldDelimiter)
                                        Else
                                            Exit Do
                                        End If
                                    Loop
                                Else
                                    hpointer = 0&
                                End If
                            End If
                            EndOF = (StreamEnd And EndOFStreamBuffer)
                            If hpointer Then 'Unescaped fields delimiter found
                                TokenEndReached = True
                                TokenEndingPos = hpointer - 1&
                                OverStringPointer = hpointer + 2&
                                SearchBeginningMark = OverStringPointer
                                RecordEndReached = False
                                tmpToken.Add LeftB$(tmpCSV(index), TokenEndingPos) 'Store the Left most Token portion from the current Index
                                Exit Do
                            Else
                                If Not UnixEscpedSep Then 'End of record
                                    If IncludeRecord Then
                                        Select Case SearchBeginningMark
                                            Case 1 'Store the full Token from the current Index
                                                tmpToken.Add MidB$(tmpCSV(index), 1, LenCurrentIndex)
                                            Case Else 'Store the Right most Token portion from the current Index
                                                tmpToken.Add MidB$(tmpCSV(index), TokenBeginningPos, LenCurrentIndex - TokenBeginningPos)
                                        End Select
                                    End If
                                End If
                            End If
                        Loop
                    Else
                        hpointer = 0&
                    End If
                    If hpointer = 0 Then
                        TokenEndReached = True
                        RecordEndReached = True
                    End If
                    If IncludeRecord Then
                        '@----------------------------------------------------------------------------
                        'Save the Token
                        Select Case SplittedToken
                            Case False 'Just save Token
                                If RequestedField(FieldID + 1&, MaxReqIndex) Then 'Filter fields
                                    Select Case dynamicType
                                        Case False
                                            If UseUnixEscapeSeq Then
                                                RecordToken.Add UnescapeField(UnescapeField(RightB$(tmpCSV(index), LenCurrentIndex _
                                                                            - TokenBeginningPos + 1), _
                                                                        fieldDelimiter, EscapeSequence_Delim), _
                                                                        QuoteChar, EscapeSequence)
                                            Else
                                                RecordToken.Add RightB$(tmpCSV(index), LenCurrentIndex _
                                                                        - TokenBeginningPos + 1)
                                            End If
                                        Case Else
                                            If UseUnixEscapeSeq Then
                                                RecordToken.Add TypeData(UnescapeField(UnescapeField(RightB$(tmpCSV(index), LenCurrentIndex _
                                                                                    - TokenBeginningPos + 1), _
                                                                                    fieldDelimiter, EscapeSequence_Delim), _
                                                                                QuoteChar, EscapeSequence), _
                                                                        FieldID, IsHeader)
                                            Else
                                                RecordToken.Add TypeData(RightB$(tmpCSV(index), LenCurrentIndex _
                                                                                - TokenBeginningPos + 1), _
                                                                        FieldID, IsHeader)
                                            End If
                                    End Select
                                End If
                            Case Else 'The Token is stored on tmpToken array
                                If RequestedField(FieldID + 1&, MaxReqIndex) Then 'Filter fields
                                    If UseUnixEscapeSeq Then
                                        RecordToken.Add UnescapeField(UnescapeField(Join$(tmpToken.items, RecordDelimiter), _
                                                                fieldDelimiter, EscapeSequence_Delim), _
                                                                        QuoteChar, EscapeSequence) 'Save whole Token
                                    Else
                                        RecordToken.Add Join$(tmpToken.items, RecordDelimiter)
                                    End If
                                End If
                                tmpToken.Clear 'Clear tmpToken content
                                SplittedToken = False 'Restart Token status
                        End Select
                    End If
                    If RecordEndReached Then
                        index = index + 1&
                        EndOFStreamBuffer = (index > MaxIndex)
                        EndOF = (StreamEnd And EndOFStreamBuffer)
                        If EndOFStreamBuffer And Not EndOF Then
                            '@----------------------------------------------
                            'Advance stream
                            LoadNewStream
                            StreamsCounter = StreamsCounter + 1
                            EndOFStreamBuffer = False
                        End If
                        OverStringPointer = 1&
                    End If
                End If
            End If
            FieldID = FieldID + 1&
        Loop While (Not RecordEndReached And Not EndOF)
        Select Case P_VECTORS_REGULAR_BOUND
            Case Is > 0
                If RecordToken.count - 1 > P_VECTORS_MAX_BOUND Then
                    P_VECTORS_MAX_BOUND = RecordToken.count - 1
                    RecordToken.ShrinkBuffer
                End If
            Case Else
                Select Case FieldID
                    Case Is > 0
                        P_VECTORS_REGULAR_BOUND = RecordToken.count - 1
                        RecordToken.ShrinkBuffer
                    Case Else
                        P_VECTORS_REGULAR_BOUND = FieldID
                End Select
                P_VECTORS_MAX_BOUND = P_VECTORS_REGULAR_BOUND
        End Select
        RecordsCount = RecordsCount + 1&
        EscapedBRS = False
        '@----------------------------------------------------------------------------
        'Start saving on given record
        If IncludeRecord Then
            outputList.Add RecordToken.items
            SavedRecords = SavedRecords + 1&
            RaiseEvent AfterImportRecord
            RecordToken.Reinitialize P_VECTORS_REGULAR_BOUND + 1
        End If
        P_VARYING_LENGTHS = (P_VECTORS_REGULAR_BOUND <> (FieldID - 1))
        If IgnoreLines Then
            '@----------------------------------------------------------------------------
            'Skip lines if needed
            StreamSkipUnwantedLines index, CSVstream, tmpCSV, MaxIndex, _
                                    RecordDelimiter, commToken, StreamWithQuotes, _
                                    StreamsCounter, IgnoreCommentLines, IgnoreEmptyLines
        End If
        EndOFStreamBuffer = (index > MaxIndex)
        EndOF = (StreamEnd And EndOFStreamBuffer)
        If EndOFStreamBuffer And Not EndOF Then
            '@----------------------------------------------
            'Advance stream
            LoadNewStream
            StreamsCounter = StreamsCounter + 1
            EndOFStreamBuffer = False
        End If
        Select Case KnownEnding
            Case False
                ImportSwitch = EndOF
            Case Else
                ImportSwitch = (SavedRecords = ImportDepth) Or EndOF
        End Select
    Loop While Not ImportSwitch
    '@----------------------------------------------------------------------------
    'Check if the array need to be shrinked at the output stage
    outputList.ShrinkBuffer
    Erase tmpCSV
    Set CSVstream = Nothing
    Set RecordToken = Nothing
    Set tmpToken = Nothing
    Exit Sub
StreamParseCSV_Error_Handler:
    Set outputList = Nothing
    Set tmpToken = Nothing
    Set CSVstream = Nothing
    Erase tmpCSV
    P_ERROR_NUMBER = err.Number
    P_ERROR_DESC = err.Description & " [Review the record #" & RecordsCount + 1 _
                                  & ", field #" & RecordToken.count + 1 _
                                   & " on the source CSV file]."
    Set RecordToken = Nothing
    P_ERROR_SOURCE = err.Source
    Exit Sub
StreamParseCSV_MissingEscapeChar:
    err.Raise Number:=vbObjectError + 9006, _
                    Source:="CSVinterface Class [StreamParseCSV]", _
                    Description:="Missing some quote char. Check the data and try again."
    Resume StreamParseCSV_Error_Handler
StreamParseCSV_NoSignificantData:
    err.Raise Number:=vbObjectError + 9010, _
                    Source:="CSVinterface Class [StreamParseCSV]", _
                    Description:="The CSV file has no significant data. This can occur when " & _
                                    "the file has only empty or commented lines that can be omitted."
    Resume StreamParseCSV_Error_Handler
End Sub
Private Sub StreamSkipUnwantedLines(ByRef idx As Long, _
                                    ByRef oStream As CSVTextStream, _
                                    ByRef arr() As String, _
                                    ByRef MaxIdx As Long, _
                                    ByRef recDelimiter As String, _
                                    ByVal CommentToken As Long, _
                                    ByRef QuotedStreamVariable As Boolean, _
                                    ByRef StreamsCounter As Long, _
                                    Optional SkipComments As Boolean = True, _
                                    Optional skipEmptyLines As Boolean = True)
    Dim CurLength As Long
    Dim charCode As Long
    Dim UnwantedLine As Boolean

start:
    '@----------------------------------------------------------------------------
    'Skip commented and blank lines if needed
    If idx <= MaxIdx Then
        Do
            CurLength = LenB(arr(idx))
            UnwantedLine = False
            If CurLength = 0 Then  'Empty line found
                If skipEmptyLines Then
                    UnwantedLine = True
                End If
            Else
                charCode = AscW(arr(idx))
                If charCode = CommentToken Then 'Commented line found
                    If SkipComments Then
                        UnwantedLine = True
                    End If
                End If
            End If
AdvanceLine:
            If UnwantedLine Then
                idx = idx + 1&
                If idx > MaxIdx Then
                    If Not oStream.atEndOfStream Then
                        '@----------------------------------------------
                        'Advance stream
                        With oStream
                            .ReadText
                            StreamsCounter = StreamsCounter + 1
                            RaiseEvent AfterLoadStream
                            QuotedStreamVariable = InStrB(1, .bufferString, QuoteChar)
                            If UseUnixEscapeSeq Then
                                tmpCSV() = Split(UnixToStandardEscapeSeq(.bufferString, UnixEscapeSeq, _
                                                EscapeSequence), RecordDelimiter)
                            Else
                                tmpCSV() = Split(.bufferString, RecordDelimiter)
                            End If
                        End With
                        arr() = Split(oStream.bufferString, recDelimiter)
                        idx = 0
                        MaxIdx = UBound(arr)
                        If arr(MaxIdx) = vbNullString Then
                            MaxIdx = MaxIdx - 1
                        End If
                    End If
                End If
            End If
        Loop While UnwantedLine And idx <= MaxIdx
    Else
        If Not oStream.atEndOfStream Then
            '@----------------------------------------------
            'Advance stream
            oStream.ReadText
            StreamsCounter = StreamsCounter + 1
            RaiseEvent AfterLoadStream
            QuotedStreamVariable = InStrB(1, oStream.bufferString, CHR_DOUBLE_QUOTES)
            arr() = Split(oStream.bufferString, recDelimiter)
            idx = 0
            MaxIdx = UBound(arr)
            If arr(MaxIdx) = vbNullString Then
                MaxIdx = MaxIdx - 1
            End If
            GoTo start
        End If
    End If
End Sub
Private Function TypeData(ByRef data As Variant, _
                            ByVal fieldIdx As Long, _
                            ByVal IsHeader As Boolean) As Variant
    If TemplateDefined Then
        If TargetsDefined Then
            If Not IsHeader Then
                Dim typingIdx As Long, maxTypingIdx As Long
                Dim typingNeed As Boolean

                typingIdx = LBound(dTTargets)
                maxTypingIdx = UBound(dTTargets)
                Do
                    typingNeed = (fieldIdx + 1 = dTTargets(typingIdx)) 'fieldIdx is base 0
                    typingIdx = typingIdx + 1
                Loop While typingIdx <= maxTypingIdx And Not typingNeed
                If typingNeed Then
                    Select Case dTTemplate(typingIdx - 1)
                        Case TypeConversion.None 'Strings
                            TypeData = data
                        Case TypeConversion.ToLong
                            TypeData = CLng(data)
                        Case TypeConversion.ToDouble
                            TypeData = CDbl(data)
                        Case TypeConversion.ToDate
                            TypeData = CDate(data)
                        Case TypeConversion.ToBoolean
                            TypeData = CBool(data)
                        Case Else
                            TypeData = data
                    End Select
                Else
                    TypeData = data
                End If
            Else 'Not type headers
                TypeData = data
            End If
        End If
    End If
End Function
Private Function UnescapeField(ByRef FieldString As String, _
                            ByRef EscapeCharacter As String, _
                            ByRef EscapeSeq As String) As String

    If InStrB(1, FieldString, EscapeSeq) Then
        UnescapeField = Join$(Split(FieldString, EscapeSeq), EscapeCharacter)
    Else
        UnescapeField = FieldString
    End If
End Function
Private Sub UnixEscapeField(ByRef FieldString As String, _
                        ByRef EscapeCharacter As String, _
                        ByRef FDeliter As String, _
                        ByRef RDelimiter As String)

    Dim FieldDelimiterInField As Boolean
    Dim MultiLineField As Boolean
    Dim QuotesInField As Boolean

    FieldDelimiterInField = InStrB(1, FieldString, FDeliter)
    MultiLineField = InStrB(1, FieldString, RDelimiter)
    QuotesInField = InStrB(1, FieldString, EscapeCharacter)
    If FieldDelimiterInField Or MultiLineField Or QuotesInField Then
        If QuotesInField Then
            FieldString = Join$(Split(FieldString, EscapeCharacter), UnixEscapeSeq)
        End If
        If FieldDelimiterInField Then
            FieldString = Join$(Split(FieldString, FDeliter), UnixEscapeSeqDel)
        End If
        If MultiLineField Then
            FieldString = Join$(Split(FieldString, RDelimiter), UnixEscapeSeqRec)
        End If
    End If
End Sub
Private Function UnixToStandardEscapeSeq(ByRef UnixEscapedString As String, _
                                    ByRef UnixEscapeSeq As String, _
                                    ByRef NewEscapeSeq As String) As String

    If InStrB(1, UnixEscapedString, UnixEscapeSeq) Then
        UnixToStandardEscapeSeq = Join$(Split(UnixEscapedString, UnixEscapeSeq), NewEscapeSeq)
    Else
        UnixToStandardEscapeSeq = UnixEscapedString
    End If
End Function

