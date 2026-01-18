VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdCSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@module
'@description A class for parsing, manipulating, and exporting CSV data. Supports lazy loading
'and streaming for efficient memory usage with large files. Can work with files, strings, or arrays.
'@example Simple example of creating a CSV from file and converting to array:
'```vb
'Dim csv as stdCSV: set csv = stdCSV.CreateFromFile("C:\data.csv")
'Dim arr as Variant: arr = csv.toArray()
'Debug.Print csv.Count()  '=> number of rows
'```
'@example Creating CSV from string:
'```vb
'Dim csv as stdCSV: set csv = stdCSV.CreateFromString("a,b,c" & vbCrLf & "1,2,3")
'csv.toRange Range("A1")
'```
'@example Creating CSV from array:
'```vb
'Dim arr(0 to 1, 0 to 2) as Variant
'arr(0,0) = "a": arr(0,1) = "b": arr(0,2) = "c"
'arr(1,0) = 1: arr(1,1) = 2: arr(1,2) = 3
'Dim csv as stdCSV: set csv = stdCSV.CreateFromArray(arr)
'csv.toCSV "C:\output.csv"
'```

Option Explicit

Public Enum stdECSVEncoding
    stdECSVEncodingUTF8
    stdECSVEncodingANSI
End Enum

Public Enum stdECSVType
    stdCSVTypeFile
    stdCSVTypeString
    stdCSVTypeArray
End Enum

Private Type CSVDialect
    fieldDelimiter As String
    recordDelimiter As String
    quoteChar As String
    escapeChar As String
End Type

Private Type FileBuffer
    filePath As String
    fileNum As Long
    filePosition As Long
    fileSize As Long
    buffer As String
    bufferPosition As Long
    bufferSize As Long
    atEndOfFile As Boolean
    isOpen As Boolean
End Type

Private Type StringBuffer
    ByRefStr As Variant 'VT_BYREF to be applied for performance gains
    position As Long
    length As Long
End Type

Private Type ParserConfig
    detectDelimiter As Boolean
    detectEncoding As Boolean
    commentsToken As String
    columnTypes As Variant
    headers As Variant
    hasHeaders As Boolean
    skipEmptyLines As Boolean
    skipCommentLines As Boolean
    bufferSize As Long
    dialect As CSVDialect
End Type

Private Type ParserState
    csvType As stdECSVType
    fileBuffer As FileBuffer
    stringBuffer As StringBuffer
    config As ParserConfig
    currentLine As Long
    isInitialized As Boolean
    headerRow() As Variant
    headerDetected As Boolean
End Type

Private Type TThis
    parserState As ParserState
End Type

Private This As TThis

'Constants
Private Const CHR_DOUBLE_QUOTES As String = """"
Private Const CHR_APOSTROPHE As String = "'"
Private Const CHR_BACKSLASH As String = "\"
Private Const ASCW_CHR_BACKSLASH As Long = 92

'Create a CSV object from file
'@constructor
'@param sPath - Path to CSV file
'@param detectDelimiter - Whether to auto-detect delimiter (default: True)
'@param detectEncoding - Whether to auto-detect encoding (default: True)
'@param commentsToken - Token that indicates comment lines (default: empty)
'@param columnTypes - Array of column types for type conversion
'@param headers - Headers array or boolean indicating if first row is headers (default: true)
'@returns - stdCSV object
Public Function CreateFromFile(ByVal sPath As String, Optional ByVal detectDelimiter As Boolean = True, Optional ByVal detectEncoding As Boolean = True, Optional ByVal commentsToken As String = vbNullString, Optional ByVal columnTypes As Variant = Empty, Optional ByVal headers As Variant = True) As stdCSV
    Set CreateFromFile = New stdCSV
    Call CreateFromFile.protInit(stdCSVTypeFile, sPath, detectDelimiter, detectEncoding, commentsToken, columnTypes, headers)
End Function

'Create a CSV object from string
'@constructor
'@param sString - CSV string data (passed by reference for performance)
'@param detectDelimiter - Whether to auto-detect delimiter (default: True)
'@param detectEncoding - Whether to auto-detect encoding (default: True)
'@param commentsToken - Token that indicates comment lines (default: empty)
'@param columnTypes - Array of column types for type conversion
'@param headers - Headers array or boolean indicating if first row is headers
'@returns - stdCSV object
Public Function CreateFromString(ByRef sString As String, Optional ByVal detectDelimiter As Boolean = True, Optional ByVal detectEncoding As Boolean = True, Optional ByVal commentsToken As String = vbNullString, Optional ByVal columnTypes As Variant = Empty, Optional ByVal headers As Variant = True) As stdCSV
    Set CreateFromString = New stdCSV
    Call CreateFromString.protInit(stdCSVTypeString, sString, detectDelimiter, detectEncoding, commentsToken, columnTypes, headers)
End Function

'Create a CSV object from array
'@constructor
'@param vArray - 2D array or jagged array containing CSV data
'@param headers - Headers array or boolean indicating if first row is headers
'@returns - stdCSV object
Public Function CreateFromArray(ByRef vArray As Variant, Optional ByVal headers As Variant = True) As stdCSV
    Set CreateFromArray = New stdCSV
    Call CreateFromArray.protInit(stdCSVTypeArray, vArray, headers)
End Function

'Initialize the CSV object
'@protected
'@param iType - Type of CSV source (file, string, or array)
'@param params - Parameters array containing source data and configuration
Friend Sub protInit(ByVal iType As stdECSVType, ParamArray params())
    This.parserState.csvType = iType
    This.parserState.isInitialized = True
    This.parserState.currentLine = 0
    This.parserState.headerDetected = False
    
    'Initialize parser config with defaults
    With This.parserState.config
        .detectDelimiter = True
        .detectEncoding = True
        .commentsToken = vbNullString
        .hasHeaders = True
        .skipEmptyLines = False
        .skipCommentLines = False
        .bufferSize = 1024 * 1024 '1MB default buffer
        .dialect.fieldDelimiter = ","
        .dialect.recordDelimiter = vbCrLf
        .dialect.quoteChar = CHR_DOUBLE_QUOTES
        .dialect.escapeChar = CHR_DOUBLE_QUOTES
    End With
    
    'Set up buffer based on type
    Select Case iType
        Case stdCSVTypeFile
            If UBound(params) >= 0 Then
                This.parserState.fileBuffer.filePath = CStr(params(0))
            End If
            If UBound(params) >= 1 Then This.parserState.config.detectDelimiter = CBool(params(1))
            If UBound(params) >= 2 Then This.parserState.config.detectEncoding = CBool(params(2))
            If UBound(params) >= 3 Then This.parserState.config.commentsToken = CStr(params(3))
            If UBound(params) >= 4 Then
                If Not IsEmpty(params(4)) Then This.parserState.config.columnTypes = params(4)
            End If
            If UBound(params) >= 5 Then
                If Not IsEmpty(params(5)) Then
                    If VarType(params(5)) = vbBoolean Then
                        This.parserState.config.hasHeaders = CBool(params(5))
                    Else
                        This.parserState.config.headers = params(5)
                        This.parserState.config.hasHeaders = True
                    End If
                End If
            End If
            Call initializeFile()
            
        Case stdCSVTypeString
            If UBound(params) >= 0 Then
                'Store string by reference for performance
                This.parserState.stringBuffer.ByRefStr = params(0)
                This.parserState.stringBuffer.length = Len(CStr(params(0)))
                This.parserState.stringBuffer.position = 1
            End If
            If UBound(params) >= 1 Then This.parserState.config.detectDelimiter = CBool(params(1))
            If UBound(params) >= 2 Then This.parserState.config.detectEncoding = CBool(params(2))
            If UBound(params) >= 3 Then This.parserState.config.commentsToken = CStr(params(3))
            If UBound(params) >= 4 Then
                If Not IsEmpty(params(4)) Then This.parserState.config.columnTypes = params(4)
            End If
            If UBound(params) >= 5 Then
                If Not IsEmpty(params(5)) Then
                    If VarType(params(5)) = vbBoolean Then
                        This.parserState.config.hasHeaders = CBool(params(5))
                    Else
                        This.parserState.config.headers = params(5)
                        This.parserState.config.hasHeaders = True
                    End If
                End If
            End If
            Call detectDialectFromString()
            
        Case stdCSVTypeArray
            If UBound(params) >= 0 Then
                'Store array reference
                This.parserState.stringBuffer.ByRefStr = params(0)
            End If
            If UBound(params) >= 1 Then
                If Not IsEmpty(params(1)) Then
                    If VarType(params(1)) = vbBoolean Then
                        This.parserState.config.hasHeaders = CBool(params(1))
                    Else
                        This.parserState.config.headers = params(1)
                        This.parserState.config.hasHeaders = True
                    End If
                End If
            End If
            'Arrays don't need dialect detection
    End Select
End Sub

'Initialize file for reading
'@private
Private Sub initializeFile()
    With This.parserState.fileBuffer
        If Dir(.filePath) = "" Then
            Err.Raise 53, "stdCSV::initializeFile", "File not found: " & .filePath
        End If
        
        .fileNum = FreeFile
        .fileSize = FileLen(.filePath)
        .filePosition = 0
        .bufferPosition = 0
        .bufferSize = This.parserState.config.bufferSize
        .atEndOfFile = False
        .isOpen = False
        
        'Open file and keep it open
        Call openFile()
        
        'Load initial buffer
        Call loadFileBuffer()
        
        'Detect dialect if needed
        If This.parserState.config.detectDelimiter Then
            Call detectDialectFromFile()
        End If
    End With
End Sub

'Open the CSV file
'@private
Private Sub openFile()
    With This.parserState.fileBuffer
        If .isOpen Then Exit Sub
        
        Open .filePath For Binary As #.fileNum
        .isOpen = True
    End With
End Sub

'Close the CSV file
'@public
'@returns - Self for chaining
Public Function CloseFile() As stdCSV
    With This.parserState.fileBuffer
        If .isOpen Then
            Close #.fileNum
            .isOpen = False
        End If
    End With
    Set CloseFile = Me
End Function

'Reopen the CSV file (if closed)
'@public
'@returns - Self for chaining
Public Function Open() As stdCSV
    If This.parserState.csvType = stdCSVTypeFile Then
        Call openFile()
    End If
    Set Open = Me
End Function

'Load buffer from file
'@private
Private Sub loadFileBuffer()
    With This.parserState.fileBuffer
        If .atEndOfFile Then Exit Sub
        
        'Ensure file is open
        If Not .isOpen Then
            Call openFile()
        End If
        
        Dim bytesToRead As Long: bytesToRead = IIf(.fileSize - .filePosition > .bufferSize, .bufferSize, .fileSize - .filePosition)
        
        If bytesToRead <= 0 Then
            .atEndOfFile = True
            .buffer = ""
            Exit Sub
        End If
        
        'Read from already-open file
        Seek #.fileNum, .filePosition + 1
        .buffer = Space$(bytesToRead)
        Get #.fileNum, , .buffer
        
        .filePosition = .filePosition + bytesToRead
        .bufferPosition = 1
        .atEndOfFile = (.filePosition >= .fileSize)
    End With
End Sub

'Detect CSV dialect from file
'@private
Private Sub detectDialectFromFile()
    Dim savedPosition As Long: savedPosition = This.parserState.fileBuffer.filePosition
    Dim savedBufferPosition As Long: savedBufferPosition = This.parserState.fileBuffer.bufferPosition
    
    'Read sample from beginning of file
    This.parserState.fileBuffer.filePosition = 0
    This.parserState.fileBuffer.bufferPosition = 1
    This.parserState.fileBuffer.atEndOfFile = False
    
    'Load initial buffer for detection
    Call loadFileBuffer()
    
    'Get sample
    Dim sample As String: sample = Left$(This.parserState.fileBuffer.buffer, IIf(Len(This.parserState.fileBuffer.buffer) > 10000, 10000, Len(This.parserState.fileBuffer.buffer)))
    Call detectDialect(sample)
    
    'Reset file position to beginning
    This.parserState.fileBuffer.filePosition = 0
    This.parserState.fileBuffer.bufferPosition = 1
    This.parserState.fileBuffer.atEndOfFile = False
    Call loadFileBuffer()
End Sub

'Detect CSV dialect from string
'@private
Private Sub detectDialectFromString()
    Dim sData As String: sData = CStr(This.parserState.stringBuffer.ByRefStr)
    Call detectDialect(sData)
End Sub

'Detect CSV dialect (simplified version)
'@private
Private Sub detectDialect(ByRef sample As String)
    Dim delimiters() As String
    Dim lineEndings() As String
    Dim i As Long, j As Long
    Dim bestDelimiter As String: bestDelimiter = ""
    Dim bestLineEnding As String: bestLineEnding = vbCrLf
    Dim maxScore As Long: maxScore = 0
    Dim score As Long: score = 0
    
    'Common delimiters to test
    ReDim delimiters(0 To 3)
    delimiters(0) = ","
    delimiters(1) = ";"
    delimiters(2) = vbTab
    delimiters(3) = "|"
    
    'Line endings to test
    ReDim lineEndings(0 To 2)
    lineEndings(0) = vbCrLf
    lineEndings(1) = vbLf
    lineEndings(2) = vbCr
    
    'Simple heuristic: count occurrences of each delimiter in first few lines
    Dim firstLines As String: firstLines = Left$(sample, IIf(Len(sample) > 1000, 1000, Len(sample)))
    
    For i = 0 To UBound(delimiters)
        score = 0
        Dim pos As Long: pos = 1
        Do While pos > 0 And pos < Len(firstLines)
            pos = InStr(pos, firstLines, delimiters(i))
            If pos > 0 Then
                score = score + 1
                pos = pos + 1
            End If
        Loop
        If score > maxScore Then
            maxScore = score
            bestDelimiter = delimiters(i)
        End If
    Next i
    
    'Detect line ending
    If InStr(1, firstLines, vbCrLf) > 0 Then
        bestLineEnding = vbCrLf
    ElseIf InStr(1, firstLines, vbLf) > 0 Then
        bestLineEnding = vbLf
    ElseIf InStr(1, firstLines, vbCr) > 0 Then
        bestLineEnding = vbCr
    Else
        bestLineEnding = vbCrLf 'Default
    End If
    
    'Detect quote character
    Dim quoteChar As String
    If InStr(1, firstLines, CHR_DOUBLE_QUOTES) > 0 Then
        quoteChar = CHR_DOUBLE_QUOTES
    ElseIf InStr(1, firstLines, CHR_APOSTROPHE) > 0 Then
        quoteChar = CHR_APOSTROPHE
    Else
        quoteChar = CHR_DOUBLE_QUOTES 'Default
    End If
    
    'Set dialect
    With This.parserState.config.dialect
        .fieldDelimiter = bestDelimiter
        .recordDelimiter = bestLineEnding
        .quoteChar = quoteChar
        .escapeChar = quoteChar
    End With
End Sub

'Parse a single line from CSV
'@param lineData - Output array containing parsed fields
'@param lineNumber - Current line number (output)
'@param fieldCount - Number of fields parsed (output)
'@returns - True if line was parsed, False if end of data
Public Function parseLine(ByRef lineData As Variant, ByRef lineNumber As Long, ByRef fieldCount As Long) As Boolean
    'Special handling for arrays - direct access
    If This.parserState.csvType = stdCSVTypeArray Then
        parseLine = parseLineFromArray(lineData, lineNumber, fieldCount)
        Exit Function
    End If
    
    Dim lineText As String: lineText = ""
    Dim lineStart As Long: lineStart = 0
    Dim lineEnd As Long: lineEnd = 0
    
    'Get next line
    If Not getNextLineText(lineText, lineStart, lineEnd) Then
        parseLine = False
        Exit Function
    End If
    
    'Skip empty lines if configured
    If This.parserState.config.skipEmptyLines And Trim$(lineText) = "" Then
        'Recurse to get next line
        parseLine = Me.parseLine(lineData, lineNumber, fieldCount)
        Exit Function
    End If
    
    'Skip comment lines if configured
    If This.parserState.config.skipCommentLines And This.parserState.config.commentsToken <> vbNullString Then
        If Left$(Trim$(lineText), Len(This.parserState.config.commentsToken)) = This.parserState.config.commentsToken Then
            'Recurse to get next line
            parseLine = Me.parseLine(lineData, lineNumber, fieldCount)
            Exit Function
        End If
    End If
    
    'Parse fields from line
    Dim fields() As Variant
    ReDim fields(0 To 0)
    Dim field As String: field = ""
    Dim inQuotes As Boolean: inQuotes = False
    Dim pos As Long: pos = 1
    Dim char As String: char = ""
    Dim i As Long: i = 0
    
    Do While pos <= Len(lineText)
        char = Mid$(lineText, pos, 1)
        
        If char = This.parserState.config.dialect.quoteChar Then
            If inQuotes Then
                'Check if escaped quote or end of field
                If pos < Len(lineText) And Mid$(lineText, pos + 1, 1) = This.parserState.config.dialect.quoteChar Then
                    'Escaped quote
                    field = field & This.parserState.config.dialect.quoteChar
                    pos = pos + 2
                ElseIf pos < Len(lineText) And Mid$(lineText, pos + 1, 1) = This.parserState.config.dialect.fieldDelimiter Then
                    'End of quoted field
                    inQuotes = False
                    pos = pos + 2
                Else
                    'End of quotes
                    inQuotes = False
                    pos = pos + 1
                End If
            Else
                'Start of quoted field
                inQuotes = True
                pos = pos + 1
            End If
        ElseIf char = This.parserState.config.dialect.fieldDelimiter And Not inQuotes Then
            'Field delimiter found
            ReDim Preserve fields(0 To UBound(fields) + 1)
            fields(UBound(fields) - 1) = field
            field = ""
            pos = pos + 1
        Else
            field = field & char
            pos = pos + 1
        End If
    Loop
    
    'Add last field
    ReDim Preserve fields(0 To UBound(fields) + 1)
    fields(UBound(fields)) = field
    
    'Handle headers
    If This.parserState.config.hasHeaders And Not This.parserState.headerDetected Then
        This.parserState.headerDetected = True
        ReDim This.parserState.headerRow(0 To UBound(fields))
        For i = 0 To UBound(fields)
            This.parserState.headerRow(i) = fields(i)
        Next i
        'Recurse to get first data line
        parseLine = Me.parseLine(lineData, lineNumber, fieldCount)
        Exit Function
    End If
    
    'Return parsed line
    This.parserState.currentLine = This.parserState.currentLine + 1
    lineNumber = This.parserState.currentLine
    fieldCount = UBound(fields) + 1
    lineData = fields
    parseLine = True
End Function

'Parse line directly from array (more efficient)
'@private
Private Function parseLineFromArray(ByRef lineData As Variant, ByRef lineNumber As Long, ByRef fieldCount As Long) As Boolean
    Dim vArray As Variant
    vArray = This.parserState.stringBuffer.ByRefStr
    
    If Not IsArray(vArray) Then
        parseLineFromArray = False
        Exit Function
    End If
    
    Dim currentLine As Long: currentLine = This.parserState.currentLine
    
    'Handle headers
    If This.parserState.config.hasHeaders And Not This.parserState.headerDetected Then
        This.parserState.headerDetected = True
        'Get first row as headers
        Dim firstRow As Variant
        
        On Error Resume Next
        Dim testVal As Variant: testVal = vArray(LBound(vArray))
        Dim isJagged As Boolean: isJagged = IsArray(testVal)
        On Error GoTo 0
        
        Dim iLB1 As Long: iLB1 = 0
        Dim iUB1 As Long: iUB1 = 0
        Dim iLB2 As Long: iLB2 = 0
        Dim iUB2 As Long: iUB2 = 0
        
        If isJagged Then
            firstRow = vArray(LBound(vArray))
        Else
            iLB1 = LBound(vArray, 1)
            iLB2 = LBound(vArray, 2)
            iUB2 = UBound(vArray, 2)
            ReDim firstRow(iLB2 To iUB2)
            Dim j As Long
            For j = iLB2 To iUB2
                firstRow(j) = vArray(iLB1, j)
            Next j
        End If
        
        If IsArray(firstRow) Then
            iLB2 = LBound(firstRow)
            iUB2 = UBound(firstRow)
            ReDim This.parserState.headerRow(0 To iUB2 - iLB2)
            For j = iLB2 To iUB2
                This.parserState.headerRow(j - iLB2) = firstRow(j)
            Next j
        End If
        
        'Recurse to get first data line
        parseLineFromArray = Me.parseLineFromArray(lineData, lineNumber, fieldCount)
        Exit Function
    End If
    
    'Get data row
    On Error Resume Next
    testVal = vArray(LBound(vArray))
    isJagged = IsArray(testVal)
    On Error GoTo 0
    
    Dim row As Variant
    Dim i As Long: i = 0
    Dim iLB1 As Long: iLB1 = 0
    Dim iUB1 As Long: iUB1 = 0
    Dim iLB2 As Long: iLB2 = 0
    Dim iUB2 As Long: iUB2 = 0
    
    If isJagged Then
        iLB1 = LBound(vArray)
        iUB1 = UBound(vArray)
        If currentLine < (iUB1 - iLB1 + 1) Then
            row = vArray(iLB1 + currentLine)
            If IsArray(row) Then
                iLB2 = LBound(row)
                iUB2 = UBound(row)
                ReDim lineData(0 To iUB2 - iLB2)
                For i = iLB2 To iUB2
                    lineData(i - iLB2) = row(i)
                Next i
                This.parserState.currentLine = This.parserState.currentLine + 1
                lineNumber = This.parserState.currentLine
                fieldCount = iUB2 - iLB2 + 1
                parseLineFromArray = True
            End If
        End If
    Else
        iLB1 = LBound(vArray, 1)
        iUB1 = UBound(vArray, 1)
        iLB2 = LBound(vArray, 2)
        iUB2 = UBound(vArray, 2)
        If currentLine < (iUB1 - iLB1 + 1) Then
            ReDim lineData(0 To iUB2 - iLB2)
            For i = iLB2 To iUB2
                lineData(i - iLB2) = vArray(iLB1 + currentLine, i)
            Next i
            This.parserState.currentLine = This.parserState.currentLine + 1
            lineNumber = This.parserState.currentLine
            fieldCount = iUB2 - iLB2 + 1
            parseLineFromArray = True
        End If
    End If
End Function

'Parse a chunk of lines from CSV
'@param chunkData - Output 2D array containing parsed rows
'@param startRow - Starting row number (output)
'@param rowCount - Number of rows in chunk (output)
'@param colCount - Number of columns (output)
'@param chunkSize - Maximum number of rows to parse (default: 1000)
'@returns - True if chunk was parsed, False if end of data
Public Function parseChunk(ByRef chunkData As Variant, ByRef startRow As Long, ByRef rowCount As Long, ByRef colCount As Long, Optional ByVal chunkSize As Long = 1000) As Boolean
    Dim rows() As Variant
    ReDim rows(1 To chunkSize, 1 To 1)
    Dim lineData As Variant
    Dim lineNumber As Long: lineNumber = 0
    Dim fieldCount As Long: fieldCount = 0
    Dim maxCols As Long: maxCols = 0
    startRow = This.parserState.currentLine + 1
    
    Dim i As Long: i = 1
    For i = 1 To chunkSize
        If Not Me.parseLine(lineData, lineNumber, fieldCount) Then
            'End of data
            If i = 1 Then
                'No data at all
                parseChunk = False
                Exit Function
            End If
            'Resize array to actual size
            ReDim Preserve rows(1 To i - 1, 1 To maxCols)
            rowCount = i - 1
            colCount = maxCols
            chunkData = rows
            parseChunk = True
            Exit Function
        End If
        
        'Update max columns
        If fieldCount > maxCols Then
            maxCols = fieldCount
            'Resize array if needed
            If maxCols > UBound(rows, 2) Then
                ReDim Preserve rows(1 To chunkSize, 1 To maxCols)
            End If
        End If
        
        'Copy line data to chunk
        Dim j As Long: j = 1
        For j = 1 To fieldCount
            rows(i, j) = lineData(j - 1)
        Next j
    Next i
    
    'Full chunk parsed
    ReDim Preserve rows(1 To chunkSize, 1 To maxCols)
    rowCount = chunkSize
    colCount = maxCols
    chunkData = rows
    parseChunk = True
End Function

'Get next line text from source
'@private
'@param lineText - Output line text
'@param lineStart - Output line start position
'@param lineEnd - Output line end position
'@returns - True if line found, False if end of data
Private Function getNextLineText(ByRef lineText As String, ByRef lineStart As Long, ByRef lineEnd As Long) As Boolean
    Select Case This.parserState.csvType
        Case stdCSVTypeFile
            getNextLineText = getNextLineFromFile(lineText, lineStart, lineEnd)
        Case stdCSVTypeString
            getNextLineText = getNextLineFromString(lineText, lineStart, lineEnd)
        Case stdCSVTypeArray
            getNextLineText = getNextLineFromArray(lineText, lineStart, lineEnd)
    End Select
End Function

'Get next line from file
'@private
Private Function getNextLineFromFile(ByRef lineText As String, ByRef lineStart As Long, ByRef lineEnd As Long) As Boolean
    With This.parserState.fileBuffer
        Dim delimiter As String: delimiter = This.parserState.config.dialect.recordDelimiter
        
        Do
            'Find line ending in current buffer
            Dim lineEndPos As Long: lineEndPos = InStr(.bufferPosition, .buffer, delimiter)
            
            If lineEndPos > 0 Then
                'Line found in buffer
                lineText = Mid$(.buffer, .bufferPosition, lineEndPos - .bufferPosition)
                .bufferPosition = lineEndPos + Len(delimiter)
                getNextLineFromFile = True
                Exit Function
            ElseIf .atEndOfFile Then
                'Last line (no ending delimiter)
                If .bufferPosition <= Len(.buffer) Then
                    lineText = Mid$(.buffer, .bufferPosition)
                    .bufferPosition = Len(.buffer) + 1
                    getNextLineFromFile = (Len(lineText) > 0)
                Else
                    getNextLineFromFile = False
                End If
                Exit Function
            Else
                'Load more data
                Call loadFileBuffer()
            End If
        Loop
    End With
End Function

'Get next line from string
'@private
Private Function getNextLineFromString(ByRef lineText As String, ByRef lineStart As Long, ByRef lineEnd As Long) As Boolean
    With This.parserState.stringBuffer
        If .position > .length Then
            getNextLineFromString = False
            Exit Function
        End If
        
        Dim sData As String: sData = CStr(.ByRefStr)
        Dim delimiter As String: delimiter = This.parserState.config.dialect.recordDelimiter
        Dim lineEndPos As Long: lineEndPos = InStr(.position, sData, delimiter)
        
        If lineEndPos > 0 Then
            lineText = Mid$(sData, .position, lineEndPos - .position)
            .position = lineEndPos + Len(delimiter)
            getNextLineFromString = True
        Else
            'Last line
            lineText = Mid$(sData, .position)
            .position = .length + 1
            getNextLineFromString = (Len(lineText) > 0)
        End If
    End With
End Function

'Get next line from array (returns as CSV string for parsing)
'@private
Private Function getNextLineFromArray(ByRef lineText As String, ByRef lineStart As Long, ByRef lineEnd As Long) As Boolean
    Dim vArray As Variant
    vArray = This.parserState.stringBuffer.ByRefStr
    
    If Not IsArray(vArray) Then
        getNextLineFromArray = False
        Exit Function
    End If
    
    Dim currentLine As Long: currentLine = This.parserState.currentLine
    
    'Check if we have headers to skip
    If This.parserState.config.hasHeaders And Not This.parserState.headerDetected Then
        'This will be handled in parseLine
        getNextLineFromArray = False
        Exit Function
    End If
    
    'Convert array row to CSV line
    On Error Resume Next
    Dim testVal As Variant: testVal = vArray(LBound(vArray))
    Dim isJagged As Boolean: isJagged = IsArray(testVal)
    On Error GoTo 0
    
    Dim i As Long: i = 0
    Dim j As Long: j = 0
    Dim iLB1 As Long: iLB1 = 0
    Dim iUB1 As Long: iUB1 = 0
    Dim iLB2 As Long: iLB2 = 0
    Dim iUB2 As Long: iUB2 = 0
    Dim row As Variant
    Dim fields() As String
    
    If isJagged Then
        iLB1 = LBound(vArray)
        iUB1 = UBound(vArray)
        If currentLine < (iUB1 - iLB1 + 1) Then
            row = vArray(iLB1 + currentLine)
            If IsArray(row) Then
                iLB2 = LBound(row)
                iUB2 = UBound(row)
                ReDim fields(iLB2 To iUB2)
                For j = iLB2 To iUB2
                    fields(j) = CStr(row(j))
                Next j
                lineText = Join(fields, This.parserState.config.dialect.fieldDelimiter)
                getNextLineFromArray = True
            End If
        End If
    Else
        iLB1 = LBound(vArray, 1)
        iUB1 = UBound(vArray, 1)
        iLB2 = LBound(vArray, 2)
        iUB2 = UBound(vArray, 2)
        If currentLine < (iUB1 - iLB1 + 1) Then
            ReDim fields(iLB2 To iUB2)
            For j = iLB2 To iUB2
                fields(j) = CStr(vArray(iLB1 + currentLine, j))
            Next j
            lineText = Join(fields, This.parserState.config.dialect.fieldDelimiter)
            getNextLineFromArray = True
        End If
    End If
End Function

'Export CSV data to Excel Range (streaming)
'@param rng - Target range (top-left cell)
Public Sub toRange(ByRef rng As Range)
    Dim currentRow As Long: currentRow = 0
    
    'Write headers if they exist
    If This.parserState.config.hasHeaders And Not IsEmpty(This.parserState.headerRow) Then
        Dim headerRange As Range
        Set headerRange = rng.Offset(0).Resize(1, UBound(This.parserState.headerRow) + 1)
        headerRange.Value2 = This.parserState.headerRow
        currentRow = 1
    End If
    
    'Reset parser state
    This.parserState.currentLine = 0
    This.parserState.headerDetected = This.parserState.config.hasHeaders
    
    'Stream data in chunks
    Dim chunkData As Variant
    Dim iRow As Long: iRow = 0
    Dim iRowCount As Long: iRowCount = 0
    Dim iColCount As Long: iColCount = 0
    Do While Me.parseChunk(chunkData, iRow, iRowCount, iColCount, 1000)
        If IsArray(chunkData) Then
            Dim targetRange As Range
            Set targetRange = rng.Offset(currentRow).Resize(iRowCount, iColCount)
            targetRange.Value2 = chunkData
            currentRow = currentRow + iRowCount
        End If
    Loop
End Sub

'Convert CSV to 2D array (loads all data)
'@returns - 2D array containing CSV data
Public Function toArray() As Variant
    Dim allRows As Collection
    Set allRows = New Collection
    Dim lineData As Variant
    Dim lineNumber As Long: lineNumber = 0
    Dim fieldCount As Long: fieldCount = 0
    Dim maxCols As Long: maxCols = 0
    
    'Reset parser state
    This.parserState.currentLine = 0
    This.parserState.headerDetected = This.parserState.config.hasHeaders
    
    'Collect all rows
    Do While Me.parseLine(lineData, lineNumber, fieldCount)
        If fieldCount > maxCols Then maxCols = fieldCount
        allRows.Add lineData
    Loop
    
    'Convert to 2D array
    If allRows.count > 0 Then
        Dim result() As Variant
        ReDim result(1 To allRows.count + IIf(This.parserState.config.hasHeaders, 1, 0), 1 To maxCols)
        
        'Add headers if they exist
        Dim headerRow As Long: headerRow = 1
        Dim i As Long: i = 1
        Dim j As Long: j = 1
        If This.parserState.config.hasHeaders And Not IsEmpty(This.parserState.headerRow) Then
            For j = 1 To UBound(This.parserState.headerRow) + 1
                result(headerRow, j) = This.parserState.headerRow(j - 1)
            Next j
            headerRow = 2
        End If
        
        'Add data rows
        For i = 1 To allRows.count
            lineData = allRows(i)
            fieldCount = UBound(lineData) + 1
            For j = 1 To fieldCount
                If j <= maxCols Then
                    result(headerRow + i - 1, j) = lineData(j - 1)
                End If
            Next j
        Next i
        
        toArray = result
    End If
End Function

'Get the number of rows in the CSV (requires full scan)
'@returns - Number of data rows (excluding headers)
Public Function Count() As Long
    Dim lineData As Variant
    Dim lineNumber As Long: lineNumber = 0
    Dim fieldCount As Long: fieldCount = 0
    
    'Reset parser state
    This.parserState.currentLine = 0
    This.parserState.headerDetected = This.parserState.config.hasHeaders
    
    'Count lines
    Do While Me.parseLine(lineData, lineNumber, fieldCount)
        Count = Count + 1
    Loop
End Function

'Get headers if they exist
'@returns - Headers as array or Empty
Public Function Headers() As Variant
    If This.parserState.config.hasHeaders Then
        If Not IsEmpty(This.parserState.headerRow) Then
            Headers = This.parserState.headerRow
        Else
            'Parse first line to get headers
            Dim lineData As Variant
            Dim lineNumber As Long: lineNumber = 0
            Dim fieldCount As Long: fieldCount = 0
            
            This.parserState.currentLine = 0
            This.parserState.headerDetected = False
            If Me.parseLine(lineData, lineNumber, fieldCount) Then
                Headers = This.parserState.headerRow
            End If
        End If
    End If
End Function

'Export CSV to file
'@param sPath - Output file path
'@param delimiter - Field delimiter (default: detected or ",")
'@param enclosure - Quote character (default: detected or """")
'@param escape - Escape character (default: detected or """")
'@param includeHeaders - Whether to include headers (default: True)
'@param encoding - File encoding (default: UTF8)
Public Sub toCSV(ByVal sPath As String, Optional ByVal delimiter As String = "", Optional ByVal enclosure As String = "", Optional ByVal escape As String = "", Optional ByVal includeHeaders As Boolean = True, Optional ByVal encoding As stdECSVEncoding = stdECSVEncodingUTF8)
    'Use provided or detected delimiters
    Dim outputDelimiter As String: outputDelimiter = IIf(delimiter = "", This.parserState.config.dialect.fieldDelimiter, delimiter)
    Dim outputQuote As String: outputQuote = IIf(enclosure = "", This.parserState.config.dialect.quoteChar, enclosure)
    Dim outputEscape As String: outputEscape = IIf(escape = "", This.parserState.config.dialect.escapeChar, escape)
    
    Dim ff As Long: ff = FreeFile
    Open sPath For Output As #ff
    
    'Reset parser state
    This.parserState.currentLine = 0
    This.parserState.headerDetected = This.parserState.config.hasHeaders
    
    'Write headers if needed
    If includeHeaders And This.parserState.config.hasHeaders Then
        Dim headers As Variant
        headers = Me.Headers()
        If Not IsEmpty(headers) Then
            Dim headerLine As String: headerLine = ""
            Dim j As Long: j = 0
            For j = 0 To UBound(headers)
                If j > 0 Then headerLine = headerLine & outputDelimiter
                headerLine = headerLine & outputQuote & CStr(headers(j)) & outputQuote
            Next j
            Print #ff, headerLine
        End If
    End If
    
    'Write data rows
    Dim lineData As Variant
    Dim lineNumber As Long: lineNumber = 0
    Dim fieldCount As Long: fieldCount = 0
    Do While Me.parseLine(lineData, lineNumber, fieldCount)
        Dim csvLine As String: csvLine = ""
        Dim j As Long: j = 0
        For j = 0 To fieldCount - 1
            If j > 0 Then csvLine = csvLine & outputDelimiter
            Dim fieldValue As String: fieldValue = CStr(lineData(j))
            'Escape quotes in field
            fieldValue = Replace(fieldValue, outputQuote, outputEscape & outputQuote)
            csvLine = csvLine & outputQuote & fieldValue & outputQuote
        Next j
        Print #ff, csvLine
    Loop
    
    Close #ff
End Sub

'Get the CSV type
'@returns - The CSV source type
Public Function GetType() As stdECSVType
    GetType = This.parserState.csvType
End Function

'Get the detected dialect
'@returns - Dictionary with dialect information
Public Function GetDialect() As Object
    Set GetDialect = CreateDictionary("fieldDelimiter", This.parserState.config.dialect.fieldDelimiter, _
                                      "recordDelimiter", This.parserState.config.dialect.recordDelimiter, _
                                      "quoteChar", This.parserState.config.dialect.quoteChar, _
                                      "escapeChar", This.parserState.config.dialect.escapeChar)
End Function

'Reset parser to beginning
'@returns - Self for chaining
Public Function Reset() As stdCSV
    This.parserState.currentLine = 0
    This.parserState.headerDetected = False
    If This.parserState.csvType = stdCSVTypeFile Then
        'Ensure file is open
        If Not This.parserState.fileBuffer.isOpen Then
            Call openFile()
        End If
        This.parserState.fileBuffer.filePosition = 0
        This.parserState.fileBuffer.bufferPosition = 1
        This.parserState.fileBuffer.atEndOfFile = False
        Call loadFileBuffer()
    ElseIf This.parserState.csvType = stdCSVTypeString Then
        This.parserState.stringBuffer.position = 1
    End If
    Set Reset = Me
End Function

'Class termination - ensure file is closed
Private Sub Class_Terminate()
    If This.parserState.csvType = stdCSVTypeFile Then
        Call CloseFile
    End If
End Sub

'Create a dictionary
'@private
Private Function CreateDictionary(ParamArray children()) As Object
    Set CreateDictionary = CreateObject("Scripting.Dictionary")
    CreateDictionary.CompareMode = vbTextCompare
    
    Dim i As Long
    For i = LBound(children) To UBound(children) Step 2
        Call CreateDictionary.add(children(i), children(i + 1))
    Next
End Function
