VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Enum StdShellType
    File
    Dir
    Multifile
End Enum
Public Enum stdShellHashType
    HashTypeMD5
    HashTypeSHA1
    HashTypeSHA256
End Enum


Private sPaths() As String
Private sPath As String 'only populated if File or Dir
Public iType As StdShellType

'@constructor
'@param sPath {String} - The path to wrap a shell object around.
Public Function Create(ByVal sPath As String) As stdShell
    'Get paths array
    Dim tPaths(0 To 0) As String: tPaths(0) = sPath
    
    'Get type of shell item
    Dim tType As StdShellType
    Select Case VBA.GetAttr(sPath)
        Case VbFileAttribute.vbDirectory, VbFileAttribute.vbVolume
            tType = Dir
        Case Else
            tType = File
    End Select
    
    'Create shell item
    Set Create = New stdShell
    Call Create.protInit(tPaths, tType)
End Function

'Create new file as shell object
'@constructor
'@param {string} Path to file to create
'@returns {stdShell} Shell object of file
Public Function CreateFile(ByVal sPath As String, Optional ByVal sContents As String) As stdShell
    Dim ff As Long: ff = FreeFile()
    Open sPath For Append As #ff
        If sContents <> "" Then Print #ff, sContents
    Close #ff
    Set CreateFile = Create(sPath)
End Function

'Create and return directory as shell object
'@constructor
'@param {string} Path to folder to create
'@returns {stdShell} Shell object of directory
Public Function CreateDir(ByVal sPath As String) As stdShell
    MkDir sPath
    Set CreateDir = Create(sPath)
End Function

Friend Sub protInit(paramPaths() As String, pType As StdShellType)
    sPaths = paramPaths
    iType = pType
    Select Case iType
      Case File, Dir
        sPath = sPaths(0)
    End Select
End Sub

Public Property Get Path() as string
    Select Case iType
        Case File, Dir
            Path = sPath
        Case else
            Err.Raise 1, "", ""
    end select
End Property

Public Property Get Size() As Long
  Select Case iType
    Case File
      Size = FileLen(sPath)
    Case Else
      Err.Raise 1, "", "Not implemented"
  End Select
End Property

Public Property Get Children() As Collection
    Dim vRet As Collection
    If iType = File Then
        Set vRet = New Collection
    ElseIf iType = Dir Then
        'Define return collection
        Set vRet = New Collection
        
        'Get child paths
        Dim tsPath As String
        tsPath = VBA.Dir(sPath & IIf(Right(sPath, 1) = "\", "*", "\*"), vbNormal Or vbDirectory)
        While tsPath <> ""
            If tsPath <> "." And tsPath <> ".." Then Call vRet.Add(stdShell.Create(sPath & "\" & tsPath))
            tsPath = VBA.Dir()
        Wend
    End If
    
    Set Children = vRet
End Property

'Attempts to read data from a file
'@returns {Byte()} - Bytes read from file
Public Function Read() As Byte()
    If iType = File Then
        Dim ff As Long: ff = FreeFile
        Open sPath For Input As #ff
            Read = Input(LOF(ff), #ff)
        Close #ff
    Else
        Err.Raise 1, "", "Cannot read this item as it is not a file"
    End If
End Function

'Attempts to read data from a file
'@returns {Byte()} - Bytes read from file
Public Function ReadText() As String
    If iType = File Then
        Dim ff As Long: ff = FreeFile
        Open sPath For Input As #ff
            ReadText = Input$(LOF(ff), #ff)
        Close #ff
    Else
        Err.Raise 1, "", "Cannot read this item as it is not a file"
    End If
End Function

'Attempts to append data onto the end of a file
'@param data {ByVal Variant} - Data to append to file
'@returns {Boolean} - Success flag
Public Function Append(ByVal data As Variant) As Boolean
    If iType = File Then
        Dim ff As Long: ff = FreeFile
        Open sPath For Append As #ff
            Print #ff, data
        Close #ff
    Else
        Err.Raise 1, "", "Cannot append data to this item as it is not a file"
    End If
End Function

'Attempts to overwrite the contents of the file
'@param data {ByVal Variant} - Data to write to file
'@returns {Boolean} - Success flag
Public Function Overwrite(ByVal data As Variant) As Boolean
    If iType = File Then
        Dim ff As Long: ff = FreeFile
        Open sPath For Output As #ff
            Print #ff, data
        Close #ff
    Else
        Err.Raise 1, "", "Cannot write data to this item as it is not a file"
    End If
End Function

'Attempts to clear the file
'@returns {Boolean} - Success flag
Public Function Clear() As Boolean
    If iType = File Then
        Dim ff As Long: ff = FreeFile
        Open sPath For Output As #ff
        Close #ff
    Else
        Err.Raise 1, "", "Cannot clear data from this item as it is not a file"
    End If
End Function

'Obtain a hash value for file contents
'@param {stdShellHashType} The type of hash to create
'@returns {String} Hex string representing hash.
Public Function Hash(ByVal iHashType As stdShellHashType) As String
    'Obtain hash encoder
    Dim hashEncoder As Object, iHashLength As Long
    Select Case iHashType
        Case HashTypeMD5
            Set hashEncoder = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
            iHashLength = 16
        Case HashTypeSHA1
            Set hashEncoder = CreateObject("System.Security.Cryptography.SHA1CryptoServiceProvider")
            iHashLength = 20
        Case HashTypeSHA256
            Set hashEncoder = CreateObject("System.Security.Cryptography.SHA256Managed")
            iHashLength = 32
    End Select
    
    Dim blockSize As Long: blockSize = 2 ^ 16
    Dim ff As Long: ff = FreeFile
    
    'Progressively build hash 1024 bytes at a time
    Open sPath For Binary Access Read As #ff
        If LOF(ff) < blockSize Then blockSize = ((LOF(ff) + 1024) \ 1024) * 1024
        Dim buffer() As Byte: ReDim buffer(blockSize - 1)
        
        Dim i As Long
        For i = 1 To LOF(ff) \ blockSize
            Get ff, , buffer
            hashEncoder.TransformBlock buffer, 0, blockSize, buffer, 0
        Next
        Get ff, , buffer
        hashEncoder.TransformFinalBlock buffer, 0, LOF(ff) Mod blockSize
        buffer = hashEncoder.Hash
        hashEncoder.Clear
    Close #ff
    
    'Output hash as hex string
    Hash = ""
    For i = 0 To (iHashLength - 1)
        Hash = Hash & IIf(buffer(i) < 16, "0", "") & Hex(buffer(i))
    Next
End Function




'Attempts to clear the file
'
Public Function Find(ByVal callable As stdICallable, Optional ByVal bRecurse As Boolean = True) As stdShell
    If iType = Dir Then
        'Get children as collection
        Dim col As Collection: Set col = Children.AsCollection
        
        Dim shellItem As stdShell
        While Not CollectionShift(col, shellItem) Is Nothing
            Dim child As stdShell
            For Each child In shellItem.Children
                If callable.Run(child) Then
                    Set Find = child
                    Exit Function
                End If
                
                If child.iType = Dir Then col.add child
            Next
        Wend
    End If
End Function
Public Function FindAll(ByVal callable As stdICallable) As Collection
    Dim vRet As Collection: Set vRet = New Collection
    If iType = Dir Then
        'Get children as collection
        Dim col As Collection: Set col = Children.AsCollection
        
        Dim shellItem As stdShell
        While Not CollectionShift(col, shellItem) Is Nothing
            Dim child As stdShell
            For Each child In shellItem.Children
                If callable.Run(child) Then
                    vRet.add child
                End If
                
                If child.iType = Dir Then col.add child
            Next
        Wend
    End If
    Set FindAll = vRet
End Function

Private Function CollectionShift(ByRef col As Collection, ByRef shellItem As stdShell) As stdShell
    Set CollectionPop = col.Item(1)
    Set shellItem = col.Item(1)
    Call col.Remove(1)
End Function
