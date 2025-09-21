VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Issues:
' CORE BEHAVIOUR
' - [ ] Missing Properties: `Attributes`, `DateModified`, `DateAccessed`, `DateCreated`, `IsReadOnly`, `IsHidden`,
' - [ ] On windows, migrate to Win32 API rather than native calls.

'I/O ENCODING
' - [ ] Append, Overwrite, Clear etc. shouldn't use `Print` as this appends a newline. Use `Write` instead. Perhaps add a AppendLine method.
' - [ ] ReadText uses Input$ (locale/ANSI), no UTF-8/UTF-16 handling.
' - [ ] No binary read/write methods; no large-file chunking helpers.
' - [ ] CreateFile always appends newline when sContents provided.
'GENERAL
' - [ ] Support for mobile devices (MTP?) Unsure how possible this is... `EShellFileTypeMobileDevice`
' - [ ] Support alternate data streams.`EShellFileTypeAlternateDataStream`
' - [ ] File hashing should use Win32 API rather than .NET COM classes. Performance much better with Win32 API.
' - [ ] Path normalisation `\` and `/` should be interchangeable.
' - [ ] Better error handling generally, with descriptive errors.

Public Enum stdShellKind
    EShellFileTypeUnknown
    EShellFileTypeRoot
    EShellFileTypeFile
    EShellFileTypeFolder
End Enum
Public Enum stdShellHashType
    EShellHashTypeMD5
    EShellHashTypeSHA1
    EShellHashTypeSHA256
End Enum

Private Type TThis
    path as string
    kind as stdShellKind
    isVirtual as boolean
    shellExtensions as Collection
End Type
Private This As TThis

'Expose type via property for external reads
Public Property Get FileType() As stdShellKind
    FileType = This.kind
End Property

'New explicit kind property
Public Property Get Kind() As stdShellKind
    Kind = This.kind
End Property

'Returns True if the shell item is deemed to be a file
'@returns - True if the shell item is a file
'@example ```vb
'stdShell.Create("C:\my\files\song.mp3").IsFile
'```
Public Property Get IsFile() As Boolean
    IsFile = This.kind = EShellFileTypeFile
End Property

'Returns True if the shell item is deemed to be a folder
'@returns - True if the shell item is a folder
'@example ```vb
'stdShell.Create("C:\my\files").IsFolder
'```
Public Property Get IsFolder() As Boolean
    IsFolder = This.kind = EShellFileTypeFolder
End Property

'Is the shell item virtual? Virtual shell items are not backed by a physical file or folder.
'@returns - True if the shell item is virtual
'@example ```vb
'stdShell.Create("C:\my\files").IsVirtual
'```
Public Property Get IsVirtual() As Boolean
    IsVirtual = This.isVirtual
End Property

'Create a new stdShell object. Kind is determined from the path, else use `AsFile` or `AsFolder` to define it.
'@constructor
'@param path - The path to wrap a shell object around.
Public Function Create(ByVal path As String) As stdShell
    Dim kind As stdShellKind
    Dim isVirtual As Boolean
    kind = privGetSystemType(path)
    isVirtual = (kind = EShellFileTypeUnknown)
    
    Set Create = New stdShell
    Call Create.protInit(path, kind, isVirtual)
End Function

'Create a virtual file handle (no I/O)
'@constructor
'@param path - Path to file
'@returns {stdShell} Virtual shell object of file
Public Function CreateFile(ByVal path As String) As stdShell
    Set CreateFile = New stdShell
    Call CreateFile.protInit(path, EShellFileTypeFile, privGetSystemType(path) = EShellFileTypeUnknown)
End Function

'Create a virtual directory handle (no I/O)
'@constructor
'@param Path to folder
'@returns {stdShell} Virtual shell object of directory
Public Function CreateFolder(ByVal path As String) As stdShell
    Set CreateFolder = New stdShell
    Call CreateFolder.protInit(path, EShellFileTypeFolder, privGetSystemType(path) = EShellFileTypeUnknown)
End Function

'Create a root shell object
'The root shell object is a virtual shell object that represents the root of the file system. It 
'contains all drives, devices and network shares.
'@constructor
'@returns - Root shell object
Public Function CreateRoot() as stdShell
    Set CreateRoot = New stdShell
    #if Mac then
        Call CreateRoot.protInit("/", EShellFileTypeFolder, True)
    #else
        Call CreateRoot.protInit("", EShellFileTypeRoot, True)
    #end if
End Function


'Create a temporary shell object. Use `AsFile` or `AsFolder` to define its type and `Ensure` to create it.
'@constructor
'@returns - Temporary shell object
Public Function CreateTemporaryItem() as stdShell
    Dim tempFolder as string: tempFolder = Environ("TEMP")
    Dim guid as string: guid = getGUID()
    Set CreateTemporaryFile = New stdShell
    Call CreateTemporaryFile.protInit(tempFolder & "\" & guid, EShellFileTypeUnknown, True)
End Function

'Initialise object
'@protected
'@param path - Path to file or folder
'@param pType - Type of shell item
'@param isVirtual - Whether the shell item is virtual
Friend Sub protInit(ByVal path As String, pType As stdShellKind, ByVal isVirtual As Boolean)
    This.path = path
    This.kind = pType
    This.isVirtual = isVirtual
    set This.shellExtensions = protGetExtensions()
End Sub

'Add a shell extension
'@static
'@param extension - Shell extension to add
Public Sub BindExtension(ByVal extension as stdIShellExtension)
    if This.shellExtensions is nothing Then Set This.shellExtensions = new Collection
    Call This.shellExtensions.Add(extension)
End Sub

'Get the shell extensions
'@protected
'@returns - Shell extensions
Public Function protGetExtensions() as Collection
    Set protGetExtensions = This.shellExtensions
End Function

'Specify that the shell item is a file. Useful when ensuring a file exists.
'@returns - File shell object
'@example ```vb
'stdShell.Create("C:\my\files\song.mp3").AsFile().Ensure()
'```
Public Function AsFile() as stdShell
    Set AsFile = CreateFile(This.path)
End Function

'Specify that the shell item is a folder. Useful when ensuring a folder exists.
'@returns - Folder shell object
'@example ```vb
'stdShell.Create("C:\my\files.stuff").AsFolder().Ensure()
'```
Public Function AsFolder() as stdShell
    Set AsFolder = CreateFolder(This.path)
End Function

'Sanitize the current path into a single legal item under the root
'Replaces illegal characters and any slashes with underscores, preserves extension for files
'@returns - New stdShell pointing to sanitized item path
Public Function ToLegalItem(optional ByVal replacement as string = "_") As stdShell
    Dim original As String: original = This.path
    original = Replace(original, "!", replacement)
    original = Replace(original, "*", replacement)
    original = Replace(original, "|", replacement)
    original = Replace(original, "<", replacement)
    original = Replace(original, ">", replacement)
    original = Replace(original, ":", replacement)
    original = Replace(original, "?", replacement)
    
    Set ToLegalItem = New stdShell
    Call ToLegalItem.protInit(original, This.kind, privGetSystemType(newPath) = EShellFileTypeUnknown)
End Function

'Convert the shell item to a FileSystemObject
'@returns - FileSystemObject
'@example ```vb
'stdShell.Create("C:\my\files\song.mp3").ToFSO()
'```
Public Function ToFSO() as Object
    select case This.kind
        case EShellFileTypeFile
            Set ToFSO = CreateObject("Scripting.FileSystemObject").GetFile(This.path)
        case EShellFileTypeFolder
            Set ToFSO = CreateObject("Scripting.FileSystemObject").GetFolder(This.path)
        case else
            Err.Raise 1, "", "Cannot convert this item to a FileSystemObject"
    end select
End Function

'Convert the shell item to an IStream
'@returns - IStream
'@example ```vb
'stdShell.Create("C:\my\files\song.mp3").ToIStream()
'```
Public Function ToIStream() as Object
    Set ToIStream = CreateObject("ADODB.Stream")
    With ToIStream
        .Type = 1 'adTypeBinary
        .Mode = 3 'adModeReadWrite
        .Open
        .LoadFromFile This.path
        .Position = 0
    End With
End Function

'Create a shortcut to the shell item
'@param path - Path to save the shortcut to
'@returns - Success flag
'@example ```vb
'stdShell.Create("C:\my\files\song.mp3").SaveAsShortcut("C:\my\files\song.lnk")
'```
Public Function SaveAsShortcut(ByVal path as string) as Boolean
    On Error GoTo EH
    Dim shortcut As Object: Set shortcut = CreateObject("WScript.Shell").CreateShortcut(path)
    shortcut.TargetPath = This.path
    shortcut.Save

    SaveAsShortcut = True
    Exit Function
EH:
    SaveAsShortcut = False
End Function

'Returns the full path of the shell item
'@returns - Full path of the shell item
'@example ```vb
'stdShell.Create("C:\my\files\song.mp3").FullPath
'```
Public Property Get FullPath() as string
    FullPath = This.path
End Property

'Returns the name of the shell item
'@returns - Name of the shell item
'@example ```vb
'stdShell.Create("C:\my\files\song.mp3").Name
'```
Public Property Get Name(Obtional ByVal includingExtension as boolean = false) as string
    If This.kind = EShellFileTypeRoot Then Exit Property
    Name = Mid(This.path, InStrRev(This.path, "\") + 1)
    if not includingExtension then Name = Left(Name, InStrRev(Name, ".") - 1)
End Property

'Returns the extension of the shell item
'@returns - Extension of the shell item
'@example ```vb
'stdShell.Create("C:\my\files\song.mp3").Extension
'```
Public Property Get Extension() as string
    If This.kind <> EShellFileTypeFile Then Exit Property
    Extension = Mid(This.path, InStrRev(This.path, ".") + 1)
End Property

' Returns True if path exists. If true, clears virtual and updates kind if unknown
'@returns - True if path exists
'@example ```vb
'If stdShell.Create("C:\my\files\song.mp3").Exists() Then
'    Debug.Print "File exists"
'End If
'```
Public Property Get Exists() As Boolean
    If This.kind = EShellFileTypeRoot Then 
        Exists = True
        Exit Property
    End if

    Dim s As String: s = This.path
    Dim k As stdShellKind: k = privGetSystemType(s)
    If k <> EShellFileTypeUnknown Then
        Exists = True
        If This.kind = EShellFileTypeUnknown Then This.kind = k
        This.isVirtual = False
    Else
        Exists = False
    End If
End Property

' Ensure file/folder exists based on intended kind; optionally create parents
'@param createParents - Whether to create parents if they don't exist
'@returns - True if the file/folder exists, false otherwise
'@example ```vb
'If stdShell.Create("C:\my\files\song.mp3").Ensure() then
'    Debug.Print "File now guaranteed to exist"
'End If
'```
Public Function Ensure(Optional ByVal createParents As Boolean = True) As Boolean
    'Root is always guaranteed to exist
    If This.kind = EShellFileTypeRoot Then 
        Ensure = True
        Exit Function
    End if
    
    'If it already exists, then we're done
    Dim s As String: s = This.path
    If Exists() Then Ensure = True: Exit Function
    Select Case This.kind
        Case EShellFileTypeFile
            If createParents Then
                Dim p As stdShell
                Set p = Me.Parent
                If Not p Is Nothing Then
                    If Not p.Ensure(True) Then Exit Function
                End If
            End If
            On Error GoTo EH
            Dim ff As Long: ff = FreeFile
            Open s For Output As #ff: Close #ff
            This.isVirtual = False
            Ensure = True
            Exit Function
        Case EShellFileTypeFolder
            On Error GoTo EH
            If createParents Then
                Dim pp As stdShell
                Set pp = Me.Parent
                If Not pp Is Nothing Then
                    If Not pp.Ensure(True) Then Exit Function
                End If
            End If
            MkDir s
            This.isVirtual = False
            Ensure = True
            Exit Function
        Case Else
            Ensure = False
            Exit Function
    End Select
EH:
    Ensure = False
End Function

' Ensure empty file; update timestamp if exists
'@returns - True if the file exists / was created, false otherwise
Public Function Touch() As Boolean
    If This.kind <> EShellFileTypeFile Then Exit Function
    Dim s As String: s = This.path
    On Error GoTo EH
    If Not Exists() Then
        Dim ff As Long: ff = FreeFile
        Open s For Output As #ff: Close #ff
        This.isVirtual = False
        Touch = True
        Exit Function
    End If
    'TODO: Refactor to use native methods
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(s) Then
        fso.GetFile(s).DateLastModified = Now
        This.isVirtual = False
        Touch = True
        Exit Function
    End If
EH:
    Touch = False
End Function

' Parent shell item
'@returns - Parent shell item
'@example ```vb
'stdShell.Create("C:\my\files\song.mp3").Parent
'```
Public Property Get Parent() As stdShell
    If This.kind = EShellFileTypeRoot Then Exit Property
    Dim p As String: p = privGetParentPath(This.path)
    If p <> "" Then Set Parent = stdShell.CreateFolder(p)
End Property


' Returns the size of the file
'@returns - Size of the file
'@example ```vb
'stdShell.Create("C:\my\files\song.mp3").Size
'```
Public Property Get Size() As Long
  Select Case This.kind
    Case EShellFileTypeFile
      Size = FileLen(This.path)
    Case Else
      Err.Raise 1, "", "Not implemented, use `GetSize` function instead."
  End Select
End Property

'Get the size of the file/folder
'@returns - Size of the file/folder
'@remark This is a recursive function that will return the size of the file/folder and all its children.
'@example ```vb
'stdShell.Create("C:\my\files\song.mp3").GetSize()
'```
Public Function GetSize() As Long
    Select Case This.kind
        Case EShellFileTypeFile
            GetSize = FileLen(This.path)
        Case else
            if not This.isVirtual Then
                Dim child as stdShell
                For each child in Children
                    GetSize = GetSize + child.GetSize()
                Next
                DoEvents
            end if
    End Select
End Function

'Obtains a collection of children (ignoring shellExtensions)
'@returns {Collection<stdShell>}
Public Property Get ChildrenRaw() As Collection
    Set ChildrenRaw = New Collection
    select case This.kind
        case EShellFileTypeFile
            'No children, Do nothing
        case EShellFileTypeFolder
            'Get child paths
            Dim tpath As String
            tpath = VBA.Dir(This.path & IIf(Right(This.path, 1) = "\", "*", "\*"), vbNormal Or vbDirectory Or vbHidden Or vbSystem)
            While tpath <> ""
                If tpath <> "." And tpath <> ".." Then Call ChildrenRaw.Add(stdShell.Create(This.path & "\" & tpath))
                tpath = VBA.Dir()
            Wend
            DoEvents
        case EShellFileTypeRoot
            'Use FSO for now... TODO: Use win32 API instead.
            Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
            Dim drive As Object
            For each drive in fso.Drives
                if drive.DriveLetter <> "" then
                    ChildrenRaw.Add stdShell.Create(drive.DriveLetter & ":\")
                end if
            Next
    end select
End Property

'Obtains a collection of descendents (ignoring shellExtensions)
'@returns {Collection<stdShell>}
Public Property Get DescendentsRaw() as Collection
    Dim ret As Collection: set ret = new Collection
    Dim child As stdShell
    For each child in ChildrenRaw
        ret.add child
        For each descendent in child.DescendentsRaw
            ret.add descendent
        Next
    Next
    Set DescendentsRaw = ret
End Property


'Obtains a collection of children (including shellExtensions)
'@returns Collection<stdShell | stdIShellExtension> - Collection of children
Public Property Get Children() as Collection
    Dim raw As Collection: Set raw = Me.ChildrenRaw
    Dim consumed as Object: set consumed = CreateObject("Scripting.Dictionary")
    Dim ret as Collection: set ret = new Collection
    
    'Extend children
    Dim ext as stdIShellExtension
    For each ext in This.shellExtensions
        Dim child As stdShell
        For each child in raw
            If not consumed.exists(child.FullPath) Then 
                If ext.IsExtensionOf(child) Then
                    Dim extendedChild as stdIShellExtension
                    set extendedChild = ext.CreateExtension(child)
                    ret.add extendedChild
                    consumed(child.FullPath) = true
                    Dim desc as Variant 'stdIShellExtension or stdShell
                    For each desc in extendedChild.DescendentsRaw
                        consumed(desc.FullPath) = true
                    next
                end if
            end if
        Next
    Next

    'Add remaining children that were not extended
    For each child in raw
        If not consumed(child.FullPath) Then
            ret.add child
        end if
    next
    
    Set Children = ret
End Property

'Attempts to read data from a file
'@returns {Byte()} - Bytes read from file
Public Function Read() As Byte()
    If This.kind = EShellFileTypeFile Then
        Dim ff As Long: ff = FreeFile
        Open This.path For Input As #ff
            Read = Input(LOF(ff), #ff)
        Close #ff
    Else
        Err.Raise 1, "", "Cannot read this item as it is not a file"
    End If
End Function

'Attempts to read data from a file
'@returns {Byte()} - Bytes read from file
Public Function ReadText() As String
    If This.kind = EShellFileTypeFile Then
        Dim ff As Long: ff = FreeFile
        Open This.path For Input As #ff
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
    On Error GoTo EH
        If This.kind = EShellFileTypeFile Then
            Dim ff As Long: ff = FreeFile
            Open This.path For Append As #ff
                Seek #ff, LOF(ff) + 1
                Put #ff, , CStr(data)
            Close #ff
        Else
            Err.Raise 1, "", "Cannot append data to this item as it is not a file"
        End If
        Exit Function
EH:
    Append = False
End Function

'Appends a line to the file
'@param line {ByVal Variant} - Line to append to file
'@returns {Boolean} - Success flag
Public Function AppendLine(ByVal line As string) As Boolean
    If This.kind = EShellFileTypeFile Then
        AppendLine = Append(data & vbCrLf)
    Else
        Err.Raise 1, "", "Cannot append line to this item as it is not a file"
    End If
End Function

'Attempts to overwrite the contents of the file
'@param data {ByVal Variant} - Data to write to file
'@returns {Boolean} - Success flag
Public Function Overwrite(ByVal data As Variant) As Boolean
    On Error GoTo EH
        If This.kind = EShellFileTypeFile Then
            Dim ff As Long: ff = FreeFile
            Open This.path For Output As #ff
                Put #ff, , data
            Close #ff
        Else
            Err.Raise 1, "", "Cannot write data to this item as it is not a file"
        End If
        Exit Function
EH:
    Overwrite = False
End Function

'Attempts to clear the file
'@returns {Boolean} - Success flag
Public Function Clear() As Boolean
    On Error GoTo EH
        If This.kind = EShellFileTypeFile Then
            Dim ff As Long: ff = FreeFile
            Open This.path For Output As #ff
            Close #ff
        Else
            Err.Raise 1, "", "Cannot clear data from this item as it is not a file"
        End If
        Exit Function
EH:
    Clear = False
End Function

'Obtain a hash value for file contents
'@param {stdShellHashType} The type of hash to create
'@returns Hex string representing hash.
'@remark {1} MD5 and SHA1 kept for legacy support.
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
    Open This.path For Binary Access Read As #ff
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




'Finds the first shell item that matches the callable
'@param callable - Callable to match shell item against
'@param bRecurse - Whether to recurse into children
'@returns - First shell item that matches the callable
'@example ```vb
'stdShell.Create("C:\my\files\song.mp3").Find(Callable)
'```
Public Function Find(ByVal callable As stdICallable, Optional ByVal bRecurse As Boolean = True) As stdShell
    If This.kind = EShellFileTypeFolder Then
        'Get children as collection
        Dim col As Collection: Set col = Children
        
        Dim shellItem As stdShell
        While Not CollectionShift(col, shellItem) Is Nothing
            Dim child As stdShell
            For Each child In shellItem.Children
                If callable.Run(child) Then
                    Set Find = child
                    Exit Function
                End If
                
                If child.Kind = EShellFileTypeFolder Then col.add child
            Next
        Wend
    End If
End Function

'Finds all shell items that match the callable
'@param callable - Callable to match shell item against
'@returns - Collection of shell items that match the callable
'@example ```vb
'stdShell.Create("C:\my\files\song.mp3").FindAll(Callable)
'```
Public Function FindAll(ByVal callable As stdICallable) As Collection
    Dim vRet As Collection: Set vRet = New Collection
    If This.kind = EShellFileTypeFolder Then
        'Get children as collection
        Dim col As Collection: Set col = Children
        
        Dim shellItem As stdShell
        While Not CollectionShift(col, shellItem) Is Nothing
            Dim child As stdShell
            For Each child In shellItem.Children
                If callable.Run(child) Then
                    vRet.Add child
                End If
                
                If child.Kind = EShellFileTypeFolder Then col.add child
            Next
        Wend
    End If
    Set FindAll = vRet
End Function

'Deletes the file/folder
'@returns {Boolean} - Success flag
'@example ```vb
'Call stdShell.Create("C:\my\files\song.mp3").Delete()
'```
Public Function Delete() As Boolean
    On Error GoTo EH
        If Exists Then
            select case This.kind
                case EShellFileTypeRoot Then Exit Function
                case EShellFileTypeFolder Then
                    Dim child as stdShell
                    For each child in Children
                        child.Delete
                    next
                    RmDir This.path
                case EShellFileTypeFile Then
                    Kill This.path
                case else
                    Err.Raise 1, "", "Cannot delete this item as it is not a file"
            end select
            This.isVirtual = True
        End If
        Delete = True
        Exit Function
EH:
    Delete = False
End Function

'Copies the file/folder to a new path
'@param path - Path to copy the file/folder to
'@returns - Success flag
'@example ```vb
'Call stdShell.Create("C:\my\files\song.mp3").CopyTo("C:\my\files\song2.mp3")
'```
Public Function CopyTo(ByVal path As String) As Boolean
    On Error GoTo EH
        If Exists Then
            select case This.kind
                case EShellFileTypeFile
                    Call CopyFile(This.path, path)
                case EShellFileTypeFolder
                    Call CopyFolder(This.path, path)
                case else
                    Err.Raise 1, "", "Cannot copy this item as it is not a file or folder"
            end select
            CopyTo = True
        End If
        Exit Function
EH:
    CopyTo = False
End Function

'Copies the file/folder to a new folder
'@param folderPath - Path to copy the file/folder to
'@returns - Success flag
'@example ```vb
'Call stdShell.Create("C:\my\files\song.mp3").CopyToFolder("C:\my\files")
'```
Public Function CopyToFolder(ByVal folderPath As String) As Boolean
    if right(folderPath, 1) <> "\" then folderPath = folderPath & "\"
    Call CopyTo(folderPath & Name(true))
End Function


'Moves the file/folder to a new path
'@param path - Path to move the file/folder to
'@returns - Success flag
'@example ```vb
'Call stdShell.Create("C:\my\files\song.mp3").MoveTo("C:\my\files\song2.mp3")
'```
Public Function MoveTo(ByVal path As String) As Boolean
    On Error GoTo EH
        If Exists Then
            Call CopyTo(path)
            Call Delete()
            MoveTo = True
        End If
        Exit Function
EH:
    MoveTo = False
End Function


'Moves the file/folder to a new folder
'@param folderPath - Path to move the file/folder to
'@returns - Success flag
'@example ```vb
'Call stdShell.Create("C:\my\files\song.mp3").MoveToFolder("C:\my\files")
'```
Public Function MoveToFolder(ByVal folderPath As String) As Boolean
    if right(folderPath, 1) <> "\" then folderPath = folderPath & "\"
    Call MoveTo(folderPath & Name(true))
End Function


'Renames the file/folder
'@param newName - New name of the file/folder
'@returns - Success flag
'@example ```vb
'Call stdShell.Create("C:\my\files\song.mp3").Rename("song2.mp3")
'```
Public Function Rename(ByVal newName As String) As Boolean
    On Error GoTo EH
        If Exists Then
            Call MoveTo(Parent.Path & "\" & newName & Extension)
            Rename = True
        End If
        Exit Function
EH:
    Rename = False
End Function

'====================
' Private helpers
'====================
'Shift a value out of a collection
'@param col - Collection to shift value out of
'@param shellItem - Shell item shifted out of collection
'@returns - Shell item shifted out of collection
Private Function CollectionShift(ByRef col As Collection, ByRef shellItem As stdShell) As stdShell
    Set CollectionShift = col.Item(1)
    Set shellItem = col.Item(1)
    Call col.Remove(1)
End Function

'Get the parent path of a given path
'@param s - Path to get the parent path of
'@returns - Parent path of the given path
Private Function privGetParentPath(ByVal s As String) As String
    Dim path As String: path = s
    If VBA.Len(path) = 0 Then Exit Function
    If Right$(path, 1) = "\" Then
        If Not privIsDriveOrShare(path) Then path = Left$(path, Len(path) - 1)
    End If
    Dim i As Long: i = InStrRev(path, "\")
    If i <= 0 Then Exit Function
    Dim parent As String: parent = Left$(path, i - 1)
    If parent = "" Then Exit Function
    If Mid$(path, 2, 1) = ":" And i = 3 Then parent = Left$(path, 3)
    If privIsDriveOrShare(parent) Then
        ' Root has no parent; return empty to stop recursion
        privGetParentPath = ""
        Exit Function
    End If
    privGetParentPath = parent
End Function

'Check if a path is a drive or share
'@param s - Path to check if it is a drive or share
'@returns - True if the path is a drive or share, false otherwise
Private Function privIsDriveOrShare(ByVal s As String) As Boolean
    If Len(s) = 3 And Mid$(s, 2, 1) = ":" And Right$(s, 1) = "\" Then
        privIsDriveOrShare = True
    ElseIf Left$(s, 2) = "\\" Then
        Dim parts() As String: parts = Split(Mid$(s, 3), "\")
        privIsDriveOrShare = (UBound(parts) < 1)
    Else
        privIsDriveOrShare = False
    End If
End Function

'Get the system type of a given path
'@param path - Path to get the system type of
'@returns - System type of the given path
Private Function privGetSystemType(ByVal path As String) As stdShellKind
    On Error GoTo EH
    If VBA.LenB(VBA.Dir(path, vbDirectory Or vbNormal Or vbHidden Or vbSystem)) = 0 Then
        privGetSystemType = EShellFileTypeUnknown
        Exit Function
    End If
    Dim attr As VbFileAttribute: attr = VBA.GetAttr(path)
    If (attr And VbFileAttribute.vbDirectory) = VbFileAttribute.vbDirectory Or _
       (attr And VbFileAttribute.vbVolume) = VbFileAttribute.vbVolume Then
        privGetSystemType = EShellFileTypeFolder
    Else
        privGetSystemType = EShellFileTypeFile
    End If
    Exit Function
EH:
    privGetSystemType = EShellFileTypeUnknown
End Function

'Generates a new GUID
'@returns - A new GUID
Private Function getGUID() As String
  Call Randomize 'Ensure random GUID generated
  getGUID = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx"
  getGUID = Replace(getGUID, "y", Hex(Rnd() And &H3 Or &H8))
  Dim i As Long: For i = 1 To 30
    getGUID = Replace(getGUID, "x", Hex$(Int(Rnd() * 16)), 1, 1)
  Next
End Function