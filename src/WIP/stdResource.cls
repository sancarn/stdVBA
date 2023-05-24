VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Implements stdIResource
Private Type ThisResource
    name as string
    path as string
    kind as EStandardResourceType
    exists as boolean
End Type
Private This as ThisResource

'**********************
'*     Constructors
'**********************

'Create from existing file
'@constructor
'@param {string} Path to file
'@returns {stdResource} File object
Public Function CreateFromFile(ByVal sPath as string) as stdResource
    Set CreateFromFile = CreateFromExisting(StandardFile, sPath)
End Function

'Create from existing dir
'@constructor
'@param {string} Path to dir
'@returns {stdResource} Dir object
Public Function CreateFromDir(ByVal sPath as string) as stdResource
    Set CreateFromDir = CreateFromExisting(StandardFolder, sPath)
End Function

'Create a new dir
'@constructor
'@param {string} Path to dir
'@returns {stdResource} Dir object
Public Function CreateDir(ByVal sPath as string) as stdResource
    'TODO: Create subfolders if they don't exist
    'Create folder
    MkDir sPath

    'Call constructor
    set CreateDir = CreateFromExisting(StandardFolder, sPath)
End Function

'Create a new file and add contents
'@constructor
'@param {string} Path to dir
'@returns {stdResource} Dir object
Public Function CreateFile(ByVal sPath as string, Optional ByVal Contents as string = "") as stdResource
    'TODO: Create subfolders if they don't exist
    'Create file
    Dim ff as Long: ff = FreeFile()
    Open sPath for output as #ff
        if Contents <> "" then Print #ff, Contents
    Close #ff

    'Call constructor
    set CreateFile = CreateFromExisting(StandardFile, sPath)
End Function


'Create a new file and add contents
'@constructor
'@param {EStandardResourceType} Type of resource
'@param {string} Path to resource
'@returns {stdResource} Dir object
Public Function CreateFromExisting(ByVal iType as EStandardResourceType, ByVal sPath as string) as stdResource
    set CreateFromExisting = new stdResource
    Call CreateFromExisting.protInit(iType, sPath)
End Function


'Initialise object
'@constructor
'@param {EStandardResourceType} Type of resource
'@param {string} Path to resource
Public Sub protInit(ByVal iType as EStandardResourceType, ByVal sPath as string)
    With This
        .name = FileNameFromPath(sPath)
        .path = sPath
        .kind = iType
        .exists = true
    end with
End Sub

'**********************
'*     Properties
'**********************

'Check whether file/folder is hidden
Public Property Get Hidden() as Boolean
    Hidden = protAttr(vbHidden)
End Property
Public Property Let Hidden(b as Boolean)
    protAttr(vbHidden) = b
End Property

'Check whether file/folder is read only
Public Property Get ReadOnly()
    ReadOnly = protAttr(vbReadOnly)
End Property
Public Property Let ReadOnly(b as Boolean)
    protAttr(vbReadOnly) = b
End Property





'******************************
'*     INTERFACE REDIRECTS
'******************************

Property Get Name() as string
  Name = stdIResource_Name() 
End Property

Property Get Path() as string
  Path = stdIResource_Path() 
End Property

Property Get LocalPaths() as string
  LocalPaths = stdIResource_LocalPaths() 
End Property

Property Get isContainer() as boolean
  isContainer = stdIResource_isContainer() 
End Property

Property Get isContentSupported() as boolean
  isContentSupported = stdIResource_isContentSupported() 
End Property

Property Get Contents(Optional ByVal iStartAt as Long, Optional ByVal iByteCount as Long) as Byte
  Contents = stdIResource_Contents(Optional ByVal iStartAt as Long, Optional ByVal iByteCount as Long) 
End Property

Property Get Size() as Long
  Size = stdIResource_Size() 
End Property

Property Get Hash() as string
  Hash = stdIResource_Hash()
End Property

'@returns {Collection<stdResource>}
'TODO: Is there an easier way to map to a new interface? Does it require a new collection?
Property Get Children() as Collection
  Dim cc as Collection
  set cc = new Collection
  
  'Cast to stdResource
  Dim c as stdResource
  For each c in stdIResource_Children()  
    cc.add c
  next
  
  set Children = cc 
End Property

Sub CopyTo(ByVal res as stdIResource) 
  Call stdIResource_CopyTo(res) 
End Sub

Sub MoveTo(ByVal res as stdIResource) 
  Call stdIResource_MoveTo(res) 
End Sub


Sub Delete() 
  Call stdIResource_Delete() 
End Sub

'@returns {stdResource}
Function CreateChild(ByVal name as string, ByVal iStandardType as EStandardResourceType, Optional ByVal contents as Variant, Optional ByVal bOverrite as boolean) as stdResource
  set CreateChild = stdIResource_CreateChild(name, iStandardType, contents, bOverrite) 
End Function

'@returns {Collection<stdResource>}
Function Identify(ByVal ContainerResource as stdIResource) as Collection
  Dim cc as Collection
  set cc = new Collection
  
  'Cast to stdResource
  Dim c as stdResource
  For each c in stdIResource_Identify(ContainerResource) 
    cc.add c
  next

  set Identify = cc
End Function








'Get whether an attribute is set / set an attribute to a value
'@remark Intentionally doesn't store the attribute in the class, as it may have been changed by another program.
Public Property Get protAttr(ByVal iAttr as vbFileAttribute) as Boolean
    protAttr = ((GetAttr(This.path) And iAttr) = iAttr)
End Property
Public Property Let protAttr(ByVal iAttr as vbFileAttribute, bLet as Boolean)
    if bLet then
        SetAttr This.path, This.attr or iAttr
    else
        SetAttr This.path, This.attr and not iAttr
    end if
End Property

'**************************************
'*     stdIResource Implementation
'**************************************

'Obtain a path to the resource
'@returns {String} The path to this resource
Public Property Get stdIResource_Name() as string
    stdIResource_Name = This.Name
End Property

'Obtain a path to the resource
'@returns {String} The path to this resource
Public Property Get stdIResource_Path() as string
    stdIResource_Path = This.Path
End Property

'Obtain hash of resource
'@return {String} The hash of the resource
Public Property Get stdIResource_Hash(Optional iType as EResourceHashType = HashTypeMD5) as string
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

    select case This.kind
        case StandardFile
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
        case StandardFolder
            'Deep hash, hash children into self then hash hashes
            Dim sHashString: sHashString = ""
            Dim c as stdIResource
            For each c in stdIResource_Children
                s = s & c.Hash(iType)
            next
            buffer = hashEncoder.HashData(s)
    end select

    'Output hash as hex string
    stdIResource_Hash = ""
    For i = 0 To (iHashLength - 1)
        stdIResource_Hash = stdIResource_Hash & IIf(buffer(i) < 16, "0", "") & Hex(buffer(i))
    Next
End Property

'Obtain size of file.
'@returns {long} size in bytes of the file.
Public Property Get stdIResource_Size()
    stdIResource_Size = FileLen(This.Path)
End Property

'Obtain local paths which build up this resource.
'@returns {String()} A local path to this resource.
'@remark {1} This may not be the original path and may be a copy of the resource.
'@remark {2} A resource may be built off numerous files, in which case an array of local paths are returned.
Public Property Get stdIResource_LocalPaths() as string()
    Redim stdIResource_LocalPaths(0 to 0)
    stdIResource_LocalPaths(0) = This.Path
End Property

'Returns True if the resource is a container for other resources, false otherwise
'@returns {Boolean} True if resource is a container, false otherwise.
Public Property Get stdIResource_isContainer() as boolean
    stdIResource_isContainer = (This.kind = StandardFolder)
End Property

'Returns true if resource supports content. For instance in SharePoint, folders can contain additional information e.g. "Postcode" etc. Implementors of this property  
'@returns {Boolean} True if resource supports content, false otherwise
Public Property Get stdIResource_isContentSupported() as boolean
    stdIResource_isContentSupported = (This.kind = StandardFile)
End Property

'Returns the contents of the resource
'@returns {Byte()} The contents of the resource.
'@throws {#438} Should only be when `isContentSupported` returns `false`.
'@remarks {1} A container may have contents.
'@devRemarks {1} Return `""` if no content 
Public Property Get stdIResource_Contents(Optional ByVal iStartAt as Long=0, Optional ByVal iByteCount as Long = -1) as Byte()
    if stdIResource_isContentSupported then
        Dim ff as Long: ff = FreeFile()
        Open path for Binary Access Read as #ff 
            if iStartAt > 0 then Seek #ff, iStartAt
            Dim iNumBytes as long: iNumBytes = iif(iByteCount = -1, LOF(ff)-iStartAt, iByteCount)
            if iByteCount = -1 then
            stdIResource_Contents = Input(iNumBytes,ff)
        Close #ff
    else
        Err.Raise 440, "stdResource", "Content not supported by this resource"
    end if
End Property
Public Property Let stdIResource_Contents(Optional ByVal iStartAt as Long=0, Optional ByVal iByteCount as Long = -1) as Byte()

End Property

'Obtains a collection of children
'@returns {Collection<stdIResource>}
Public Property Get stdIResource_Children() as Collection
    set stdIResource_Children = new Collection
    Dim res as stdIResource: 
End Function

'Copies the current resource to another existing resource container.
'@param {stdIResource} The resource to copy this resource to.
'@throws {...}
Public Sub stdIResource_CopyTo(ByVal res as stdIResource)
    select case This.kind 
        case StandardFile
            Call res.CreateChild(stdIResource_Name, StandardFile, stdIResource_Contents)
        case StandardFolder
            Call res.CreateChild(stdIResource_Name, StandardFolder)
    end select
End Sub
'_ Example implementation:
'_   Function CopyTo(res as stdIResource)
'_     resType = iif(me.isContainer, StandardFolder, StandardFile)
'_     Call res.CreateChildEx(me.name, xxx, me.contents, true)
'_   End Function
'_ Example2 MultiFile implementation:
'_   Function CopyTo(res as stdIResource)
'_     For res in me.resources
'_       resType = iif(res.isContainer, StandardFolder, StandardFile)
'_       Call res.CreateChildEx(file.name, xxx, me.contents, true)
'_     next
'_   End Function

'Moves the current resource to another existing resource container.
'@param {stdIResource} The resource to move this resource to.
'@throws {...}
Public Sub stdIResource_MoveTo(ByVal res as stdIResource)
    Call stdIResource_CopyTo(res)
    Call stdIResource_Delete()
End Sub

'Deletes the file
'@throws {...}
Public Sub stdIResource_Delete()
    select case This.kind
        case StandardFile
            Kill This.Path
        case StandardFolder
            Dim child as stdResource
            For each child in stdIResource_ChildrenCollection
                child.Delete
            next
            RmDir This.path
    end select
End Sub

'@param {string} Name of the resource to create. I.E. FileName / FolderName
'@param {EStandardResourceType}  
'@param {Variant<Number | String | Byte()>} The data the newly created resource should contain.
'@remark It's encouraged that developers provide their own `CreateChild` routine with improved type information if exists. E.G. `CreateChild(name, iType as ESharepointFileType, ...)
Public Function stdIResource_CreateChild(ByVal name as string, ByVal iStandardType as EStandardResourceType, Optional ByVal contents as Variant, Optional ByVal bOverrite as boolean) as stdIResource
    if stdIResource_isContainer then
        select case iStandardType
            case StandardFile
                ff = FreeFile()
                Open File stdIResource_Path & "\" & name as #ff for Output
                    if not isEmpty(contents) then Print #ff, contents
                Close #ff
            case StandardFolder
                MkDir stdIResource_Path & "\" & name
        end select
    else
        Err.Raise ERR_NOT_IMPL, "", "Not implemented"
    end if
End Function

'Identify this resource within some ContainerResource. Return an array of strings equal to all paths consumed by this class
'@param {stdIResource} A container to identify this resource within.
'@returns {Collection<stdIResource>}
'@remark This function is implemented for ResourceManagers to allocate files correctly.
Public Function stdIResource_Identify(ByVal ContainerResource as stdIResource) as Collection
    'All children can be a `stdResource`
    set stdIResource_Identify = ContainerResource.ChildrenCollection
End Property