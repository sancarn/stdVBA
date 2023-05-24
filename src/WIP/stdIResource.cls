VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdIResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = false
Attribute VB_Exposed = False

Enum EStandardResourceType
  StandardFile
  StandardFolder
End Enum

Enum EResourceHashType
  HashTypeMD5
  HashTypeSHA1
  HashTypeSHA256
End Enum


'Obtain name of resource
'@return {String} The name of the resource
Public Property Get Name() as string: End Property

'Obtain a path to the resource
'@returns {String} The path to this resource
Public Property Get Path() as string: End Property

'Obtain local paths which build up this resource.
'@returns {String()} A local path to this resource.
'@remark {1} This may not be the original path and may be a copy of the resource.
'@remark {2} A resource may be built off numerous files, in which case an array of local paths are returned.
Public Property Get LocalPaths() as string(): End Property

'Returns True if the resource is a container for other resources, false otherwise
'@returns {Boolean} True if resource is a container, false otherwise.
Public Property Get isContainer() as boolean: End Property

'Returns true if resource supports content. For instance in SharePoint, folders can contain additional information e.g. "Postcode" etc. Implementors of this property  
'@returns {Boolean} True if resource supports content, false otherwise
Public Property Get isContentSupported() as boolean: End Property

'Returns the contents of the resource
'@returns {Byte()} The contents of the resource.
'@throws {#438} Should only be when `isContentSupported` returns `false`.
'@remarks {1} A container may have contents.
'@devRemarks {1} Return `""` if no content 
Public Property Get Contents(Optional ByVal iStartAt as Long, Optional ByVal iByteCount as Long) as Byte(): End Property
Public Property Let Contents(Optional ByVal iStartAt as Long, Optional ByVal iByteCount as Long) as Byte(): End Property

'The total size in Bytes of the resource
'@returns {Long} The total size in bytes of the resource
Public Property Get Size() as Long: End Property

'Obtain hash of resource
'@return {String} The hash of the resource
Public Property Get Hash(Optional iType as EResourceHashType) as string: End Property

'Obtains a collection of children
'@returns {Collection<stdIResource>}
Public Property Get Children() as Collection: End Function

'Copies the current resource to another existing resource container.
'@param {stdIResource} The resource to copy this resource to.
'@throws {...}
Public Sub CopyTo(ByVal res as stdIResource): End Sub
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
Public Sub MoveTo(ByVal res as stdIResource): End Sub

'Deletes the file
'@throws {...}
Public Sub Delete(): End Sub

'@param {string} Name of the resource to create. I.E. FileName / FolderName
'@param {EStandardResourceType}  
'@param {Variant<Number | String | Byte()>} The data the newly created resource should contain.
'@remark It's encouraged that developers provide their own `CreateChild` routine with improved type information if exists. E.G. `CreateChild(name, iType as ESharepointFileType, ...)
Public Function CreateChild(ByVal name as string, ByVal iStandardType as EStandardResourceType, Optional ByVal contents as Variant, Optional ByVal bOverrite as boolean) as stdIResource: End Function
'It's advised you keep 

'Identify this resource within some ContainerResource. Return an array of strings equal to all paths consumed by this class
'@param {stdIResource} A container to identify this resource within.
'@returns {Collection<stdIResource>} Collection of resources which are supported by this class.
'@remark This function is implemented for ResourceManagers to allocate files correctly.
Public Function Identify(ByVal ContainerResource as stdIResource) as Collection: End Property