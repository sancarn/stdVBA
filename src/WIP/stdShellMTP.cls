VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdShellMTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type TThis
    shellFolder as Object
End Type
Private This As TThis

'Create a stdShellMTP object from a device name
'@param deviceName - Device name to create the stdShellMTP object from
'@returns - stdShellMTP object
'@example ```vb
'stdShellMTP.CreateFromDeviceName("My Phone")
'```
Public Function CreateFromDeviceName(ByVal deviceName as string) as stdShellMTP
  Const CSIDL_DRIVES = 17
  Dim sh As Object: Set sh = CreateObject("Shell.Application")
  Dim thisPC As Object: Set thisPC = sh.Namespace(CSIDL_DRIVES)
  Dim it As Object
  For Each it In thisPC.items
    If it.Name = deviceName Then
      Set CreateFromDeviceName = New stdShellMTP
      Call CreateFromDeviceName.protInit(it)
      Exit Function
    End If
  Next it
  Err.Raise 1, "stdShellMTP#CreateFromDeviceName", "Device not found"
End Function

'Create a stdShellMTP object from a ShellFolder
'@param ShellFolder - ShellFolder to create the stdShellMTP object from
'@returns - stdShellMTP object
'@example ```vb
'stdShellMTP.CreateFromShellFolder(ShellFolder)
'```
Public Function CreateFromShellFolder(ByVal ShellFolder as Object) as stdShellMTP
    Set CreateFromShellFolder = New stdShellMTP
    Call CreateFromShellFolder.protInit(ShellFolder)
End Function

'Create a collection of all portable devices
'@returns - Collection of all portable devices
'@example ```vb
'stdShellMTP.CreateAllPortableDevices()
'```
Public Function CreateAllPortableDevices() as Collection
    Dim ret As Collection: Set ret = New Collection
    Dim sh As Object: Set sh = CreateObject("Shell.Application")
    Dim thisPC As Object: Set thisPC = sh.Namespace(CSIDL_DRIVES)
    Dim it As Object
    For Each it In thisPC.items
        ret.Add CreateFromShellFolder(it)
    Next it
    Set CreateAllPortableDevices = ret
End Function

'Initialise the stdShellMTP object
'@protected
'@param ShellFolder - ShellFolder to create the stdShellMTP object from
Friend Sub protInit(ByVal ShellFolder as Object)
    Set This.shellFolder = ShellFolder
End Sub

'Get the children of the stdShellMTP object
'@returns - Collection of children
'@example ```vb
'stdShellMTP.CreateFromDeviceName("My Phone").Children()
'```
Public Function Children() as Collection
    Dim ret As Collection: Set ret = New Collection
    Dim it As Object
    For Each it In This.GetFolder().items
        ret.Add CreateFromShellFolder(it)
    Next it
    Set Children = ret
End Function

'Get the parent of the stdShellMTP object
'@returns - Parent of the stdShellMTP object
'@example ```vb
'stdShellMTP.CreateFromDeviceName("My Phone").Parent
'```
Public Property Get Parent() as stdShellMTP
    If This.shellFolder.parentFolder is nothing then
        Set Parent = Nothing
        Exit Property
    End If
    Set Parent = CreateFromShellFolder(This.shellFolder.parentFolder)
End Property

'Get the path of the stdShellMTP object
'@returns - Path of the stdShellMTP object
'@example ```vb
'stdShellMTP.CreateFromDeviceName("My Phone").Path
'```
Public Property Get Path() as string
    If This.Parent is nothing then
        Path = "mtp://"
    else
        Dim parentPath as string: parentPath = This.Parent.Path
        Path = parentPath & "/" & This.Name
    End If
End Property

'Get the name of the stdShellMTP object
'@returns - Name of the stdShellMTP object
'@example ```vb
'stdShellMTP.CreateFromDeviceName("My Phone").Name
'```
Public Property Get Name(Optional ByVal includingExtension as boolean = false) as string
    Name = This.shellFolder.name
    if not includingExtension then Name = Left(Name, InStrRev(Name, ".") - 1)
End Property

'Get the existence of the stdShellMTP object
'@returns - Existence of the stdShellMTP object
'@example ```vb
'stdShellMTP.CreateFromDeviceName("My Phone").Exists
'```
Public Property Get Exists() as boolean
    Exists = This.shellFolder.exists
End Property