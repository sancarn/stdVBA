VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdIShellExtension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Checks if the extension is of the given item
'@param item - The item to check if the extension is of
'@returns - True if the extension is of the given item, false otherwise
'@example ```vb
'stdIShellExtension.IsExtensionOf(stdShell.Create("C:\my\files\song.mp3"))
'```
'@exampleImplementation ```vb
'Public Function IsExtensionOf(ByVal item as stdShell) as Boolean
'  IsExtensionOf = item.isFile and item.Extension = "zip"
'End Function
'```
Public Function IsExtensionOf(ByVal item as stdShell) as Boolean: End Function

'Creates an extension of the given item
'@param item - The item to create an extension of
'@returns - The extension
'@example ```vb
'stdIShellExtension.CreateExtension(stdShell.Create("C:\my\files\song.mp3")).Object.Name
'```
'@exampleImplementation ```vb
'Public Function CreateExtension(ByVal item as stdShell) as stdIShellExtension
'  Set CreateExtension = New stdZipExtension
'  Call CreateExtension.protInit(item)
'End Function
'```
Public Function CreateExtension(ByVal item as stdShell) as stdIShellExtension:End Function

'Returns the raw object / class, this is critical for situations like in stdLambda where type inference
'cannot be used to convert to a extension object, so it remains as the interface stdIShellExtension unless
'this method is called.
'@returns - Raw object / class
'@example ```vb
'stdIShellExtension.CreateExtension(stdShell.Create("C:\my\files\song.mp3")).Object.Name
'```
'@exampleImplementation ```vb
'Public Property Get Object() as Object
'  Set Object = Me
'End Property
'```
Public Property Get Object() as Object: End Property

'Returns a collection of descendents (ignoring shellExtensions)
'@returns - Collection of descendents
'@example ```vb
'stdIShellExtension.DescendentsRaw().Count
'```
'@exampleImplementation ```vb
'Public Function DescendentsRaw() As Collection
'  Set DescendentsRaw = New Collection
'  For each descendent in This.Files
'    DescendentsRaw.Add descendent
'  Next
'End Function
'```
Public Function DescendentsRaw() As Collection: End Function