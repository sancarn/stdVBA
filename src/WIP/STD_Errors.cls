VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "STD_Errors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Ensure to parse stack in messages:
'  Error in ClassName::MethodName  /  ModuleName::MethodName  reading "<message>"
'Long term ideas:
'  Dim sb as Object
'  set sb = STD_Types_StringBuilder.Create()
'  sb.[(                                   ]
'  sb.[  {"id":1,"message":"Cool message"},]
'  sb.[  {"id":2,"message":"Sad message"}  ]
'  sb.[)                                   ]
'  STD.Errors.AddNamespace("CustomErrors",STD.Data.JSON.Parse(sb))
'  STD.Errors.Raise("CustomErrors",1) '==> "Cool Message" (+ stack info)

'STANDARD NAMESPACES:
'-----------------------------
'Namespace,ID <-- Message
'-----------------------------
'STD.Class,1  <-- [ERROR: #{Module}::#{Method}] Init routine called on intitialised object.
'STD.Class,2  <-- [ERROR: #{Module}::#{Method}] Constructor "#{Method}" called on initialised object. Use on class instead.
'STD.Class,3  <-- [ERROR: #{Module}::#{Method}] Cannot Call Method "#{Method}" of #{Module} class.
'STD.Class,4  <-- [ERROR: #{Module}::#{Method}] Cannot Get Property "#{Method}" of #{Module} class.
'STD.Class,5  <-- [ERROR: #{Module}::#{Method}] Cannot Let Property "#{Method}" of #{Module} class.
'STD.Class,6  <-- [ERROR: #{Module}::#{Method}] Cannot Set Property "#{Method}" of #{Module} class.


Public Namespaces as Object

Public Sub Raise(Optional ByVal ErrorNamespace as string = "", Optional ByVal ErrorId as Integer = -1, Optional ByVal CustomMessage as String = "")
  Err.Number = -1
  Err.Source = "" 'Stack trace!
  Err.Description = CustomMessage

End Sub

Public Sub Namespace_Add(ByVal ErrorNamespace as string, Optional dict as Object)

End Sub
Public Sub Namespace_Remove(ByVal ErrorNamespace as string)

End Sub

Public Sub Error_Type_Add(ByVal ErrorNamespace as string, ByVal ErrorId as Integer, ByVal CustomMessage as String)

End Sub