VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "STD_Types_Regex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

private p_initialised as boolean
private p_pattern as string
private p_namesDict as object
private p_matchDict as object

Public Property Let Pattern(val as string)
  if not p_initialised then
    
  else
    
  end if
End Property

Public Function Create(ByVal pattern as string, ByVal flags as string) as STD_Types_Regex
  Set Create = new STD_Types_Regex
  
  
  Create.init()
End Function
Friend Sub init()
  p_initialised = true
End Sub