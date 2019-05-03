VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JSON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute Value.VB_UserMemId = 0

Private pValue as Variant
Private pType as String


Public Property Get Value() as Variant
  
End Property
Public Property Let Value(v as variant)

End Property

Public Property Get Photo() as IPhoto
  if pType = "CF_BITMAP" then
    
  elseif pType = "CF_METAFILEPIC" then

  else
    Err.Raise(1,"stdClipboard::Photo","Invalid clipboard format")
  end if
End Property

Public Property Set Photo(v as IPhoto)
  'Set format based on passed photo v
End Property

Public Function GetBytes() as Byte()
  
End Function