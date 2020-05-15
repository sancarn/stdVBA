Implements IJSONSerializable

'Thinking this
Public children as Collection '<stdRegistry>

'This would be a decent interface!
Public WithEvents values as Dictionary

Function IJSONSerializable_ToJSON()
  'Defer JSON creation to Dictionary class
  Dim interface as IJSONSerializable
  set interface = getDescendents()
  IJSONSerializable_ToJSON = interface.ToJSON()
End Function

Function getDescendents() as Dictionary 'Dictionary<members:Dictionary, children:Collection<Dictionary<...>>>

End Function




Public Function values_GetValue(ByVal sKey as string, ByVal vVal as Variant)

End Function
Public Function values_LetValue(ByVal sKey as string, ByVal vVal as Variant)

End Function



'With stdRegistry.Create("HKCR/Software")
'  'Loop over keys
'  dim reg as stdRegistry
'  For each reg in  .children()
'    '...
'  Next
'
'  'Set some values
'  .values("SomeKeyName") = "hello world"
'  .values("SomeKeyName.type") = "REG_SZ"
'  
'End With













Private Enum RootKeys
  HKEY_CURRENT_USER = &H80000001
End Enum

Private Enum RegTypes
    REG_SZ = 1
    REG_EXPAND_SZ = 2
    REG_BINARY = 3
    REG_DWORD = 4
    REG_MULTI_SZ = 7
End Enum


Public Function ToJSON() as string
  Dim sTest as String
  sTest = "{"
  sTest = sTest & """members"":["
  sTest = sTest & members.map("ToJSON").join(",")
  sTest = sTest & "],"
  sTest = sTest & """children"":["
  sTest = sTest & children.map("ToJSON").join(",")
  sTest = sTest & "]"
  sTest = sTest & "}"
  ToJSON = sTest
End Function
Public Function IJSONSerializable_ToJSON() to string
  IJSONSerializable_ToJSON = ToJSON()
End Function



Sub EnumKeyDirs()
    'Debug.Print CreateObject("WScript.Shell").RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\Shell\AttachmentExecute")

    Dim temp As Object
    Dim strComputer As String
    Dim rPath As String
    Dim arrSubKeys, arrValues()
    Dim strAsk
        
    
    
    Set temp = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

    rPath = "SOFTWARE\Microsoft\Internet Explorer"
    temp.EnumKey HKEY_CURRENT_USER, rPath, arrSubKeys
    
    'temp.EnumValues HKEY_CURRENT_USER, rPath, arrValues, arrSubKeys
    If VarType(arrSubKeys) <> vbNull Then
        For Each strAsk In arrSubKeys
            Debug.Print strAsk
        Next
    End If
End Sub


Sub EnumKeyDir(ByVal sPath As String)
    Dim reg As Object: Set reg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    Dim arrKeys As Variant: reg.EnumKey HKEY_CURRENT_USER, sPath & IIf(Right(sPath, 1) = "\", "", "\"), arrKeys
    If VarType(arrKeys) <> vbNull Then
        For i = 0 To UBound(arrKeys)
            Debug.Print sPath & IIf(Right(sPath, 1) = "\", "", "\") & arrKeys(i)
        Next
    Else
        Debug.Print "... No keys found under this path ..."
    End If
End Sub
Sub EnumKeyVal(ByVal sPath As String, Optional iRoot As Long = HKEY_CURRENT_USER)
    Dim reg As Object: Set reg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    Dim arrKeys As Variant, arrTypes() As RegTypes, Value As Variant
    reg.EnumValues iRoot, sPath, arrKeys, arrTypes
    If VarType(arrKeys) <> vbNull Then
        For i = 0 To UBound(arrKeys)
            Select Case arrTypes(i)
            Case RegTypes.REG_SZ
                reg.GetStringValue iRoot, sPath, arrKeys(i), Value
                Debug.Print "  " & arrKeys(i) & ": (REG_SZ) """ & Value & """"
            Case RegTypes.REG_EXPAND_SZ
                reg.GetExpandedStringValue iRoot, sPath, arrKeys(i), Value
                Debug.Print "  " & arrKeys(i) & ": (REG_EXPAND_SZ) """ & Value & """"
            
            Case RegTypes.REG_MULTI_SZ
                reg.GetMultiStringValue iRoot, sPath, arrKeys(i), Value
                Dim s, cs As String
                For Each s In Value
                    cs = cs & """" & s & """" & ", "
                Next
                Debug.Print "  " & arrKeys(i) & ": (REG_MULTI_SZ) [" & left(cs, Len(cs) - 2) & "]"
                
            Case RegTypes.REG_BINARY
                Debug.Print "  " & arrKeys(i) & ": #BINARY"
            Case RegTypes.REG_DWORD
                reg.GetDWORDValue iRoot, sPath, arrKeys(i), Value
                Debug.Print "  " & arrKeys(i) & ": " & Value
            Case Else
                Debug.Print "  " & arrKeys(i) & ": UNKNOWN"
            End Select
        Next
    Else
        Debug.Print "... No values found under this path ..."
    End If
End Sub

