VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Structure of a stdRegistry tree
'stdReg<Key>
'|- Keys
'|  |- stdReg<Key>
'|  |- |- Keys
'|  |  |- Items
'|  |  |  |- stdReg<Item>
'|  |  |  |- stdReg<Item>
'|  |  |  |- stdReg<Item>
'|  |  |  |- stdReg<Item>
'|  |- stdReg<Key>
'|  |- |- Keys
'|  |  |- Items
'|  |  |  |- stdReg<Item>
'|  |- stdReg<Key>
'|  |- |- Keys
'|  |  |- Items
'|  |  |  |- stdReg<Item>
'|  |  |  |- stdReg<Item>
'|- Items
'|  |- stdReg<Item>
'|  |- stdReg<Item>
'|  |- stdReg<Item>
'|  |- stdReg<Item>
'
'Typically in RegEdit Items are instead called "Values" but given that Values contain a Value which might be an array, I figured it might be confusing. Thus rename to items.
'However some of the terminology in this source code refers to values




Public Enum ERegistryType
    iCalculate = 0
    iKey = 1
    iItem = 2
End Enum

Public Enum ERegistryValueType
    Value_None = 0               'REG_NONE
    Value_String = 1             'REG_SZ
    Value_String_WithEnvVars = 2 'REG_EXPAND_SZ  'use the ExpandEnvironmentStrings function to obtain full string
    Value_Binary = 3             'REG_BINARY
    Value_DWORD = 4              'REG_DWORD
    Value_DWORD_LE = Value_DWORD 'REG_DWORD_LITTLE_ENDIAN
    Value_DWORD_BE = 5           'REG_DWORD_BIG_ENDIAN
    Value_Link = 6               'REG_LINK
    Value_String_Array = 7       'REG_MULTI_SZ
    Value_QWORD = 11             'REG_QWORD
    Value_QWORD_LE = Value_QWORD 'REG_QWORD_LITTLE_ENDIAN
End Enum

Public Enum ERegistryAccess
    PERM_ALL_ACCESS = &HF003F
    PERM_CREATE_LINK = &H20
    PERM_CREATE_SUB_KEY = &H4
    PERM_ENUMERATE_SUB_KEYS = &H8
    PERM_EXECUTE = &H20019
    PERM_NOTIFY = &H10
    PERM_QUERY_VALUE = &H1
    PERM_READ = &H20019
    PERM_SET_VALUE = &H2
    PERM_WOW64_32KEY = &H200
    PERM_WOW64_64KEY = &H100
    PERM_WRITE = &H20006
End Enum

Private Enum ERegistryRoot
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_CURRENT_CONFIG = &H80000005
End Enum

Private Enum ESystemError
  ERROR_SUCCESS = &H0
  ERROR_FILE_NOT_FOUND = &H2
End Enum




#If VBA7 Then
    Private Declare PtrSafe Function RegOpenKeyExA Lib "advapi32.dll" (ByVal key As ERegistryRoot, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As LongPtr) As Long
    Private Declare PtrSafe Function RegQueryValueExA Lib "advapi32.dll" (ByVal hKey As LongPtr, ByVal sValueName As String, ByVal lpReserved As LongPtr, ByVal lpType As LongPtr, ByVal lpData As LongPtr, ByRef lpcbData As Long) As Long
    Private Declare PtrSafe Function RegEnumKeyExA Lib "advapi32" (ByVal hKey As LongPtr, ByVal dwIndex As Long, ByVal lpName As String, ByRef lpcbName As Long, ByVal lpReserved As LongPtr, ByVal lpClass As LongPtr, ByVal lpsbClass As LongPtr, ByVal lpftLastWriteTime As LongPtr) As Long
    Private Declare PtrSafe Function RegEnumValueA Lib "advapi32" (ByVal hKey As LongPtr, ByVal dwIndex As Long, ByVal lpName As String, ByRef lpcbName As Long, ByVal lpReserved As LongPtr, ByVal lpType As LongPtr, ByVal lpData As LongPtr, ByVal lpcbData As LongPtr) As Long
    Private Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As LongPtr) As Long
    Private Declare PtrSafe Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As LongPtr, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As LongPtr, ByVal cbData As Long) As Long
#Else
    Private Enum LongPtr
      [_]
    End Enum
    Private Declare Function RegOpenKeyExA Lib "advapi32.dll" (ByVal key As ERegistryRoot, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As LongPtr) As Long
    Private Declare Function RegQueryValueExA Lib "advapi32.dll" (ByVal hKey As LongPtr, ByVal sValueName As String, ByVal lpReserved As LongPtr, ByVal lpType As LongPtr, ByVal lpData As LongPtr, ByRef lpcbData As Long) As Long
    Private Declare Function RegEnumKeyExA Lib "advapi32" (ByVal hKey As LongPtr, ByVal dwIndex As Long, ByVal lpName As String, ByRef lpcbName As Long, ByVal lpReserved As LongPtr, ByVal lpClass As LongPtr, ByVal lpsbClass As LongPtr, ByVal lpftLastWriteTime As LongPtr) As Long
    Private Declare Function RegEnumValueA Lib "advapi32" (ByVal hKey As LongPtr, ByVal dwIndex As Long, ByVal lpName As String, ByRef lpcbName As Long, ByVal lpReserved As LongPtr, ByVal lpType As LongPtr, ByVal lpData As LongPtr, ByVal lpcbData As LongPtr) As Long
    Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As LongPtr) As Long
    Private Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As LongPtr, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As LongPtr, ByVal cbData As Long) As Long
#End If

#If Win64 Then
  Private const NULL_PTR as LongLong = 0^
#Else
  Private Const NULL_PTR As Long = 0
#End If


'Internal state
Private Type TThis
  path As String
  Type As ERegistryType
  root As ERegistryRoot
  access As ERegistryAccess
  query As String
  queryHandle As LongPtr
  writeHandle As LongPtr
  name As String
  valtype As ERegistryValueType
End Type
Private This As TThis

'Create a new stdReg object from a path with required access
'@param path - Path to registry key or value
'@param access - Access required to open key i.e. `PERM_READ`, `PERM_WRITE`, `PERM_ALL_ACCESS`
'@param normalisePath - If `true`, then path will be normalised to use `"\"` as path separator
'@returns - `stdReg` object
Public Function Create(ByVal path As String, Optional ByVal access As ERegistryAccess = ERegistryAccess.PERM_READ, Optional ByVal normalisePath As Boolean = True) As stdReg
  Set Create = New stdReg
  Call Create.protInit(path, ERegistryType.iCalculate, access, normalisePath)
End Function

'Create a new stdReg object from a key path with required access
'@param path - Path to registry key
'@param access - Access required to open key i.e. `PERM_READ`, `PERM_WRITE`, `PERM_ALL_ACCESS`
'@param normalisePath - If `true`, then path will be normalised to use `"\"` as path separator
'@returns - `stdReg` object representing registry key
Public Function CreateFromKey(ByVal path As String, Optional ByVal access As ERegistryAccess = ERegistryAccess.PERM_READ, Optional ByVal normalisePath As Boolean = True) As stdReg
  Set CreateFromKey = New stdReg
  Call CreateFromKey.protInit(path, ERegistryType.iKey, access, normalisePath)
End Function

'Create a new stdReg object from a value path with required access
'@param path - Path to registry value
'@param access - Access required to open key i.e. `PERM_READ`, `PERM_WRITE`, `PERM_ALL_ACCESS`
'@param normalisePath - If `true`, then path will be normalised to use `"\"` as path separator
'@returns - `stdReg` object representing registry value
Public Function CreateFromItem(ByVal path As String, Optional ByVal access As ERegistryAccess = ERegistryAccess.PERM_READ, Optional ByVal normalisePath As Boolean = True) As stdReg
  Set CreateFromItem = New stdReg
  Call CreateFromItem.protInit(path, ERegistryType.iItem, access, normalisePath)
End Function

'Create a new stdReg object from a key path with required access
'@protected
'@param path - Path to registry key
'@param iType - Type of stdReg object to create i.e. `iKey`, `iItem`
'@param access - Access required to open key i.e. `PERM_READ`, `PERM_WRITE`, `PERM_ALL_ACCESS`
'@param normalisePath - If `true`, then path will be normalised to use `"\"` as path separator
Friend Sub protInit(ByVal path As String, Optional ByVal iType As ERegistryType = 0, Optional ByVal access As ERegistryAccess = ERegistryAccess.PERM_READ, Optional ByVal normalisePath As Boolean = True)
  Static roots As Variant: If IsEmpty(roots) Then roots = Array( _
    "HKEY_CLASSES_ROOT", HKEY_CLASSES_ROOT, _
    "HKEY_CURRENT_USER", HKEY_CURRENT_USER, _
    "HKEY_LOCAL_MACHINE", HKEY_LOCAL_MACHINE, _
    "HKEY_USERS", HKEY_USERS, _
    "HKEY_CURRENT_CONFIG", HKEY_CURRENT_CONFIG, _
    "HKCR", HKEY_CLASSES_ROOT, _
    "HKCU", HKEY_CURRENT_USER, _
    "HKLM", HKEY_LOCAL_MACHINE _
  )
  
  'In theory / is a valid character in a Key/Value name therefore this is optional.
  If normalisePath Then This.path = Replace(path, "/", "\")
  
  This.path = path
  This.Type = iType
  This.access = access
  
  'Obtain root and query for API
  Dim i As Long
  For i = 0 To UBound(roots) Step 2
    If path Like roots(i) & "*" Then
      This.root = roots(i + 1)
      This.query = Mid(This.path, Len(roots(i)) + 1)
      If left(This.query, 1) = "\" Then
        This.query = Mid(This.query, 2)
      End If
      Exit For
    End If
  Next
  
  'Try to open key, if fails, then assume value and try again
  If This.Type = iCalculate Then
    Select Case RegOpenKeyExA(This.root, This.query, 0, This.access, This.queryHandle)
      Case ERROR_SUCCESS
        This.Type = iKey
        Call RegCloseKey(This.queryHandle)
      Case Else
        This.Type = iItem
    End Select
  End If
  
  'Identify key location and value name
  If This.query <> "" Then
    Dim parts As Variant: parts = Split(This.query, "\")
    Dim iub As Long: iub = UBound(parts)
    If iub <> -1 Then
      This.name = parts(iub)
      
      'If value then trim name from key path
      If This.Type = iItem Then
        'Accessing default value via path
        If This.name = "(Default)" And normalisePath Then
          This.name = ""
          This.Type = iItem
        End If
        
        'Trim query based on name
        If iub <> 0 Then
          ReDim Preserve parts(0 To iub - 1)
          This.query = Join(parts, "\")
        Else
          This.query = ""
        End If
      End If
    Else
      This.query = ""
      This.name = ""
    End If
  Else
    This.name = This.path
    'Shorthands
    If This.name = "HKCR" Then This.name = "HKEY_CLASSES_ROOT"
    If This.name = "HKCU" Then This.name = "HKEY_CURRENT_USER"
    If This.name = "HKLM" Then This.name = "HKEY_LOCAL_MACHINE"
  End If
  
  'Open key
  If RegOpenKeyExA(This.root, This.query, 0, This.access, This.queryHandle) = ERROR_SUCCESS Then
      If This.Type = iItem Then This.valtype = pRegistryQueryValueType()
  Else
      Err.Raise 1, "stdReg", "ERROR: Opening registry key failed"
  End If
End Sub


'*******************
'* GENERIC METHODS *
'*******************

'Obtain name of registry key
'@returns - Name of registry key
Public Property Get name() As String
  If This.name = "" Then
    name = "(Default)"
  Else
    name = This.name
  End If
End Property

'Obtain type of registry object. Either `ERegistryType.iKey` or `ERegistryType.iItem`.
'@returns - Type of registry Either `ERegistryType.iKey` or `ERegistryType.iItem`.
Public Property Get RegType() As ERegistryType
  RegType = This.Type
End Property

'Obtain whether registry key/item is root
'@returns - Whether registry key/item is root
Public Property Get isRoot() As Boolean
  isRoot = This.query = ""
End Property

'Obtain the path of the registry key/item
'@returns - Path of the registry key/item
Public Property Get path() As String
  path = This.path & IIf(right(This.path, 9) <> "(Default)" And This.name = "", "(Default)", "")
End Property

'Obtain the parent of the registry key/item
'@returns - Parent of the registry key/item
Public Property Get Parent() As stdReg
  If isRoot Then
    Set Parent = Nothing
  Else
    Dim iLastSlash As Long: iLastSlash = InStrRev(This.path, "\")
    Dim sPath As String: sPath = left(This.path, iLastSlash - 1)
    Set Parent = stdReg.CreateFromKey(sPath)
  End If
End Property

'***************
'* KEY METHODS *
'***************

'Get all keys and items
'@returns Collection<stdReg> - Collection of stdReg keys and items
'@throws 1, "ERROR: Only stdReg keys have subkeys"
'@throws 1, "ERROR: Only stdReg keys have items"
Public Property Get children() As Collection
  Set children = Keys
  
  Dim o As stdReg
  For Each o In Items
    Call children.Add(o)
  Next
End Property

'Get all subkeys
'@returns Collection<stdReg> - Collection of stdReg keys
'@throws 1, "ERROR: Only stdReg keys have subkeys"
Public Property Get Keys() As Collection
  Set Keys = New Collection
  Select Case This.Type
    Case ERegistryType.iKey
      'Enumerate keys
      Dim i As Long: i = -1
      Do
        i = i + 1
        Dim iSize As Long: iSize = 256
        Dim sName As String: sName = Space(iSize)
        Dim result As Long: result = RegEnumKeyExA(This.queryHandle, i, sName, iSize, 0, 0, 0, 0)
        If result = ERROR_SUCCESS Then
          Call Keys.Add(CreateFromKey(This.path & "\" & left(sName, iSize)))
        End If
      Loop While result = ERROR_SUCCESS
    Case Else
      Err.Raise 1, "stdReg", "ERROR: Only stdReg keys have subkeys"
  End Select
End Property

'Get all items
'@returns Collection<stdReg> - Collection of stdReg items
'@throws 1, "ERROR: Only stdReg keys have items"
Public Property Get Items() As Collection
  Set Items = New Collection
  Select Case This.Type
    Case ERegistryType.iKey
      'Enumerate keys
      Dim i As Long: i = -1
      Do
        i = i + 1
        Dim iSize As Long: iSize = 256
        Dim sName As String: sName = Space(iSize)
        Dim result As Long: result = RegEnumValueA(This.queryHandle, i, sName, iSize, 0, 0, 0, 0)
        If result = ERROR_SUCCESS Then
          Call Items.Add(CreateFromItem(This.path & "\" & left(sName, iSize)))
        End If
      Loop While result = ERROR_SUCCESS
    case else
      Err.Raise 1, "stdReg", "ERROR: Only stdReg keys have items"
  End Select
End Property

'****************
'* ITEM METHODS *
'****************

'Get the item type of this stdReg Item
'@returns - Type of stdReg Item
'@throws 1, "ERROR: Only stdReg items have values"
Public Property Get ItemType() As ERegistryValueType
  If This.Type = iItem Then
    ItemType = This.valtype
  Else
    Err.Raise 1, "stdReg", "ERROR: Only stdReg items have values"
  End If
End Property

'Get / Set value
'@returns - Value of stdReg Item
'@throws 1, "stdReg", "ERROR: Unknown value type"
'@throws 2, "stdReg", "ERROR: Cannot get value (0x" & Hex(iStatus) & ")"
'@throws 3, "", "ERROR: Only stdReg items have values"
Public Property Get value() As Variant
  If This.Type <> ERegistryType.iItem Then Err.Raise 3, "", "ERROR: Only stdReg items have values"
  Dim iSize As Long: iSize = pRegistryQueryValueSize()
  Dim vRet As Variant, iStatus As Long
  Select Case This.valtype
    Case ERegistryValueType.Value_Binary
      Dim binary() As Byte
      ReDim binary(1 To iSize)
      iStatus = RegQueryValueExA(This.queryHandle, This.name, 0, 0, VarPtr(binary(1)), iSize)
      value = binary
    Case ERegistryValueType.Value_DWORD
      Dim DWORD As Long
      iStatus = RegQueryValueExA(This.queryHandle, This.name, 0, 0, VarPtr(DWORD), iSize)
      value = DWORD
    Case ERegistryValueType.Value_DWORD_BE
      Dim DWORD_BE As Long
      iStatus = RegQueryValueExA(This.queryHandle, This.name, 0, 0, VarPtr(DWORD_BE), iSize)
      'TODO: SWAP
      iStatus = 666
    Case ERegistryValueType.Value_None
      value = Empty
    Case ERegistryValueType.Value_QWORD
      Dim QWORD As Currency
      iStatus = RegQueryValueExA(This.queryHandle, This.name, 0, 0, VarPtr(QWORD), iSize)
      value = QWORD
    Case ERegistryValueType.Value_String, ERegistryValueType.Value_String_WithEnvVars, ERegistryValueType.Value_Link
      Dim bSZ() As Byte
      ReDim bSZ(1 To iSize)
      iStatus = RegQueryValueExA(This.queryHandle, This.name, 0, 0, VarPtr(bSZ(1)), iSize)
      
      Dim sRet As String, i As Long
      For i = 1 To iSize
          If bSZ(i) = 0 Then Exit For
          sRet = sRet & Chr(bSZ(i))
      Next
      
      'TODO: Expand vars(?)
      'If This.valtype = ERegistryValueType.Value_String_WithEnvVars Then
      '
      'End If
      
      value = sRet
    Case ERegistryValueType.Value_String_Array
      Dim bMULTISZ() As Byte
      ReDim bMULTISZ(1 To iSize)
      iStatus = RegQueryValueExA(This.queryHandle, This.name, 0, 0, VarPtr(bMULTISZ(1)), iSize)
      
      Dim sRetArray() As String, iItem As Long: iItem = 0
      ReDim sRetArray(0 To 0)
      For i = 1 To iSize
        If bMULTISZ(i) = 0 Then
          'If \0\0 then stop, else add item to array
          If bMULTISZ(i + 1) = 0 Then
            Exit For
          Else
            ReDim Preserve sRetArray(0 To UBound(sRetArray) + 1)
            iItem = iItem + 1
          End If
        Else
          'Concatenate
          sRetArray(iItem) = sRetArray(iItem) & Chr(bMULTISZ(i))
        End If
      Next
    Case Else
      Err.Raise 1, "stdReg", "ERROR: Unknown value type"
  End Select
  
  If iStatus <> 0 Then Err.Raise 2, "stdReg", "ERROR: Cannot get value (0x" & Hex(iStatus) & ")"
End Property
Public Property Let value(vNew As Variant)
  Call SetValue(vNew, This.valtype)
End Property

'Set value and change type
'@param vNew  - Value to set item to
'@param iType - Type of value to set
Public Sub SetValue(ByVal vNew As Variant, ByVal iType As ERegistryValueType)
  If This.Type <> ERegistryType.iItem Then Err.Raise 3, "", "ERROR: Only stdReg items have values"
  Call pRegistryInitWriteHandle
  
  Dim iStatus As Long
  Select Case iType
    Case ERegistryValueType.Value_Binary
      Dim bNew() As Byte: bNew = vNew
      Dim lb As Long: lb = LBound(bNew)
      Dim ub As Long: ub = UBound(bNew)
      iStatus = RegSetValueExA(This.writeHandle, This.name, 0, iType, VarPtr(bNew(lb)), ub - lb + 1)
    Case ERegistryValueType.Value_DWORD
      Dim DWORD As Long: DWORD =  vNew
      iStatus = RegSetValueExA(This.writeHandle, This.name, 0, iType, VarPtr(DWORD), LenB(DWORD))
    Case ERegistryValueType.Value_DWORD_BE
      Dim DWORD_BE As Long: DWORD_BE =  vNew
      'TODO: SWAP
      iStatus = 666
      if iStatus = ERROR_SUCCESS Then iStatus = RegQueryValueExA(This.queryHandle, This.name, 0, 0, VarPtr(DWORD_BE), LenB(DWORD))
    Case ERegistryValueType.Value_None
      iStatus = RegSetValueExA(This.writeHandle, This.name, 0, iType, 0, 0)
    Case ERegistryValueType.Value_QWORD
      Dim QWORD As Currency: QWORD = vNew
      iStatus = RegSetValueExA(This.writeHandle, This.name, 0, iType, VarPtr(QWORD), LenB(QWORD))
    Case ERegistryValueType.Value_String, ERegistryValueType.Value_String_WithEnvVars, ERegistryValueType.Value_Link
      Dim bSZ() As Byte: bSZ = vNew
      Dim ilb As Long: ilb = LBound(bSZ)
      Dim iub As Long: iub = UBound(bSZ)
      iStatus = RegSetValueExA(This.writeHandle, This.name, 0, iType, VarPtr(bSZ(ilb)), iub-ilb+1)
    Case ERegistryValueType.Value_String_Array
      'TODO: TBC 
      Err.Raise 1, "stdReg", "ERROR: Not implemented"
    Case Else
      Err.Raise 1, "stdReg", "ERROR: Unknown value type"
  End Select
  
  If iStatus <> 0 Then Err.Raise 2, "stdReg", "ERROR: Cannot set value (0x" & Hex(iStatus) & ")"
End Sub


'*******************
'* Private Helpers *
'*******************

'Initialise write handle
'@private
Private Sub pRegistryInitWriteHandle()
  If This.writeHandle = NULL_PTR Then
    If RegOpenKeyExA(This.root, This.query, 0, ERegistryAccess.PERM_WRITE, This.writeHandle) <> ERROR_SUCCESS Then
      Err.Raise 1, "stdReg", "ERROR: Insufficient permissions to write to this registry key."
    End If
  End If
End Sub

'Query value type
'@private
'@returns - Value type of registry item
Private Function pRegistryQueryValueType() As ERegistryValueType
  Dim iType As ERegistryValueType
  Call RegQueryValueExA(This.queryHandle, This.name, 0, VarPtr(iType), 0, 0)
  pRegistryQueryValueType = iType
End Function

'Query value size
'@private
'@returns - Size of registry item
Private Function pRegistryQueryValueSize() As Long
  Call RegQueryValueExA(This.queryHandle, This.name, 0, 0, 0, pRegistryQueryValueSize)
End Function

'Close handles
'@destructor
'@private
Private Sub Class_Terminate()
  Call RegCloseKey(This.queryHandle)
End Sub
