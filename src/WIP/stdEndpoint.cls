VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdEndpoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

#If VBA7 Then
  Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
  Private Declare PtrSafe Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
#Else
  Private Enum LongPtr
    [_]
  End Enum
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
  Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
#End If

Private Type TSingleton
  apiObjects As Collection
End Type
Private Type TThis
  Singleton As TSingleton
End Type
Private This as TThis

'Creates a shared object
'@param apiObject - Object to share
'@param name - Optional name of the connection
'@returns - The shared object
Public Function CreateSharedObject(ByVal apiObject as object, Optional ByVal name as string = "Main") As Object
  If This.Singleton.apiObjects is nothing then set this.singleton.apiObjects = new Collection
  Call This.Singleton.apiObjects.add(apiObject)
  
  Dim AppName as string: AppName = "VBA-HWND: 0x" & Hex(Application.Hwnd)
  Dim Section as string: Section = Workbook.Name
  Call SaveSetting(AppName, Section, name, ObjPtr(apiObject))
  
  set CreateSharedObject = apiObject
End Function

'Returns a shared object
'@param wb - Workbook to create connection for
'@param name - Optional name of the connection
'@returns - Connection object
Public Function CreateConnection(ByVal wb as Workbook, Optional ByVal name as string = "Main") as Object
  Dim app as Application: Set app = wb.Application
  Dim AppName as string: AppName = "VBA-HWND: 0x" & Hex(app.Hwnd)
  Dim Section as string: Section = wb.Name
  Dim apiObject as object
  Dim ptr as LongPtr: ptr = CLngPtr(GetSetting(AppName, Section, name, 0))
  If ptr = 0 Then
    Set apiObject = Nothing
  Else
    Set apiObject = Deref(ptr)
  End If
  
  Set CreateConnection = apiObject
End Function


'Deref pointer to IUnknown helper
'@param memAddress - Pointer to IUnknown
'@returns - IUnknown
'@devRemark DO NOT RETURN `result`
Private Function Deref(ByVal memAddress As LongPtr) As IUnknown
  Dim result As IUnknown
  CopyMemory result, ByVal VarPtr(memAddress), LenB(memAddress)
  Set Deref = result                             'Set and add ref
  ZeroMemory result, LenB(memAddress)
End Function

'Converts a string to a LongPtr
'@param str - String to convert
'@returns - LongPtr
Private Function CLngPtr(BYVal str as string) as LongPtr
  #If VBA7 Then
    CLngPtr = CLngLng(str)
  #Else
    CLngPtr = CLng(str)
  #End If
End Function

'TODO: May be better using stdCOM and ROT instead. Will still need to store the ROT name in the registry, in case of state loss.
'* https://www.mrexcel.com/board/threads/reference-and-remotely-manipulate-userforms-loaded-in-seperate-workbooks-or-in-seperate-excel-instances-via-file-monikers.1161038/
'* 