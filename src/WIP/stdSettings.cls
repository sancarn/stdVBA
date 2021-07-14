VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = true
Private pSettingsCache As Object

Private pSettingsSheet As Worksheet

'TABLE pSettingsTable:
'  Of the form:
'    | Identity              | Key         | Value     | Tags          |
'    |-----------------------|-------------|-----------|---------------|
'    | System                | DataPtr     | 123456789 |               |
'    | <<domain>>\<<user>>   | Filter      | true      | Auto;Filter   |
'    | <<domain>>\<<user>>   | FreezePanes | false     | Freeze;Panes  |
'    | <<domain>>\<<user>>   | Access      | admin     | Access;System |
'    | <<domain>>\<<user>>   | LoggedIn    | 20190924  | Access;System |
Private pSettingsTable As ListObject

'STRING pIdentity
'  Of the form:
'    "<domain>\<user>"
Private pIdentity As String
Private pInitialised As Boolean
Public Event OnChange(ByVal sIdentity As String, ByVal sKey As String, ByVal Value As Variant)

'TODO:
'* Use custom XML parts for storing settings instead of sheets.

Public Function Create(Optional sSheetName As String = "shStdSettings", Optional sTableName As String = "StdSettingsTable") As StdSettings
    If pSettingsCache Is Nothing Then Set pSettingsCache = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(pSettingsCache(sSheetName & ";" & sTableName)) Then
        Set Create = pSettingsCache(sSheetName & ";" & sTableName)
        Exit Function
    End If
    
    
    Dim SettingsSheet As Worksheet
    Dim SettingsTable As ListObject
    
    'Try to get existing SettingsSheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = sSheetName Then
            Set SettingsSheet = ws
            Exit For
        End If
    Next
    
    'Create SettingsSheet if not created already
    If SettingsSheet Is Nothing Then
        Set SettingsSheet = ThisWorkbook.Sheets.Add()
        With SettingsSheet
            'Set name
            .Name = sSheetName
            
            ''Make sheet invisible
            .Visible = xlSheetVeryHidden
            
            'Set codename
            On Error Resume Next
                ThisWorkbook.VBProject.VBComponents(SettingsSheet.CodeName).Name = sSheetName
            On Error GoTo 0
            
            'Initialise table list object
            SettingsSheet.Range("A1:D1").Value = Split("Identity;Key;Value;Tags", ";")
            Dim tbl As ListObject
            Set tbl = .ListObjects.Add(xlSrcRange, SettingsSheet.Range("$A$1:$D$2"), , xlYes)
            tbl.Name = sTableName
        End With
    End If
    
    'Get list object
    Set SettingsTable = SettingsSheet.ListObjects(sTableName)
    
    'Create settings object
    Dim ret As StdSettings
    Set ret = New StdSettings
    Call ret.Init(SettingsSheet, SettingsTable)
    
    'Cache ret
    Set pSettingsCache(sSheetName & ";" & sTableName) = ret
    
    'Return settings object
    Set Create = ret
End Function

'Initialise object
Friend Sub Init(Sheet As Worksheet, Table As ListObject)
    Set pSettingsSheet = Sheet
    Set pSettingsTable = Table
    pIdentity = GetIdentity()
    pInitialised = True
End Sub

Public Function Export() As Boolean
    If pInitialised Then
        
    Else
        'Error
    End If
End Function

Public Function Import() As Boolean
    If pInitialised Then
        
    Else
        'Error
    End If
End Function

'GETTING SETTINGS
Public Property Get system(key As String) As Range
    If pInitialised Then
        Set system = GetSettingsRange("System", key)
    Else
        'Error
    End If
End Property

Public Property Get user(key As String) As Range
    If pInitialised Then
        Set user = GetSettingsRange(pIdentity, key)
        Exit Property
    Else
        'Error
    End If
End Property

'SETTING SETTINGS
Public Property Let system(key As String, val As Variant)
    If pInitialised Then
        GetSettingsRange("System", key).Value = val
        RaiseEvent OnChange("System", key, val)
        Exit Property
    Else
        'Error
    End If
End Property

Public Property Let user(key As String, val As Variant)
    If pInitialised Then
        GetSettingsRange(pIdentity, key).Value = val
        RaiseEvent OnChange(pIdentity, key, val)
        Exit Property
    Else
        'Error
    End If
End Property


Private Function GetIdentity() As String
    Dim identity As String
    GetIdentity = Environ("USERDNSDOMAIN") & "\" & Environ("USERNAME")
End Function

Private Function GetSettingsRange(ByVal sIdentity As String, ByVal sKey As String) As Range
    Dim arr As Variant
    arr = pSettingsTable.DataBodyRange.Value
    
    'Try to find property in list
    For i = 1 To UBound(arr, 1)
        If arr(i, 1) = sIdentity Then
            If arr(i, 2) = sKey Then
                Set GetSettingsRange = pSettingsSheet.Cells(i + 1, 3)
                Exit Function
            End If
        End If
    Next
    
    'If it doesn't exist, add it
    Dim lr As ListRow
    Set lr = pSettingsTable.ListRows.Add()
    lr.Range.Value = Array(sIdentity, sKey, "", "")
    Set GetSettingsRange = lr.Range(1, 3)
    Exit Function
End Function

