VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#If Mac Then
    Private pSettingsCache As Dictionary
#Else
    Private pSettingsCache As Object
#End If

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
Public Event OnChange(ByVal sIdentity As String, ByVal sKey As String, ByVal value As Variant)


'TODO:
'* Use custom XML parts for storing settings instead of sheets.

Public Function Create(Optional sSheetName As String = "shStdSettings", Optional sTableName As String = "StdSettingsTable") As stdSettings
    #If Mac Then
        If pSettingsCache Is Nothing Then Set pSettingsCache = New Dictionary
    #Else
        If pSettingsCache Is Nothing Then Set pSettingsCache = CreateObject("Scripting.Dictionary")
    #End If
    
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
                #If Mac Then
                    'Don't bother
                #Else
                    ThisWorkbook.VBProject.VBComponents(SettingsSheet.CodeName).Name = sSheetName
                #End If
            On Error GoTo 0
            
            'Initialise table list object
            SettingsSheet.Range("A1:D1").value = Split("Identity;Key;Value;Tags", ";")
            Dim tbl As ListObject
            Set tbl = .ListObjects.Add(xlSrcRange, SettingsSheet.Range("$A$1:$D$2"), , xlYes)
            tbl.Name = sTableName
        End With
    End If
    
    'Get list object
    Set SettingsTable = SettingsSheet.ListObjects(sTableName)
    
    'Create settings object
    Dim ret As stdSettings
    Set ret = New stdSettings
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
Public Property Get system(KEY As String) As Range
    If pInitialised Then
        Set system = GetSettingsRange("System", KEY)
    Else
        'Error
    End If
End Property

Public Property Get user(KEY As String) As Range
    If pInitialised Then
        Set user = GetSettingsRange(pIdentity, KEY)
        Exit Property
    Else
        'Error
    End If
End Property

'SETTING SETTINGS
Public Property Let system(KEY As String, val As Variant)
    If pInitialised Then
        GetSettingsRange("System", KEY).value = val
        RaiseEvent OnChange("System", KEY, val)
        Exit Property
    Else
        'Error
    End If
End Property

Public Property Let user(KEY As String, val As Variant)
    If pInitialised Then
        GetSettingsRange(pIdentity, KEY).value = val
        RaiseEvent OnChange(pIdentity, KEY, val)
        Exit Property
    Else
        'Error
    End If
End Property


Private Function GetIdentity() As String
    Dim identity As String
    #If Mac Then
        'Not going to get domain easily on a mac, so adding in the SandboxGroupId instead
        GetIdentity = Environ("DOTNET_SANDBOX_APPLICATION_GROUP_ID") & Application.PathSeparator & Environ("LOGNAME")
    #Else
        'recommend change "\" to Application.PathSeparator
        
        GetIdentity = Environ("USERDNSDOMAIN") & "\" & Environ("USERNAME")
    #End If
End Function

Private Function GetSettingsRange(ByVal sIdentity As String, ByVal sKey As String) As Range
    Dim arr As Variant
    arr = pSettingsTable.DataBodyRange.value
    
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
    lr.Range.value = Array(sIdentity, sKey, "", "")
    Set GetSettingsRange = lr.Range(1, 3)
    Exit Function
End Function

