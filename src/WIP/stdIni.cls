VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private IniPath As String
Public Data As Object
Public Comments As Object
Private Initialised As Boolean
Public AutoSave As Boolean

Const DefaultIniPath As String = "H:\MailSettings.ini"
Const NewLineIdentifier As String = "3512db75-c35b-447b-b79a-386c905237de"
Const CommentSectionIdentifier As String = "6c843687-d1e5-4960-8f92-adf6266ec609"
Const CommentFooter As String = "2a1b354a-eca4-4095-908c-c76570d1c8a6"
Const DefaultSection As String = ""

Private Sub Class_Initialize()
    IniPath = DefaultIniPath
End Sub

'Readonly property path
Public Property Get Path() As String
  Path = IniPath
End Property

'Constructor
Function From(sIniPath As String) As Ini
    Dim x As Ini
    Set x = New Ini
    Call x.init(sIniPath)
    Set From = x
End Function
Friend Sub init(sIniPath As String)
  IniPath = sIniPath
  Call Read
  Initialised = True
End Sub

Public Sub Puts(key As String, val As String, Optional section As String = DefaultSection)
  Data(section)(key) = val
End Sub
Public Function Gets(key As String, Optional section As String = DefaultSection)
  Gets = Data(section)(key)
End Function

Public Sub Save()
  'Get file num
  Dim ff As Long: ff = VBA.FileSystem.FreeFile()
  
  'Open file for writing
  Open IniPath & ".tmp" For Output As ff
  
  'Loop over keys
  Dim vSection As Variant, vKey As Variant
  For Each vSection In Data.keys()
    'If vSection is an object then error
    If IsObject(vSection) Then GoTo SectionObjectError
    
    'Print comments if any exist
    If CStr(Comments(vSection)(CommentSectionIdentifier)) <> "" Then Print #ff, Comments(vSection)(CommentSectionIdentifier)
    
    'Print section header
    Print #ff, "[" & vSection & "]"
    
    For Each vKey In Data(vSection).keys()
      'If key or value is an object then error
      If IsObject(vKey) Then GoTo KeyObjectError
      If IsObject(Data(vSection)(vKey)) Then GoTo ValObjectError
      
      'Print comments if any exist
      If CStr(Comments(vSection)(vKey)) <> "" Then Print #ff, Comments(vSection)(vKey)
      
      'Print key,value pair
      If CStr(Data(vSection)(vKey)) <> "" Then Print #ff, vKey & "=" & Data(vSection)(vKey)
    Next
  Next
  
  'Write footer data if provided
  If CStr(Comments(CommentFooter)) <> "" Then Print #ff, Comments(CommentFooter)
  Close #ff
  
  'Write data to main file
  Kill IniPath
  Name IniPath & ".tmp" As IniPath
  Exit Sub
SectionObjectError:
  Err.Raise 100, "Ini::Save", "SaveError: Error in section header. Ini file specification does not support object serialisation!"
  Close #ff
  Kill IniPath & ".tmp"
  Exit Sub
KeyObjectError:
  Err.Raise 100, "Ini::Save", "SaveError: Error in key of section " & vSection & " header. Ini file specification does not support object serialisation!"
  Close #ff
  Kill IniPath & ".tmp"
  Exit Sub
ValObjectError:
  Err.Raise 100, "Ini::Save", "SaveError: Error in value of " & vSection & "!" & vKey & ". Ini file specification does not support object serialisation!"
  Close #ff
  Kill IniPath & ".tmp"
  Exit Sub
End Sub


Private Sub Read()
  'Initialise data and comments
  Set Data = CreateObject("Scripting.Dictionary")
  Set Comments = CreateObject("Scripting.Dictionary")
  
  'Get file num
  Dim ff As Long: ff = VBA.FileSystem.FreeFile()
  
  'Error tracking
  Dim iLineNum As Long: iLineNum = 0
  
  'Open file
  On Error GoTo NoFile
    Open IniPath For Input As ff
  On Error GoTo 0
  
  'Setup
  Dim sLine As String, sSection As String, sConcatComment As String
  
  'Set default section
  sSection = DefaultSection
  
  While Not EOF(ff)
    'Increment line for tracking
    iLineNum = iLineNum + 1
    
    'Get line as string
    Line Input #ff, sLine
    
    'Simple parsing of line
    '*************************
    'Comment:
    If sLine Like ";*" Then
      'Concatenate comments ready for writing to file
      sConcatComment = sConcatComment & IIf(sConcatComment = "", "", vbCrLf) & sLine
    
    'Section header
    ElseIf sLine Like "[[]*[]]" Then
      'Get section name
      sSection = Mid(sLine, 2, Len(sLine) - 2)
      
      'Create section name if not present
      If IsEmpty(Data(sSection)) Then Set Data(sSection) = CreateObject("Scripting.Dictionary")
      
      'Create comments for section
      If IsEmpty(Comments(sSection)) Then Set Comments(sSection) = CreateObject("Scripting.Dictionary")
      
      'Set comments of section to sConcatComment, then reset it
      Comments(sSection)(CommentSectionIdentifier) = sConcatComment
      sConcatComment = ""
      
    'Potentially data?
    ElseIf InStr(1, sLine, "=") > 0 Then
      'Split data from value
      Dim v As Variant: v = Split(sLine, "=", 2)
      
      'Set the data of the section to the value
      Data(sSection)(v(0)) = v(1)
      
      'Set the comments for this value and reset
      Comments(sSection)(v(0)) = sConcatComment
      sConcatComment = ""
    
    'Throw parsing error
    Else
      'Ignore blank lines:
      If sLine <> "" Then
        Err.Raise 100, "Ini::Read()", "ParseError: Unknown data format at line " & iLineNum & " of ini file."
      End If
    End If
  Wend
  
  'Write final comments (if any)
  Comments(CommentFooter) = sConcatComment
  Close #ff
  Exit Sub

NoFile:
  Set Data(DefaultSection) = CreateObject("Scripting.Dictionary")
  Set Comments(DefaultSection) = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
  If AutoSave Then Call Save
End Sub
