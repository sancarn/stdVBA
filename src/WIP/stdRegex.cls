VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdRegex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Adds new features:
'* Named capturing groups
'* Modern symbols e.g. \A and \Z symbols
'* ...?+ --> (?:...?)+
'
'The class works as follows:
'  * Input regex is transpiled to ES6 convention
'  * Regex is executed using VBScript
'  * Regex object returned with new features.


private p_initialised as boolean
private p_pattern as string
private p_cPattern as string
private p_flags as string
private p_namesDict as object
private p_matchDict as object
private p_regex as object

Public Property Let Pattern(val as string)
  if not Me is stdRegex then
    Call init(val,p_flags)
  else
    'ERR_RAISE("Please create a new class")
  end if
End Property
Public Property Let Flags(val as string)
  if not Me is stdRegex then
    Call init(p_pattern,val)
  else
    'ERR_RAISE("Please create a new class")
  end if
End Property

'stdRegex[/hello world/i]
Public Function Create2(ByVal rx as string)
  Attribute Create2.VB_UserMemID = -5
  'regex given like /asdf/flags
End Function

'stdRegex.create("hello world","i")
Public Function Create(ByVal pattern as string, ByVal flags as string) as stdRegex
  Set Create = new stdRegex
  Create.init(pattern,flags)
End Function

Friend Sub init(ByVal pattern as string, ByVal flags as string)
  p_pattern = pattern
  p_flags = flags
  set p_namesDict = CreateObject("Scripting.Dictionary")
  set p_matchDict = CreateObject("Scripting.Dictionary")
  set p_regex = CreateObject("VBScript.RegExp")

  'Named regexp used to detect capturing groups and named capturing groups
  Static NamedRegexp As Object
  If NamedRegexp Is Nothing Then
    Set NamedRegexp = CreateObject("VBScript.RegExp")
    NamedRegexp.pattern = "\((?:\?\<(.*?)\>)?"
    NamedRegexp.Global = True
  End If
  
  Static FreeSpace As Object
  If FreeSpace Is Nothing Then
    Set FreeSpace = CreateObject("VBScript.RegExp")
    FreeSpace.pattern = "\s+"
    FreeSpace.Global = True
  End If

  'TODO:...

  p_initialised = true
End Sub










Function RegexMatch(ByVal haystack As String, ByVal pattern As String, Optional ByVal options As String) As Object
  'Cache regexes for optimisation
  Static CachedRegex As Object
  Static CachedNames As Object
  If CachedRegex Is Nothing Then Set CachedRegex = CreateObject("Scripting.Dictionary")
  If CachedNames Is Nothing Then Set CachedNames = CreateObject("Scripting.Dictionary")
  
  'Named regexp used to detect capturing groups and named capturing groups
  Static NamedRegexp As Object
  If NamedRegexp Is Nothing Then
    Set NamedRegexp = CreateObject("VBScript.RegExp")
    NamedRegexp.pattern = "\((?:\?\<(.*?)\>)?"
    NamedRegexp.Global = True
  End If
  
  Static FreeSpace As Object
  If FreeSpace Is Nothing Then
    Set FreeSpace = CreateObject("VBScript.RegExp")
    FreeSpace.pattern = "\s+"
    FreeSpace.Global = True
  End If
  
  'If cached pattern doesn't exist, create it
  If Not CachedRegex(pattern) Then
    
    
    'Create names/capture group object
    Dim testPattern As String, oNames As Object
    testPattern = pattern
    testPattern = Replace(testPattern, "\\", "asdasd")
    testPattern = Replace(testPattern, "\(", "asdasd")
    
    'Store names for optimisation
    Set CachedNames(options & ")" & pattern) = NamedRegexp.Execute(testPattern)
    
    'Create new VBA valid pattern
    Dim newPattern As String
    newPattern = NamedRegexp.Replace(pattern, "(")
    
    'Create regexp from new pattern
    Dim oRegexp As Object
    Set oRegexp = CreateObject("VBScript.RegExp")
    
    'Set regex options
    Dim i As Integer
    For i = 1 To Len(flags)
        Select Case Mid(flags, i, 1)
            Case "i"
                oRegexp.ignoreCase = True
            Case "g"
                oRegexp.Global = True
            Case "x"
                newPattern = FreeSpace.Replace(newPattern, "(?:)")
            Case "s"
                newPattern = Replace(newPattern, "\\", "973ed556-6a75-45d6-b0c2-8c2d0e2431c9")
                newPattern = Replace(newPattern, "\.", "7ae2088d-1b1d-4ee1-8e38-956db49c12a2")
                newPattern = Replace(newPattern, ".", "(?:.|\s)")
                newPattern = Replace(newPattern, "7ae2088d-1b1d-4ee1-8e38-956db49c12a2", "\.")
                newPattern = Replace(newPattern, "973ed556-6a75-45d6-b0c2-8c2d0e2431c9", "\\")
            Case "m"
                oRegexp.MultiLine = True
        End Select
    Next
    
    'Set pattern
    oRegexp.pattern = newPattern
    
    'Store regex for optimisation
    
    Set CachedRegex(options & ")" & pattern) = oRegexp
  End If
  
  'Get matches object
  Dim oMatches As Object
  Set oMatches = CachedRegex(options & ")" & pattern).Execute(haystack)
  
  'Get names object
  Dim CName As Object
  Set CName = CachedNames(options & ")" & pattern)
  
  'Create dictionary to return
  Dim oRet As Object
  Set oRet = CreateObject("Scripting.Dictionary")
  
  'Fill dictionary with names and indexes
  '0 = Whole match
  '1,2,3,... = Submatch 1,2,3,...
  '"Count" stores the count of matches
  '"<<NAME>>" stores the match of a specified name
  For i = 1 To CName.Count
    oRet(i) = oMatches(0).Submatches(i - 1)
    If Not IsEmpty(CName(i - 1).Submatches(0)) Then oRet(CName(i - 1).Submatches(0)) = oMatches(0).Submatches(i - 1)
  Next i
  oRet(0) = oMatches(0)
  oRet("Count") = CName.Count
  Set RegexMatch = oRet
End Function