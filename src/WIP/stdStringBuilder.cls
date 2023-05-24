VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'StringBuilder
'
'StringBuilder is specifically designed to build large strings which are mainly made from boiler plate code.
'Instead of code like this:
'
'    Dim str as string
'    str = str & "<div className=""MarkdownEditor"">" & vbCrLf
'    str = str & "  <h3>Input</h3>" & vbCrLf
'    str = str & "  <label htmlFor=""markdown-content"">" & vbCrLf
'    str = str & "    Enter some markdown" & vbCrLf
'    str = str & "  </label>" & vbCrLf
'    str = str & "  <textarea" & vbCrLf
'    str = str & "    id=""markdown-content""" & vbCrLf
'    str = str & "    onChange=""" & handleChange & """" & vbCrLf & _ " & vbCrLf
'    str = str & "    defaultValue=""" & state.value & """" & vbCrLf
'    str = str & "  />" & vbCrLf
'    str = str & "  <h3>Output</h3>" & vbCrLf
'    str = str & "  <div" & vbCrLf
'    str = str & "    className=""content""" & vbCrLf
'    str = str & "    dangerouslySetInnerHTML=" & getRawMarkup() & vbCrLf
'    str = str & "  />" & vbCrLf
'    str = str & "</div>"
'    renderHTML(str)
'
'You can now use code like this:
'    'IMPORTANT!!! Only Object (aka "IDispatch") can use square bracket syntax!
'    Dim sb as Object
'    set sb = StringBuilder.Create()
'    sb.TrimBehaviour = RTrim
'
'    'Inject variables into string
'    sb.InjectionVariables.add "{this.handleChange}", handleChange
'    sb.InjectionVariables.add "{this.state.value}", state.value
'    sb.InjectionVariables.add "{this.getRawMarkup()}", getRawMarkup()
'
'    'Build string
'    sb.[<div className="MarkdownEditor">                 ]
'    sb.[  <h3>Input</h3>                                 ]
'    sb.[  <label htmlFor="markdown-content">             ]
'    sb.[    Enter some markdown                          ]
'    sb.[  </label>                                       ]
'    sb.[  <textarea                                      ]
'    sb.[    id="markdown-content"                        ]
'    sb.[    onChange="{this.handleChange}"               ]
'    sb.[    defaultValue="{this.state.value}"            ]
'    sb.[  />                                             ]
'    sb.[  <h3>Output</h3>                                ]
'    sb.[  <div                                           ]
'    sb.[    className="content"                          ]
'    sb.[    dangerouslySetInnerHTML={this.getRawMarkup()}]
'    sb.[  />                                             ]
'    sb.[</div>                                           ]
'    renderHTML(sb)
'
'Which significantly improves maintainability of the code.
Option Explicit

' must be >= 2
Private Const DEFAULT_MINIMUM_CAPACITY As Long = 16

' StaticStringBuilder from https://github.com/sihlfall/vba-stringbuilder
' Code copied in here in order to keep everything in one file.
Private Type SsbTy
  Active As Integer            ' index of the currently active buffer (0 or 1)
  Buffer(0 To 1) As String     ' .Buffer(.Active) is the currently active buffer
  Capacity As Long             ' current allocated capacity in characters
  Length As Long               ' current length of the string, in characters
  MinimumCapacity As Long      ' minimum capacity set (>= 2)
End Type

Private InjectionVariablesDictionary As Object
Private Tainted As Boolean
Private ProcessedString As String
Private Ssb As SsbTy

Public JoinStr As String
Public TrimBehaviour As TrimBehaviourEnum
Public Enum TrimBehaviourEnum
  NoTrim
  LTrim
  RTrim
  Trim
End Enum

Public Property Get InjectionVariables() As Object
  If InjectionVariablesDictionary Is Nothing Then Set InjectionVariablesDictionary = CreateObject("Scripting.Dictionary")
  Set InjectionVariables = InjectionVariablesDictionary
End Property

Public Property Get MinimumCapacity() As Long
  MinimumCapacity = Ssb.MinimumCapacity
End Property
Public Property Let MinimumCapacity(desiredMinimumCapacity As Long)
  If desiredMinimumCapacity >= 2 Then Ssb.MinimumCapacity = desiredMinimumCapacity Else Ssb.MinimumCapacity = 2
End Property

Public Property Get Str() As String
Attribute Str.VB_UserMemId = 0
  If Tainted Then Call RefreshFromRaw
  Str = ProcessedString
End Property
Public Property Let Str(s As String)
  Dim sLength As Long
  With Ssb
    sLength = Len(s)
    ' We accept only 100% space overhead.
    If .Capacity < sLength Or (2 * sLength < .Capacity And .Capacity > .MinimumCapacity) Then ClearReserve sLength
    If sLength > 0 Then Mid$(.Buffer(.Active), 1, sLength) = s
    .Length = sLength
  End With
  
  Tainted = True
End Property

'Appends the string parsed to the main string Str
Public Function Append(s As String) As Variant
Attribute Append.VB_UserMemId = -5
  Dim Length As Long, lengthJoin As Long, nRequired As Long
  
  'Trim based on TrimBehaviour
  Select Case TrimBehaviour
    Case TrimBehaviourEnum.LTrim
      s = VBA.LTrim(s)
    Case TrimBehaviourEnum.RTrim
      s = VBA.RTrim(s)
    Case TrimBehaviourEnum.Trim
      s = VBA.Trim(s)
  End Select
  
  'Build String
  With Ssb
    Length = Len(s): lengthJoin = Len(JoinStr)
    If Length + lengthJoin = 0 Then Exit Function
    nRequired = .Length + Length + lengthJoin
    If nRequired > .Capacity Then SwitchToLargerBuffer nRequired
    If lengthJoin > 0 Then Mid$(.Buffer(.Active), .Length + 1, lengthJoin) = JoinStr
    If Length > 0 Then Mid$(.Buffer(.Active), .Length + lengthJoin + 1, Length) = s
    .Length = nRequired
  End With
  
  'Set tainted
  Tainted = True
End Function

'Creates an instance of stdStringBuilder
Public Function Create() As stdStringBuilder
  Set Create = New stdStringBuilder
End Function

Public Function Test()
  Dim sb As Object
  Set sb = Create()
  sb.JoinStr = "-"
  sb.Str = "Start"
  sb.TrimBehaviour = RTrim
  sb.InjectionVariables.Add "@1", "cool"
  sb.[This is a really cool multi-line    ]
  sb.[string which can even include       ]
  sb.[symbols like " ' # ! / \ without    ]
  sb.[causing compiler errors!!           ]
  sb.[also this has @1 variable injection!]
  Test = sb.Str = "Start-This is a really cool multi-line-string which can even include-symbols like "" ' # ! / \ without-causing compiler errors!!-also this has cool variable injection!"
End Function

'==============================
' PRIVATE METHODS
'==============================

Private Sub RefreshFromRaw()
  With Ssb
    ProcessedString = Left$(.Buffer(.Active), .Length)
  End With
  
  'Replace keys
  Dim key
  If Not InjectionVariablesDictionary Is Nothing Then
    For Each key In InjectionVariablesDictionary.Keys()
      ProcessedString = VBA.Replace(ProcessedString, key, InjectionVariables(key))
    Next
  End If
  
  'No longer tainted
  Tainted = False
End Sub

Private Sub ClearReserve(ByVal nRequired As Long)
  ' Clear everything
  ' Allocate a buffer that is able to hold nRequired characters.
  ' The new buffer size is calculated by repeatedly growing the minimum capacity by 50%.
  With Ssb
    .Active = 0
    .Length = 0
    .Buffer(1) = vbNullString
    If nRequired > 0 Then
      .Capacity = .MinimumCapacity
      Do
        If .Capacity >= nRequired Then Exit Do
        .Capacity = .Capacity + .Capacity \ 2
      Loop
      .Buffer(0) = String(.Capacity, 0)
    Else
      .Capacity = 0
      .Buffer(0) = vbNullString
    End If
  End With
  Tainted = False
  ProcessedString = vbNullString
End Sub

Private Sub SwitchToLargerBuffer(ByVal nRequired As Long)
  ' Allocate buffer that is able to hold nRequired characters.
  ' The new buffer size is calculated by repeatedly growing the current size by 50%.
  ' Copy string over to the new buffer.
  ' Deallocate the old buffer.
  With Ssb
    If .Capacity < .MinimumCapacity Then .Capacity = .MinimumCapacity
    Do
      If .Capacity >= nRequired Then Exit Do
      .Capacity = .Capacity + .Capacity \ 2
    Loop
    .Buffer(1 - .Active) = String(.Capacity, 0)
    Mid$(.Buffer(1 - .Active), 1, .Length) = .Buffer(.Active)
    .Buffer(.Active) = vbNullString
    .Active = 1 - .Active
  End With
End Sub

Private Sub Class_Initialize()
  Ssb.MinimumCapacity = DEFAULT_MINIMUM_CAPACITY
  JoinStr = vbCrLf
End Sub


