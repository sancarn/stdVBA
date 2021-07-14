VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = true
Public stack As stdArray
Public RaiseClient As Object 'Allow Raise through a custom userform. Userform should expose a Raise(sMessage, Criticality, Title, StackTrace) method.
Public Event OnStackAdd(ByRef oStack as stdArray, ByRef oElement as object)
Public Event OnStackPop(ByRef oStack as stdArray, ByRef oElement as object)


'TODO: Implement a ResumeNext() function for skipping past errors that are known, likely to use stdSentry
'With stdError.ResumeNext()
'    '...
'End With
'If stdError.Errors.Length > 0 Then
'    
'End If
'Implementation detail: set ResumeNext = stdSentry.CreateFromMethods(stdError,"zprot_RNOn","zprot_RNOff")

'TODO: Tracking performance of function calls on the stack.
'         stdError.TrackPerformance = true/false
'



Public Sub RefreshStack()
    Set stack = stdArray.Create()
End Sub

'TODO: It'd be cool if we could look backwards at the stack in order to get the calling object name, instead of providing it in sMethodName.
Public Sub AddStack(ByVal sMethodName As String)
    If stack Is Nothing Then Set stack = stdArray.Create()
    Dim oElement as object: set oElement = CreateDict("Name",sMethodName)
    RaiseEvent OnStackAdd(stack, oElement)
    stack.Push oElement
End Sub

Public Sub PopStack()
    If stack Is Nothing Then
        Raise "stdError::StackPop() - No stack created", vbCritical
    Else
        Dim oElement as object: set oElement = stack.Pop()
        RaiseEvent OnStackPop(stack, oElement)
    End If
End Sub


Public Function Raise(Optional ByVal sMessage As String = "", Optional ByVal Criticality As VBA.VbMsgBoxStyle = VBA.VbMsgBoxStyle.vbExclamation, Optional ByVal Title As String = "VBA Error", Optional ByVal isCritical As Boolean = True) As VBA.VbMsgBoxResult
    'Build stack trace if available
    Dim sStackTrace As String
    sStackTrace = GetTrace()
    
    '
    If RaiseClient Is Nothing Then
        'Start full message
        Dim sFullMessage As String
        sFullMessage = "Error in routine """ & sMessage & """" & vbCrLf & sStackTrace
        
        'Return and raise
        Raise = MsgBox(sFullMessage, Criticality, Title)
    Else
        On Error GoTo ErrorOccurred:
            Raise = RaiseClient.Raise(sMessage, Criticality, Title, sStackTrace)
        On Error GoTo 0
    End If
    
    'Stop process if critical
    If isCritical Then
      'Reset stack
      Me.RefreshStack
      
      End
    End If
    
    Exit Function
ErrorOccurred:
    'Ensure the error is raised
    Set RaiseClient = Nothing
    Me.AddStack "stdError::Raise()"
        Raise = Raise("stdError::Raise() Error in oMsgClient::Raise() " & Err.Description)
    Me.PopStack
End Function


'TODO: If we can get the stack trace directly by walking the VBA stack this would be amazing... Instead we currently have to manually add and remove method names to the stack.
Public Function GetTrace() as String
    'Build stack trace if available
    Dim sStackTrace As String
    If Not stack Is Nothing Then
        sStackTrace = "Trace:" & vbCrLf
        
        Dim i As Long
        For i = 1 To stack.Length
            sStackTrace = sStackTrace & Space((i - 1) * 3) & "|- " & stack.item(i)("Name") & vbCrLf
        Next
    End If

    GetTrace = sStackTrace
End Function

Private Function CreateDict(ParamArray args() As Variant)
  'Create dictionary
  Dim o As Object: Set o = CreateObject("Scripting.Dictionary")
  
  'Loop over parameters and create dictionary
  Dim i As Long
  For i = LBound(args) To UBound(args) Step 2
      'If is object then set
      If IsObject(args(i + 1)) Then
          Set o(args(i)) = args(i + 1)
      Else
          o(args(i)) = args(i + 1)
      End If
  Next
    
  'Return dictionary
  Set CreateDict = o
End Function
