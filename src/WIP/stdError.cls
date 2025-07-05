VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Type TThis
    stack as stdArray
    RaiseClient As Object 'Allow Raise through a custom userform. Userform should expose a Raise(sMessage, Criticality, Title, StackTrace) method.
End Type
Private This as TThis

'New Syntax:
'  sub test
'    With stdError.getSentry("test")
'      '...
'      .Raise "Some message"
'      '...
'    End With
'  end sub
Public Function getSentry(ByVal sSubName As String, ParamArray args() as variant) As Object 'stdSentry
  If TypeName(stdSentry) = "stdSentry" Then
    Dim vArgs As Variant: vArgs = args
    Set getSentry = stdSentry.Create( _
        stdCallback.CreateFromObjectMethod(stdError, "addStack").Bind(sSubName, vArgs), _
        stdCallback.CreateFromObjectMethod(stdError, "popStack") _
    ).Run()
  End If
End Function

Public Property Get stack() As stdArray
    'Return the stack object, creating it if it doesn't exist.
    If this.stack Is Nothing Then
        Set this.stack = stdArray.Create()
    End If
    Set stack = this.stack
End Property

Public Sub RefreshStack()
    Set this.stack  = stdArray.Create()
End Sub

'TODO: It'd be cool if we could look backwards at the stack in order to get the calling object name, instead of providing it in sMethodName.
Public Sub AddStack(ByVal sMethodName As String, Optional ByRef args As Variant = Empty)
    Dim sArgs As String: sArgs = ""
    if Not isEmpty(args) then
        Dim iArgs As Long: iArgs = UBound(args) - LBound(args) + 1
        If iArgs > 0 Then sArgs = SerializeArgs(args)
    end if
    If this.stack Is Nothing Then Set this.stack = stdArray.Create()
    Call this.stack.Push(sMethodName & "(" & sArgs & ")")
End Sub

Public Sub PopStack()
    If this.stack Is Nothing Then
        Call Raise( "stdError::StackPop() - No stack created", vbCritical)
    ElseIf this.stack.Length = 0 Then
        Call Raise( "stdError::StackPop() - Stack empty", vbCritical)
    Else
        this.stack.Pop
    End If
End Sub


Public Function Raise(Optional ByVal sMessage As String = "", Optional ByVal Criticality As VBA.VbMsgBoxStyle = VBA.VbMsgBoxStyle.vbExclamation, Optional ByVal Title As String = "VBA Error", Optional ByVal isCritical As Boolean = True) As VBA.VbMsgBoxResult
    'Build stack trace if available
    Dim sStackTrace As String
    sStackTrace = GetTrace()
    
    '
    If this.RaiseClient Is Nothing Then
        'Start full message
        Dim sFullMessage As String
        sFullMessage = "Error in routine """ & sMessage & """" & vbCrLf & sStackTrace
        
        'Return and raise
        Raise = MsgBox(sFullMessage, Criticality, Title)
    Else
        On Error GoTo ErrorOccurred:
            Raise = this.RaiseClient.Raise(sMessage, Criticality, Title, sStackTrace)
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
    Call AddStack("stdError::Raise(" & SerializeArgs(Array("sMessage",sMessage,"Criticality",Criticality, "title", Title, "isCritical", isCritical))  & ")")
        Raise = Raise("stdError::Raise() Error in oMsgClient::Raise() " & Err.Description)
    Call PopStack
End Function


'TODO: If we can get the stack trace directly by walking the VBA stack this would be amazing... Instead we currently have to manually add and remove method names to the stack.
Public Function GetTrace() As String
    'Build stack trace if available
    Dim sStackTrace As String
    If Not this.stack Is Nothing Then
        sStackTrace = "Trace:" & vbCrLf
        
        Dim i As Long
        For i = 1 To this.stack.Length
            sStackTrace = sStackTrace & Space((i - 1) * 3) & "|- " & this.stack.item(i) & vbCrLf
        Next
    End If

    GetTrace = sStackTrace
End Function


Sub Test()
    With stdError
        .RefreshStack
        .AddStack "A"
            .AddStack "B"
                .AddStack "C"
                .PopStack
                .AddStack "D"
                .PopStack
                .AddStack "E"
                    Debug.Print .Raise("Some Error occurred")
                
                    Set RaiseClient = stdArray
                    Debug.Print .Raise("Some Other Error occurred")
                    .AddStack "F"
                        
                        'Purposely overstacking the buffer
                        .AddStack "Something crazy and cool"
                            .AddStack "g"
                                .AddStack "g"
                                    .AddStack "g"
                                        .AddStack "g"
                                            .AddStack "g"
                                                .AddStack "g"
                                                    .AddStack "g"
                                                        .AddStack "g"
                                                            .AddStack "g"
                                                                .AddStack "g"
                                                                    .AddStack "g"
                                                                        .AddStack "g"
                                                                            .AddStack "g"
                                                                                .AddStack "g"
                                                                                    .AddStack "g"
                                                                                        .AddStack "g"
                                                                                            .AddStack "g"
                                                                                                .AddStack "g"
                                                                                                    .AddStack "g"
                                                                                                        .AddStack "g"
                                                                                                            .AddStack "g"
                                                                                                                .AddStack "g"
                                                                                                                    .AddStack "g"
                                                                                                                        .AddStack "g"
                                                                                                                            .AddStack "g"
                                                                                                                                .AddStack "g"
                                                                                                                                    .AddStack "g"
                                                                                                                                        .AddStack "g"
                                                                                                                                            .AddStack "Something crazy and cool"
                                                                                                                                            .Raise "Poop"
                                                                                                                                            'This stack trace will barely be visible in the default alert box.
                                                                                                                                            'Suggest instead using RaiseClient with edit control.
                        
    End With
End Sub

'Helper function to serialize the arguments for the stack trace.
'@param args - Name-value pairs of arguments to serialize. E.G. Array("a", 1, "b", 2)
'@return - String representation of the arguments. E.G. "a: 1, b: 2"
Private Function SerializeArgs(ByRef args As Variant) As String
    Dim i As Long
    Dim sArgs As String
    For i = LBound(args) To UBound(args) step 2
        If i > LBound(args) Then sArgs = sArgs & ", "
        If i + 1 <= UBound(args) Then
            sArgs = sArgs & args(i) & ": " & Serialize(args(i + 1))
        Else
            sArgs = sArgs & args(i) & ": Unknown" 'If there's no value, just use the name.
        End If
    Next i
    SerializeArgs = sArgs
End Function

'Serialize a value to a string representation.
'@param value - The value to serialize.
'@return - String representation of the value.
Private Function Serialize(ByVal value As Variant) As String
    'Serialize the value to a string representation.
    select case VarType(value)
        case vbString
            Serialize = """" & value & """"
        case vbBoolean
            Serialize = CStr(value)
        case vbInteger, vbLong, vbSingle, vbDouble
            Serialize = CStr(value)
        case vbDate
            Serialize = Format(value, "yyyy-mm-dd hh:nn:ss")
        case vbNull
            Serialize = "Null"
        case vbEmpty
            Serialize = "Empty"
        case vbObject
            'Serialize as #<Foo:0x00007ff9c18f7d10>
            If value Is Nothing Then
                Serialize = "Nothing"
            ElseIf TypeName(value) = "stdLambda" Then
                Serialize = "#<stdLambda" & ":" & Hex(ObjPtr(value)) & " Equation='" &   value.Equation & "'>"
            Else
                Serialize = "#<" & TypeName(value) & ":" & Hex(ObjPtr(value)) & ">"
            End If
        Case vbError
            Serialize = "#Error!"
        case Else
            Serialize = "UnknownType"
    end select
End Function