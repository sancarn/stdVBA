VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type TStack
  initialised as Boolean
  index as Long
  arr() as string
End Type
Private Type TThis
    stack as TStack
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
  Dim v: v = args
  Call stdError.addStack(sSubName, v)

  'A fresh instance acts as a stack popper.
  set getSentry = New stdError
End Function

Public Sub Class_Terminate()
    'Only do this on fresh instances, not the singleton.
    if Me is stdError then Exit Sub
    
    'On class termination, pop the stack.
    On Error Resume Next
    Call stdError.PopStack
End Sub

Public Property Set RaiseClient(ByVal v As Object)
  Set This.RaiseClient = v
End Property

Public Property Get stack(ByRef index as Long) As string()
    With this.stack
      stack = .arr
      index = .index
    end with
End Property

Public Sub RefreshStack()
    With this.stack
      Redim .arr(0 To 0)
      .index = 0
    end with
End Sub

'TODO: It'd be cool if we could look backwards at the stack in order to get the calling object name, instead of providing it in sMethodName.
Public Sub AddStack(ByVal sMethodName As String, Optional ByRef args As Variant = Empty)
    Dim sArgs As String: sArgs = ""
    if Not isEmpty(args) then
        Dim iArgs As Long: iArgs = UBound(args) - LBound(args) + 1
        If iArgs > 0 Then sArgs = SerializeArgs(args)
    end if
    if not this.stack.initialised then this.stack = TStack_Create()
    Call TStack_Push(this.stack, sMethodName & "(" & sArgs & ")")
End Sub

Public Sub PopStack()
    If not this.stack.initialised Then
        Call Raise( "stdError::StackPop() - No stack created", vbCritical)
    ElseIf this.stack.index = 0 Then
        Call Raise( "stdError::StackPop() - Stack empty", vbCritical)
    Else
        Call TStack_PopNoReturn(This.stack)
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
    If this.stack.initialised Then 
        sStackTrace = "Trace:" & vbCrLf
        
        Dim i As Long
        For i = 1 To this.stack.index - 1
            sStackTrace = sStackTrace & Space((i - 1) * 3) & "|- " & this.stack.arr(i) & vbCrLf
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
                
                    Set RaiseClient = tbc
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
    Select Case VarType(value)
        Case vbString
            Serialize = """" & value & """"
        Case vbBoolean
            Serialize = CStr(value)
        Case vbInteger, vbLong, vbSingle, vbDouble
            Serialize = CStr(value)
        Case vbDate
            Serialize = Format(value, "yyyy-mm-dd hh:nn:ss")
        Case vbNull
            Serialize = "Null"
        Case vbEmpty
            Serialize = "Empty"
        Case vbObject
            'Serialize as #<Foo:0x00007ff9c18f7d10>
            If value Is Nothing Then
                Serialize = "Nothing"
            ElseIf TypeName(value) = "stdLambda" Then
                Dim success As Boolean
                Serialize = "#<stdLambda" & ":" & Hex(ObjPtr(value)) & " Equation='" & value.SendMessage("equation", success, Array()) & "'>"
            Else
                Serialize = "#<" & TypeName(value) & ":" & Hex(ObjPtr(value)) & ">"
            End If
        Case vbError
            Serialize = "#Error!"
        Case Else
            Serialize = "UnknownType"
    End Select
End Function

'Create a new stack structure
'@returns - A new TStack structure
Private Function TStack_Create() As TStack
  With TStack_Create
    .initialised = True
    .index = 0
    ReDim .arr(0 To 10)
  End With
End Function

'Given a `stack` structure, append `item` to the stack.
'@param stack - stack of elements
'@param item - Item to insert into stack
Private Sub TStack_Push(ByRef stack as TStack,ByRef item as string)
  With stack
    if not .initialised then
      Err.Raise 1001, "stdError", "TStack_Push() - Stack not initialised"
      Exit Sub
    End if

    Dim ub As Long: ub = UBound(.arr)
    Dim size As Long: size = ub + 1
    If .index > ub Then
        ReDim Preserve .arr(0 To size * 2)
    End If
  
    .arr(.index) = item
    .index = .index + 1
  End With
End Sub


'Given a `stack` structure, pop `item` from the stack.  
'@param stack - stack of elements
'@returns - The popped item
Private Sub TStack_PopNoReturn(ByRef stack As TStack)
  With stack
    if not .initialised then
      Err.Raise 1001, "stdError", "TStack_Pop() - Stack not initialised"
      Exit Sub
    End if

    if .index < 0 then
      Err.Raise 1002, "stdError", "TStack_Pop() - Stack empty"
      Exit Sub
    End if

    'If the stack is too large, shrink it to half its size.
    Dim size As Long: size = UBound(.arr) + 1
    If .index < size / 3 Then
        ReDim Preserve .arr(0 To CLng(size / 2))
    End If

    
    .index = .index - 1
    'TStack_Pop = .arr(.index)
  End With
End Sub