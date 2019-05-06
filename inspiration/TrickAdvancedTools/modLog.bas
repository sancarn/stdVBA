Attribute VB_Name = "modLog"
' // modLog.bas - logging
' // © Krivous Anatoly Anatolevich (The trick), 2016

Option Explicit

Public Enum eADDIN_ERRORS
    ERR_UNABLE_TO_HOOK_FUNCTIONS = vbObjectError + 1
    ERR_UNABLE_TO_INITIALIZE_CHECKING
End Enum

' // Show message with error information
Public Sub ErrorLog( _
           ByRef ProcedureName As String)
    
    MsgBox "Error has occured in " & ProcedureName & ": " & Err.Number & " " & Err.Description & _
           vbNewLine & "LastDllError: " & Err.LastDllError & " " & SystemErrorToString(Err.LastDllError)

End Sub

' //
Private Function SystemErrorToString( _
                 ByVal ErrorNumber As Long) As String
    Dim lMsg As Long
    
    SystemErrorToString = Space$(32767)
    
    lMsg = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
                         ByVal 0&, _
                         ErrorNumber, _
                         0, _
                         SystemErrorToString, _
                         Len(SystemErrorToString), _
                         ByVal 0&)

    SystemErrorToString = Left$(SystemErrorToString, lMsg)
    
End Function


