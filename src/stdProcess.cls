


'Full example shown below, including the necessary structures
#If VBA7 Then
    Private Declare PtrSafe Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, _ 
            lpProcessAttributes As tSecurityAttributes, lpThreadAttributes As tSecurityAttributes, ByVal bInheritHandles As Long, ByVal dwCreationFlags As ECreationFlags, _
            ByRef lpEnvironment As Any, ByVal lpCurrentDriectory As String, ByRef lptStartupInfo As tStartupInfo, ByRef lpProcessInformation As tProcessInformation) As LongPtr

    Private Type tProcessInformation
        hProcess As LongPtr
        hThread As LongPtr
        dwProcessId As Long
        dwThreadId As Long
    End Type

    Private Type tStartupInfo
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Byte
        hStdInput As LongPtr
        hStdOutput As LongPtr
        hStdError As LongPtr
    End Type

    Private Type tSecurityAttributes
        nLength As Long
        lpSecurityDescriptor As LongPtr
        bInheritHandle As Long
    End Type


#Else
    Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, _ 
            lpProcessAttributes As tSecurityAttributes, lpThreadAttributes As tSecurityAttributes,ByVal bInheritHandles As Long, ByVal dwCreationFlags As ECreationFlags, _ 
            ByRef lpEnvironment As Any, ByVal lpCurrentDriectory As String, ByRef lptStartupInfo As tStartupInfo, ByRef lpProcessInformation As tProcessInformation) As Long

    Private Type tProcessInformation
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
    End Type

    Private Type tStartupInfo
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Byte
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
    End Type

    Private Type tSecurityAttributes
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
    End Type 
#End If

Const INFINITE = &HFFFF
Const STARTF_USESHOWWINDOW = &H1
Private Enum ECreationFlags
    CREATE_BREAKAWAY_FROM_JOB        = &H01000000
    CREATE_DEFAULT_ERROR_MODE        = &H04000000
    CREATE_NEW_CONSOLE               = &H00000010
    CREATE_NEW_PROCESS_GROUP         = &H00000200
    CREATE_NO_WINDOW                 = &H08000000
    CREATE_PROTECTED_PROCESS         = &H00040000
    CREATE_PRESERVE_CODE_AUTHZ_LEVEL = &H02000000
    CREATE_SECURE_PROCESS            = &H00400000
    CREATE_SEPARATE_WOW_VDM          = &H00000800
    CREATE_SHARED_WOW_VDM            = &H00001000
    CREATE_SUSPENDED                 = &H00000004
    CREATE_UNICODE_ENVIRONMENT       = &H00000400
    DEBUG_ONLY_THIS_PROCESS          = &H00000002
    DEBUG_PROCESS                    = &H00000001
    DETACHED_PROCESS                 = &H00000008
    EXTENDED_tStartupInfo_PRESENT     = &H00080000
    INHERIT_PARENT_AFFINITY          = &H00010000
    PRIORITY_CLASS_NORMAL            = &H00000020
    PRIORITY_CLASS_IDLE              = &H00000040
    PRIORITY_CLASS_HIGH              = &H00000080
End Enum
Private Enum EStartupFlags
    STARTF_FORCEONFEEDBACK    = &H00000040
    STARTF_FORCEOFFFEEDBACK   = &H00000080
    STARTF_PREVENTPINNING     = &H00002000
    STARTF_RUNFULLSCREEN      = &H00000020
    STARTF_TITLEISAPPID       = &H00001000
    STARTF_TITLEISLINKNAME    = &H00000800
    STARTF_UNTRUSTEDSOURCE    = &H00008000
    STARTF_USECOUNTCHARS      = &H00000008
    STARTF_USEFILLATTRIBUTE   = &H00000010
    STARTF_USEHOTKEY          = &H00000200
    STARTF_USEPOSITION        = &H00000004
    STARTF_USESHOWWINDOW      = &H00000001
    STARTF_USESIZE            = &H00000002
    STARTF_USESTDHANDLES      = &H00000100
End Enum

Private pProcessId as long
Private pThreadId as long
#if VBA7 then
    Private pProcessHandle as LongPtr
    Private pThreadHandle as LongPtr
#else
    Private pProcessHandle as Long
    Private pThreadHandle as Long
#end if

Private pStartup as tStartupInfo


Public Function Create(ByVal sCmd as string, Optional ByVal sWorkDir as string = "", Optional ByVal winStyle As VbAppWinStyle = VbAppWinStyle.vbHide, Optional ByVal dwCreationFlags as ECreationFlags = ECreationFlags.PRIORITY_CLASS_NORMAL, Optional ByVal dwStartupFlags as EStartupFlags = EStartupFlags.STARTF_USESHOWWINDOW) as stdProcess
    set Create = new stdProcess
    Call Create.InitFromLaunch(sCmd,sWorkDir,winStyle,dwCreationFlags,dwStartupFlags)
End Function
Public Sub InitFromLaunch(ByVal sCmd as string, Optional ByVal lpCurrentDirectory as string = vbNullString, Optional ByVal winStyle As VbAppWinStyle = VbAppWinStyle.vbHide, Optional ByVal dwCreationFlags as ECreationFlags = ECreationFlags.PRIORITY_CLASS_NORMAL, Optional ByVal dwStartupFlags as EStartupFlags = EStartupFlags.STARTF_USESHOWWINDOW)
    Dim lpProcessAttributes As tSecurityAttributes: lpProcessAttributes.nLength = Len(lpProcessAttributes)
    Dim lpThreadAttributes As tSecurityAttributes: lpThreadAttributes.nLength = Len(lpThreadAttributes)
    Dim lpStartupInfo As tStartupInfo: lpStartupInfo = createStartupInfo(winStyle,dwStartupFlags)
    Dim lpProcessInformation as tProcessInformation
    if CreateProcess(vbNullString, sCmd, lpProcessAttributes, lpThreadAttributes, False, dwCreationFlags, 0&, lpCurrentDirectory, lpStartupInfo, lpProcessInformation) then
        pProcessId = lpProcessInformation.dwProcessId
        pThreadId = lpProcessInformation.dwThreadId
        pProcessHandle = lpProcessInformation.hProcess
        pThreadHandle = lpProcessInformation.hThread
        pStartup = lpStartupInfo
    else
        '
    end if
End Sub

Public Function CreateFromProcessId(ByVal pID as long) as stdProcess
    '
End Function
Public Sub InitFromProcessId(ByVal pID as long)
    pProcessId = pID
    Dim wmiObject as object
	Set wmiObject = GetObject("WINMGMTS:\\.\ROOT\cimv2:Win32_Process.Handle=" & id)

End Sub


Public Function CreateFromExisting(ByVal query as stdICallable) as stdProcess
    '
End Function

'@returns {Collection<stdProcess>}
Public Function CreateFromAllExisting(ByVal query as stdICallable) as Collection
    '
End Function

Public Property Get id() as Long
    id = pProcessId
End Property

'@returns {Object<SWbemObjectEx>}
Public Property Winmgmt() as Object
    set Winmgmt = GetObject("WINMGMTS:\\.\ROOT\cimv2:Win32_Process.Handle=" & pProcessId)
End Property




'
'@param
Private Function createStartupInfo(ByVal winStyle As VbAppWinStyle, ByVal dwCreationFlags as EStartupFlags) as tStartupInfo
    Dim tSI as tStartupInfo
    sinfo.dwFlags = dwCreationFlags
    sinfo.wShowWindow = winStyle
    tSI.cb = lenb(tSI)
    createStartupInfo = tSI
End Function