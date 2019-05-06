VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   210
      TabIndex        =   1
      Top             =   360
      Width           =   4065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Download"
      Height          =   495
      Left            =   1110
      TabIndex        =   0
      Top             =   1110
      Width           =   1860
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CLSID_ProgressDialog As String = "{F8383852-FCD3-11d1-A6B9-006097DF5BD4}"
Private Const IID_IProgressDialog As String = "{EBBC7C04-315E-11D2-B62F-006097DF5BD4}"
Private Declare Function IIDFromString Lib "ole32.dll" (ByVal lpsz As Long, ByVal lpiid As Long) As Long
Private Declare Function CoCreateInstance Lib "ole32.dll" (ByVal rclsid As Long, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, ByVal rIID As Long, ByRef ppv As Long) As Long
Private Const CLSCTX_INPROC_SERVER As Long = 1

Implements IBindStatusCallback ' http://msdn.microsoft.com/en-us/library/ie/ms775060%28v=vs.85%29.aspx
Private m_IBinder As IBindStatusCallback
Private m_URL As String
Private m_IProgress As Object ' http://msdn.microsoft.com/en-us/library/windows/desktop/bb775248%28v=vs.85%29.aspx

Private Declare Function URLDownloadToCacheFileW Lib "URLMON.dll" (ByVal lpunk As Long, ByVal lpcstr As Long, ByVal lpFile As Long, ByVal dwNameLen As Long, ByVal dwReserved As Long, ByVal pIB As Long) As Long


Private Sub Command1_Click()

    If Text1.Text = vbNullString Then
        MsgBox "Paste or type a url into the text box", vbExclamation + vbOKOnly
        Exit Sub
    End If

    Dim sFile As String, lResult As Long
    
    sFile = String$(255, vbNullChar)
    Set m_IBinder = New IBindStatusCallback
    Set m_IBinder.Owner = Me
    m_IBinder.InitThunks
    m_URL = Text1.Text
    lResult = URLDownloadToCacheFileW(0&, StrPtr(m_URL), _
                            StrPtr(sFile), 255, 0&, m_IBinder.GetInterfacePointer())
    If lResult = 0 Then
        If MsgBox("Done Downloading" & vbNewLine & "Kill downloaded file?", vbYesNo + vbDefaultButton2, "Confirmation") = vbYes Then
            On Error Resume Next
            Kill Left$(sFile, InStr(sFile, vbNullChar) - 1)
        End If
    Else
        MsgBox "Downloading Failed" & vbNewLine & "Return Result: " & lResult, vbInformation + vbOKOnly
    End If
End Sub

Private Sub IBindStatusCallback_OnDownloadEnd(Result As Long)
    If ObjPtr(m_IProgress) Then
        ' 5th method: HRESULT StopProgressDialog();
        m_IBinder.pvCallFunction_COM ObjPtr(m_IProgress), 16&
        Set m_IProgress = Nothing
    End If
End Sub

Private Sub IBindStatusCallback_OnDownloadStart()

    Dim aGUID(0 To 7) As Long, lPtr As Long
    IIDFromString StrPtr(CLSID_ProgressDialog), VarPtr(aGUID(0))
    IIDFromString StrPtr(IID_IProgressDialog), VarPtr(aGUID(4))
    
    If CoCreateInstance(VarPtr(aGUID(0)), 0&, CLSCTX_INPROC_SERVER, VarPtr(aGUID(4)), lPtr) = 0& Then
        If lPtr = 0& Then Exit Sub
        Set m_IProgress = m_IBinder.pvPointerToObject(lPtr)
        ' 4th method: StartProgressDialog([in]  HWND hwndParent,IUnknown *punkEnableModless,DWORD dwFlags,LPCVOID pvReserved
        m_IBinder.pvCallFunction_COM lPtr, 12&, Me.hWnd, 0&, 0&, 0&
        ' 6th method: HRESULT SetTitle([in]  PCWSTR pwzTitle);
        m_IBinder.pvCallFunction_COM lPtr, 20&, StrPtr("Welcome to VTable Interfaces")
        ' 11th method: HRESULT SetLine(DWORD dwLineNum,[in]  PCWSTR pwzString,BOOL fCompactPath,LPCVOID pvReserved)
        m_IBinder.pvCallFunction_COM lPtr, 40&, 2&, StrPtr("Waiting..."), 0&, 0&
        m_IBinder.pvCallFunction_COM lPtr, 40&, 1&, StrPtr(m_URL), 1&, 0&
    Else
        Debug.Print "failed to create progress dialog"
    End If

End Sub

Private Sub IBindStatusCallback_OnProgress(Current As Long, Max As Long, Code As Long, Abort As Boolean)
    If ObjPtr(m_IProgress) = 0& Then Exit Sub
    If Code = 5& Then       ' init message for data receipt
        ' 11th method: HRESULT SetLine(DWORD dwLineNum,[in]  PCWSTR pwzString,BOOL fCompactPath,LPCVOID pvReserved)
        m_IBinder.pvCallFunction_COM ObjPtr(m_IProgress), 40&, 2&, StrPtr(""), 0&, 0&
        If Max < 1& Then        ' size of download unknown
            ' 4th method: StartProgressDialog([in]  HWND hwndParent,IUnknown *punkEnableModless,DWORD dwFlags,LPCVOID pvReserved
            m_IBinder.pvCallFunction_COM ObjPtr(m_IProgress), 12&, Me.hWnd, 0&, &H20&, 0& ' show as marquee
        End If
    ElseIf Code = 6 Then    ' continuing messages for data receipt
        ' 8th method: BOOL HasUserCancelled()
        Abort = (m_IBinder.pvCallFunction_COM(ObjPtr(m_IProgress), 28&) <> 0)
        ' 9th method: HRESULT SetProgress([in]  DWORD dwCompleted,[in]  DWORD dwTotal;
        If Not Abort Then m_IBinder.pvCallFunction_COM ObjPtr(m_IProgress), 32&, Current, Max
    End If
End Sub

Private Sub Text1_GotFocus()
    With Text1
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub
