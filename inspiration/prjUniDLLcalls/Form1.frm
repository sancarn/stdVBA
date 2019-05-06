VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1365
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   1365
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   690
      Left            =   435
      TabIndex        =   0
      Top             =   285
      Width           =   3510
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Sub Command1_Click()
    Dim c As cUniversalDLLCalls
    Dim sBuffer As String, lLen As Long
    Set c = New cUniversalDLLCalls
    
'/// 1st four examples show 2 ways of calling an ANSI function & 2 ways of calling a Unicode function
    ' example of calling ANSI function, passing strings ByRef
    Debug.Print "ANSI string parameters, ByRef..."
    lLen = c.CallFunction_DLL("user32.dll", "GetWindowTextLengthA", STR_NONE, CR_LONG, CC_STDCALL, Me.hWnd)
    sBuffer = String$(lLen, vbNullChar)
    ' STR_ANSI + string variable name = ByRef
    lLen = c.CallFunction_DLL("user32.dll", "GetWindowTextA", STR_ANSI, CR_LONG, CC_STDCALL, Me.hWnd, sBuffer, lLen + 1&)
    Debug.Print vbTab; "form caption is: "; Left$(StrConv(sBuffer, vbUnicode), lLen); "<<<"
    
    ' example of calling ANSI function, passing strings ByVal
    Debug.Print "ANSI string parameters, ByVal..."
    lLen = c.CallFunction_DLL("user32.dll", "GetWindowTextLengthA", STR_NONE, CR_LONG, CC_STDCALL, Me.hWnd)
    sBuffer = String$(lLen, vbNullChar)
    ' STR_NONE + string variable name = ByVal. Note: Only use ANSI ByRef if string sole purpose is a buffer
    lLen = c.CallFunction_DLL("user32.dll", "GetWindowTextA", STR_NONE, CR_LONG, CC_STDCALL, Me.hWnd, StrPtr(sBuffer), lLen + 1&)
    Debug.Print vbTab; "form caption is: "; Left$(StrConv(sBuffer, vbUnicode), lLen); "<<<"
    
    ' example of calling UNICODE function, passing strings ByRef
    Debug.Print "Unicode string parameters, ByRef..."
    lLen = c.CallFunction_DLL("user32.dll", "GetWindowTextLengthW", STR_NONE, CR_LONG, CC_STDCALL, Me.hWnd)
    sBuffer = String$(lLen, vbNullChar)
    ' STR_UNICODE + string variable name = ByRef
    lLen = c.CallFunction_DLL("user32.dll", "GetWindowTextW", STR_UNICODE, CR_LONG, CC_STDCALL, Me.hWnd, sBuffer, lLen + 1&)
    Debug.Print vbTab; "form caption is: "; Left$(sBuffer, lLen); "<<<"
    
    ' example of calling UNICODE function, passing strings ByVal
    Debug.Print "Unicode string parameters, ByVal..."
    lLen = c.CallFunction_DLL("user32.dll", "GetWindowTextLengthW", STR_NONE, CR_LONG, CC_STDCALL, Me.hWnd)
    sBuffer = String$(lLen, vbNullChar)
    ' STR_NONE + StrPtr(variable name) = ByVal
    lLen = c.CallFunction_DLL("user32.dll", "GetWindowTextW", STR_NONE, CR_LONG, CC_STDCALL, Me.hWnd, StrPtr(sBuffer), lLen + 1&)
    Debug.Print vbTab; "form caption is: "; Left$(sBuffer, lLen); "<<<"
    
'/// UDT/Array examples
    ' example of passing a structure
    Dim tRect As RECT
    Debug.Print "UDT/structure parameters, ByRef..."
    Call c.CallFunction_DLL("user32.dll", "GetWindowRect", STR_NONE, CR_LONG, CC_STDCALL, Me.hWnd, VarPtr(tRect))
    Debug.Print vbTab; "window position on screen: L"; CStr(tRect.Left); ".T"; CStr(tRect.Top); "   R"; CStr(tRect.Right); ".B"; CStr(tRect.Bottom)
    
    ' the RECT structure is 16 bytes, we can use an array of Long if we like
    Dim aRect(0 To 3) As Long
    Debug.Print "Array parameters, ByRef..."
    Call c.CallFunction_DLL("user32.dll", "GetWindowRect", STR_NONE, CR_LONG, CC_STDCALL, Me.hWnd, VarPtr(aRect(0)))
    Debug.Print vbTab; "window position on screen: L"; CStr(aRect(0)); ".T"; CStr(aRect(1)); "   R"; CStr(aRect(2)); ".B"; CStr(aRect(3))
    
    
'/// CDecl function call
    Dim sFmt As String
    sBuffer = String$(1024, vbNullChar)
    sFmt = "P1=%s, P2=%d, P3=%.4f, P4=%s"
    ' unicode version of the function
    Debug.Print "CDecl Unicode parameters, ByRef..."
    lLen = c.CallFunction_DLL("msvcrt.dll", "swprintf", STR_UNICODE, CR_LONG, CC_CDECL, sBuffer, sFmt, "ABC", 123456, 1.23456, "xyz")
    Debug.Print vbTab; "printf: "; Left$(sBuffer, lLen)
    ' ANSI version of the function, same parameters
    Debug.Print "CDecl ANSI parameters, ByRef..."
    lLen = c.CallFunction_DLL("msvcrt.dll", "sprintf", STR_ANSI, CR_LONG, CC_CDECL, sBuffer, (sFmt), "ABC", 123456, 1.23456, "xyz")
    Debug.Print vbTab; "printf: "; Left$(StrConv(sBuffer, vbUnicode), lLen)
    
''/// COM object call
    ' All VB objects inherit from IUnknown (which has 3 virtual functions)
    ' IPicture inherits from IUnknown and has several virtual functions
    ' This example will call the 1st function which is now the 4th function, preceeded by IUnknown's 3 functions
    
    ' NOTE: simple example. We can declare a IPicture interface via VB, but many interfaces are not exposed,
    ' and this example indicates how to get a pointer to the interface & call functions from that pointer.
    ' But just like any function, you must research to determine the VTable order & function parameter
    ' requirements. Do not assume that some page describing the interface functions lists the functions
    ' in VTable order. That assumption will lead to crashes.
    Dim IID_IPicture As Long, aGUID(0 To 3) As Long, lPicHandle As Long

    Const IUnknownQueryInterface As Long = 0&   ' IUnknown vTable offset to Query implemented interfaces
    Const IUnknownRelease As Long = 8&          ' IUnkownn vTable offset to decrement reference count
    Const IPictureGetHandle As Long = 12&       ' 4th VTable offset from IUnknown
    ' GUID for IPicture {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    c.CallFunction_DLL "ole32.dll", "CLSIDFromString", STR_UNICODE, CR_LONG, CC_STDCALL, "{7BF80980-BF32-101A-8BBB-00AA00300CAB}", VarPtr(aGUID(0))
    c.CallFunction_COM ObjPtr(Me.Icon), IUnknownQueryInterface, CR_LONG, CC_STDCALL, VarPtr(aGUID(0)), VarPtr(IID_IPicture)
    If IID_IPicture <> 0& Then
        ' get the icon handle & then Release the IPicture interface. QueryInterface calls AddRef internally
        c.CallFunction_COM IID_IPicture, 12&, CR_LONG, CC_STDCALL, VarPtr(lPicHandle)
        c.CallFunction_COM IID_IPicture, IUnknownRelease, CR_LONG, CC_STDCALL
    End If
    Debug.Print "COM interface call example..."
    Debug.Print vbTab; "Me.Icon.Handle = "; Me.Icon.Handle; " IPicture.GetHandle = "; lPicHandle

' The PointerToString methods are a courtesy
'/// simple example to return a string from a pointer
    sFmt = "LaVolpe"
    Debug.Print "PointerToStringA & PointerToStringW examples..."
    sBuffer = c.PointerToStringW(StrPtr(sFmt))  ' unicode example
    Debug.Print vbTab; sBuffer; "<<<"
    sFmt = StrConv(sFmt, vbFromUnicode)
    sBuffer = c.PointerToStringA(StrPtr(sFmt))  ' ANSI example
    Debug.Print vbTab; sBuffer; "<<<"
    
End Sub

Private Sub Form_Load()
    Me.Caption = "Universal DLL Calls"
    Command1.Caption = "Click for Simple Examples" & vbCrLf & "Printed to Immediate Window"
End Sub
