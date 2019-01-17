VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "STD_Runtimes_CLR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Requires references to:
'   C:\Windows\Microsoft.NET\Framework\v4.0.30319\mscoree.tlb    "Common language runtime execution engine"
'and
'   C:\Windows\Microsoft.NET\Framework\v4.0.30319\mscorlib.tlb

Const CLR$ = "v4.0.30319"
Public AppDomain As mscorlib.AppDomain
Public RuntimeHost As mscoree.CorRuntimeHost

Private Declare PtrSafe Function CorBindToRuntimeEx Lib "mscoree" ( _
    ByVal pwszVersion As LongPtr, _
    ByVal pwszBuildFlavor As LongPtr, _
    ByVal startupFlags As Long, _
    ByRef rclsid As Long, _
    ByRef riid As Long, _
    ByRef ppvObject As mscoree.CorRuntimeHost) As Long ' ppvObject As mscoree.CorRuntimeHost

Private Declare PtrSafe Function VariantCopy Lib "oleaut32" (dest, src) As Long

''
' Creates a .Net object with the CLR 4 without registration.  '
''
Public Function CreateInstance(assembly As String, typeName As String) As Variant
  VariantCopy CreateInstance, Me.domain.CreateInstanceFrom(assembly, typeName).Unwrap
End Function

Private Sub LoadLibrary(AssemblyName As String, Optional AppDomain = 0)
  If Not Me.domain Then
  
End Sub


Sub test()
  Dim o As Object
  o = CreateInstance("C:\Windows\Microsoft.NET\Framework64\v4.0.30319\mscorlib.dll", "")
End Sub


Private Sub Class_Initialize()
  Set Me.RuntimeHost = GetHost()
  Set Me.AppDomain = GetDefaultDomain(Me.RuntimeHost)
End Sub

Private Function GetHost() As mscoree.CorRuntimeHost
    Dim hr&, t&(0 To 7)    'host As mscoree.CorRuntimeHost
    t(0) = &HCB2F6723: t(1) = &H11D2AB3A: t(2) = &HC000409C: t(3) = &H3E0AA34F
    t(4) = &HCB2F6722: t(5) = &H11D2AB3A: t(6) = &HC000409C: t(7) = &H3E0AA34F

    hr = CorBindToRuntimeEx(StrPtr(CLR), 0, 3, t(0), t(4), GetHost)
    If hr And -2 Then Err.Raise hr
End Function

Private Function GetDefaultDomain(host As mscoree.CorRuntimeHost) As mscorlib.AppDomain
  Dim domain As mscorelib.AppDomain
  If domain Is Nothing Then
    host.Start
    host.GetDefaultDomain domain
  End If
  Set GetDefaultDomain = domain
End Function

Public Function StartDomain(Optional BaseDirectory = "") As mscorlib.AppDomain
  Set StartDomain = Me.AppDomain.GetType().InvokeMember_3("CreateDomain", &H158, vbNull, vbNull, Array("", BaseDirectory, , Nothing))
End Function

Public Function StopDomain(ByRef AppDomain As mscorlib.AppDomain)
  'ICorRuntimeHost::UnloadDomain
  
End Function

Private Function NETLoadLibrary(AssemblyName As String, Optional ad As mscorlib.AppDomain = Nothing)
  If ad = Nothing Then
    Set ad = Me.AppDomain
  End If
  Dim asm As mscorlib.assembly
  Set asm = ad.Load_2(AssemblyName)
  If asm = Nothing Then
    Dim typeOfAssembly As mscorlib.Type, args As Variant
    Set typeOfAssembly = Me.AppDomain.GetType().assembly.GetType()
    args = Array(AssemblyName)
    
    Set asm = typeOfAssembly.InvokeMember_3("LoadWithPartialName", &H158, vbNull, vbNull, args)
    If asm = Nothing Then
      Set asm = typeOfAssembly.InvokeMember_3("LoadFrom", &H158, vbNull, vbNull, args)
    End If
  End If
  
  Set NETLoadLibrary = asm
End Function

Public Function CompileAssembly(Code As String, References() As String, ProviderAssembly As String, ProviderType As String, FileName As String, CompilerOptions) As mscorlib.assembly
  Dim asmProvider As mscorlib.assembly
  Set asmProvider = NETLoadLibrary(ProviderAssembly, Me.AppDomain)
  If asmProvider Is Nothing Then Exit Function
  
  Dim codeProvider As Variant
  Set codeProvider = asmProvider.CreateInstance(ProviderType)
  If codeProvider Is Nothing Then Exit Function
  
  Dim codeCompiler As Variant
  Set codeCompiler = codeProvider.CreateCompiler()
  If codeCompiler Is Nothing Then Exit Function
  
  Dim asmSystem As assembly
  asmSystem = IIf(ProviderAssembly = "System", asmProvider, NETLoadLibrary("System", Me.AppDomain))
  If asmSystem Is Nothing Then Exit Function
  
  'Set up parameters for compiler
  Dim prms As Variant
  Set prms = CLRCreateObject(asmSystem, "System.CodeDom.Compiler.CompilerParameters", References)
  prms.OutputAssembly = FileName
  prms.GenerateInMemory = FileName = ""
  prms.GenerateExecutable = Right(FileName, 3) = "exe"
  prms.CompilerOptions = CompilerOptions
  prms.IncludeDebugInformation = True
  
  'Compile!
  Dim compilerRes As Object
  compilerRes = codeCompiler.CompileAssemblyFromSource(prms, Code)
  
  Dim error_count As Long
  error_count = compilerRes.errors.Count
  
  If error_count > 0 Then
    Dim error As Variant
    For Each error In compilerRes.errors
      Debug.Assert False
    Next
  End If
  
  If FileName = "" Then
    Set CompileAssembly = compilerRes("CompiledAssembly")
  Else
    Set CompileAssembly = Nothing
  End If
End Function

Public Function CompileCS(Code As String, References() As String, FileName As String, CompilerOptions As String) As assembly
  Set CompileCS = CompileAssembly(Code, References, "System", "Microsoft.CSharp.CSharpCodeProvider", FileName, CompilerOptions)
End Function

Public Function CompileVB(Code As String, References() As String, FileName As String, CompilerOptions As String) As assembly
  Set CompileVB = CompileAssembly(Code, References, "System", "Microsoft.VisualBasic.VBCodeProvider", FileName, CompilerOptions)
End Function

