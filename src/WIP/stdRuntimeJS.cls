VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdRuntimeJS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FURTHER READING
'------------------------------
'https://www.autohotkey.com/boards/viewtopic.php?t=5739
'https://www.autohotkey.com/boards/viewtopic.php?f=6&t=4555
'https://github.com/microsoft/ChakraCore/wiki/JavaScript-Runtime-%28JSRT%29-Overview




'JsProjectWinRTNamespace    - Projects .NET namespace to javascript. E.G. JsProjectWinRTNamespace("Winodws.Foundation")
'JsCreateNamedFunction(...) - Bind native function to javascript
'JSCreateObject()           - Createa a new js object
'JsRunScript()
'JsRunSerializedScriptWithCallback()


'Porting table:
'"x*",value ==> "x", VarPtr(value)
'"ptr" ==> LongPtr
'"Float" ==> Float
'"Double" ==> Double
'"Char" ==> Byte
'






'IE:
'  LibHandle =  LoadLibrary("jscript9")
'  jscript9\JsCreateRuntime()
'  jscript9\JSCreateContext()
'  jscript9\JsSetCurrentContext()
'  jscript9\JsGetGlobalObject()
'  init("jscript9",runtime,context)


'Edge:
'  Call LoadLibrary("chakra")
'  chakra\JSCreateRuntime()
'  chakra\JSCreateContext()
'  chakra\JsSetCurrentContext()
'  chakra\JsGetGlobalObject()
'  init("jscript9",runtime,context)

'init::
'  this._dll := dll
'  this._runtime := runtime
'  this._context := context
'  this._dsp := this._JsToVt(globalObject) 'get virtual table. Here it'd be better to use a COM interface wrapper probably COMObject.CreateFromPointer(ptr,{SomeMethod:dispID,...})


'GETTING/SETTING DATA:
'  jscript9\JsValueToVariant()
'  chakra\JsValueToVariant()
'ALTERNATIVE:
'  set Poop = myObject
'  ENGINE\JsGetPropertyIdFromName("poop", hostPropertyId)            ' Get a property ID for the name "host"
'  hostRef := JSVariantToValue(Poop)                                 ' Get a JsValueRef for our Host object
'  ENGINE\JsSetProperty(globalObject, hostPropertyId, hostRef, true) ' Pass our Host object to the script engine
'  inJS: poop.pie()

'DELETE
'  IE:
'    jscript9\JsSetCurrentContext()
'    jscript9\JsDisposeRuntime()
'  Edge:
'    chakra\JsSetCurrentContext()
'    chakra\JsDisposeRuntime()
'    FreeLibrary(LibHandle)

Option Explicit

Private Const DLL_JScript = "jscript9.dll"
Private Const DLL_CHAKRA = "chakra.dll"


'Core
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Any, ByVal wParam As Any, ByVal lParam As Any) As Long
Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long

'FOR IE
'----------------------------------------
'For initialisation:
'Private Declare Function IECreateRuntime Lib "jscript9.dll" Alias "JsCreateRuntime" (jsRuntimeAttributes As Integer, jsRuntimeVersion As Long, JsThreadServiceCallback As LongPtr, jsRuntimeHandle As LongPtr) As Long
Private Declare PtrSafe Function IECreateRuntime Lib "jscript9.dll" Alias "JsCreateRuntime" (ByVal jsRuntimeAttributes As Integer, ByVal jsRuntimeVersion As Long, ByVal JsThreadServiceCallback As LongPtr, ByVal jsRuntimeHandle As LongPtr) As Long
Private Declare PtrSafe Function IECreateContext Lib "jscript9.dll" Alias "JsCreateContext" (ByVal jsRuntimeHandle As LongPtr, ByVal IDebugApplication As LongPtr, ByVal jsContextRef As LongPtr) As Long
Private Declare PtrSafe Function IESetCurrentContext Lib "jscript9.dll" Alias "JsSetCurrentContext" (ByVal jsContextRef As LongPtr) As Long
Private Declare PtrSafe Function IEGetGlobalObject Lib "jscript9.dll" Alias "JsGetGlobalObject" (ByVal jsGlobalValueRef As LongPtr) As Long

'For uninitialisation:
Private Declare PtrSafe Function IEDisposeRuntime Lib "jscript9.dll" Alias "JsDisposeRuntime" (ByVal jsRuntimeHandle As LongPtr) As Long

'Getting/Setting values:
Private Declare PtrSafe Function IEVariantToValue Lib "jscript9.dll" Alias "JsVariantToValue" (ByVal src As LongPtr, ByVal dest As LongPtr) As Long
Private Declare PtrSafe Function IEValueToVariant Lib "jscript9.dll" Alias "JsValueToVariant" (ByVal src As LongPtr, ByVal dest As LongPtr) As Long

'Executing JavaScript
Private Declare PtrSafe Function IERunScript Lib "jscript9.dll" Alias "JsRunScript" (ByVal code As String, ByVal JsSourceContext As Integer, ByVal JsSourceURL As String, ByVal JsValueRef As LongPtr) As Long

'Getting exception information
Private Declare PtrSafe Function IEGetAndClearException Lib "jscript9.dll" Alias "JsGetAndClearException" (ByVal JsValueRef As LongPtr) As Long


'----------------------------------------

'FOR Edge
'----------------------------------------
'For initialisation:
Private Declare PtrSafe Function EdgeCreateRuntime Lib "chakra.dll" Alias "JsCreateRuntime" (ByVal jsRuntimeAttributes As Integer, ByVal JsThreadServiceCallback As LongPtr, ByVal jsRuntimeHandle As LongPtr) As Long 'Note: ThreadServiceCallback seems pointless in VBA
Private Declare PtrSafe Function EdgeCreateContext Lib "chakra.dll" Alias "JsCreateContext" (ByVal jsRuntimeHandle As LongPtr, ByVal jsContextRef As LongPtr) As Long
Private Declare PtrSafe Function EdgeSetCurrentContext Lib "chakra.dll" Alias "JsSetCurrentContext" (ByVal jsContextRef As LongPtr) As Long
Private Declare PtrSafe Function EdgeGetGlobalObject Lib "chakra.dll" Alias "JsGetGlobalObject" (ByVal jsGlobalValueRef As LongPtr) As Long

'For uninitialisation:
Private Declare PtrSafe Function EdgeDisposeRuntime Lib "chakra.dll" Alias "JsDisposeRuntime" (ByVal jsRuntimeHandle As LongPtr) As Long

'Getting/Setting values:
Private Declare PtrSafe Function EdgeVariantToValue Lib "chakra.dll" Alias "JsVariantToValue" (ByVal src As LongPtr, ByVal dest As LongPtr) As Long
Private Declare PtrSafe Function EdgeValueToVariant Lib "chakra.dll" Alias "JsValueToVariant" (ByVal src As LongPtr, ByVal dest As LongPtr) As Long

'Executing JavaScript
Private Declare PtrSafe Function EdgeRunScript Lib "chakra.dll" Alias "JsRunScript" (ByVal code As String, ByVal JsSourceContext As Integer, ByVal JsSourceURL As String, ByVal JsValueRef As LongPtr) As Long

'Getting exception information
Private Declare PtrSafe Function EdgeGetAndClearException Lib "chakra.dll" Alias "JsGetAndClearException" (ByVal JsValueRef As LongPtr) As Long

'Extra function - project WinRT namespaces to JS runtime:
Private Declare PtrSafe Function EdgeProjectWinRTNamespace Lib "chakra.dll" Alias "JsProjectWinRTNamespace" (ByVal sNamespace As String) As Long

'----------------------------------------

'Runtime attributes: (these can be combined through adding, E.G:
'   jsRuntimeAttribute.AllowScriptInterrupt + jsRuntimeAttribute.EnableExperimentalFeatures
'will both allow script interruption and enable experimental js features.
Private Enum jsRuntimeAttribute
  None = &H0
  DisableBackgroundWork = &H1
  AllowScriptInterrupt = &H2
  EnableIdleProcessing = &H4
  DisableNativeCodeGeneration = &H8
  DisableEval = &H10
  EnableExperimentalFeatures = &H20
  DispatchSetExceptionsToDebugger = &H40
  DisableFatalOnOOM = &H80
End Enum

'Not sure what the enumeration here is, but AHK uses -1 for this param so maybe -1 is latest?
Private Enum jsRuntimeVersion
  JsRuntimeVersion10
  JsRuntimeVersion11 = -1
  JsRuntimeVersionEdge 'Deprecated
End Enum


Private Const S_OK = 0



Private initialised As Boolean
Private pVersion As String
Private pHModule As LongPtr
Private pRuntime As Variant
Private pContext As Variant
Public Window As Object      'Global object (even though no actual window available)

Public Function Create(Optional ByVal version As String = "IE") As stdRuntimeJS
  Set Create = New stdRuntimeJS
  Call Create.Init(version)
End Function


Public Sub Init(Optional ByVal version As String = "Latest")
  Dim sError As String
  
  If Not initialised Then
    'Set version property
    pVersion = version
    If pVersion = "Latest" Then
      If initVersion("Edge") = "" Then
        pVersion = "Edge"
      ElseIf initVersion("IE") = "" Then
        pVersion = "IE"
      Else
        sError = "Cannot initialise latest JS engine..."
      End If
    Else
      sError = initVersion(version)
      If sError <> "" Then GoTo Error
    End If
  Else
    sError = "Error not initialised": GoTo Error
  End If
  Exit Sub
  
Error:
  'Call stdError.raise(sError)
  'Debug.Assert False
End Sub


'#TODO: Not entirely sure why but I keep getting the same error on retl (no matter if I go IE or Edge)
'       0x10001 (Which I believe to be JsErrorInvalidArgument). Wondering whether last arg shouldn't be a Pointer?
Private Function initVersion(version As String) As String
  'Versions = ["IE","Edge"]
  Dim hGlobal As LongPtr, sError As String, retl As Long
  Select Case pVersion
    Case "IE"
      pHModule = LoadLibrary(DLL_JScript)
      If pHModule <> 0 Then
        'Create the javascript runtime...
        Dim var
        retl = IECreateRuntime(0, -1, 0, VarPtr(pRuntime))
        If retl = S_OK Then
          'Create an execution context
          Dim ctxt As LongPtr
          retl = IECreateContext(VarPtr(pRuntime), 0, VarPtr(ctxt))
          If retl = S_OK Then
            Debug.Print "Got here!"
            'Set current execution context
            Call IESetCurrentContext(VarPtr(pContext))
            
            'Get global variable as COM object
            Call IEGetGlobalObject(hGlobal)
            Set Window = helperJS2VB(hGlobal)
            
            sError = ""
          Else
            sError = "Cannot create IE runtime context. Error (0x" & Hex(retl) & ")": GoTo Error
          End If
        Else
          sError = "Cannot create IE runtime. Error (0x" & Hex(retl) & ")": GoTo Error
        End If
      Else
        sError = "Couldn't load JScript9 DLL. Error (0x" & Hex(retl) & ")": GoTo Error
      End If
    Case "Edge"
      pHModule = LoadLibrary(DLL_CHAKRA)
      If pHModule <> 0 Then
        'Create the javascript runtime...
        retl = EdgeCreateRuntime(0, 0, VarPtr(pRuntime)) 'KEEP GETTING ERRORS?
        If retl = S_OK Then
          'Create an execution context
          If EdgeCreateContext(VarPtr(pRuntime), VarPtr(pContext)) = S_OK Then
            'Set current execution context
            Call EdgeSetCurrentContext(VarPtr(pContext))
            
            'Get global variable as COM object
            Call EdgeGetGlobalObject(hGlobal)
            Set Window = helperJS2VB(hGlobal)
            
            sError = ""
          Else
            sError = "Cannot create Edge runtime context. Error (0x" & Hex(retl) & ")": GoTo Error
          End If
        Else
          sError = "Cannot create Edge runtime. Error (0x" & Hex(retl) & ")": GoTo Error   '& getJsError(retl)
        End If
      Else
        sError = "Couldn't load Chakra DLL. Maybe Edge isn't installed? Error (0x" & Hex(retl) & ")": GoTo Error
      End If
    Case Else
      sError = "Error incompatible version. Error (0x" & Hex(retl) & ")": GoTo Error
  End Select
  Exit Function
Error:
  Debug.Print sError
  Call Terminate
  initVersion = sError
  'Debug.Assert False
End Function

'Loads the file given
Public Sub Require(sPath As String)
  
End Sub

Public Sub Run(script As String)
  
  Dim v As LongPtr, sError As String
  Select Case pVersion
    Case "IE"
      Call IERunScript(script, 0, "source.js", VarPtr(v))
    Case "Edge"
      Call EdgeRunScript(script, 0, "source.js", VarPtr(v))
    Case Else
      sError = "Unknown version code": GoTo Error
  End Select
  Exit Sub
Error:
  'stdError.raise(sError)
  Debug.Assert False
  
End Sub

Public Function Eval()

End Function

Public Sub AddObject(ByVal sName As String, ByRef obj As Object)
  
End Sub

Public Sub ProjectWinRTNamespace(sNamespace As String)
  Dim sError As String
  Select Case pVersion
    Case "IE"
      sError = "Projecting WinRT namespaces is not possible in IE9": GoTo Error
    Case "Edge"
      Call EdgeProjectWinRTNamespace(sNamespace)
  End Select
  Exit Sub
Error:
  'Call stdError.raise(sError)
End Sub



Private Function helperJS2VB(js As LongPtr) As Variant
  Dim v As Variant, sError As String
  Select Case pVersion
    Case "IE"
      Call IEVariantToValue(js, VarPtr(v))
    Case "Edge"
      Call EdgeVariantToValue(js, VarPtr(v))
    Case Else
      sError = "Unknown version code": GoTo Error
  End Select
  
  'Return data
  If IsObject(v) Then
    Set helperJS2VB = v
  Else
    helperJS2VB = v
  End If
  Exit Function
Error:
  'stdError.raise(sError)
End Function
Private Function helperVB2JS(vb As Variant) As LongPtr
  Dim v As LongPtr, sError As String
  Select Case pVersion
    Case "IE"
      Call IEVariantToValue(VarPtr(vb), v)
    Case "Edge"
      Call EdgeVariantToValue(VarPtr(vb), v)
    Case Else
      sError = "Unknown version code": GoTo Error
  End Select
  
  'Return data
  helperVB2JS = v
  Exit Function
Error:
  'Call stdError.raise(sError)
End Function

Private Sub Class_Terminate()
  Call Terminate
End Sub

Public Sub Terminate()
  If Not IsEmpty(pRuntime) Then
    Select Case pVersion
      Case "IE"
        Call IEDisposeRuntime(VarPtr(pRuntime))
        pRuntime = Empty
      Case "Edge"
        Call EdgeDisposeRuntime(VarPtr(pRuntime))
        pRuntime = Empty
      Case Else
        'Do nothing?
    End Select
  End If
  
  'Free library
  Call FreeLibrary(pHModule)
End Sub

'Can't find many of these error codes but good to keep them here in their true form
'Can find error names here: https://docs.microsoft.com/en-us/microsoft-edge/hosting/chakra-hosting/jserrorcode-enumeration
'and here:  https://github.com/mjrgh/PinballY/blob/master/PinballY/JavascriptEngine.cpp#L1397
'but no values
Private Function getJsError(i As Long) As String
  Select Case i
    Case &H0: getJsError = "JsNoError"
    Case &H10001: getJsError = "JsErrorInvalidArgument"
    Case &H10002: getJsError = "JsErrorNullArgument"
    Case &H30002: getJsError = "JsErrorScriptCompile"
  End Select
End Function







'stdRuntimeJs.test
Public Sub test()
  Debug.Assert False
  Dim edge As stdRuntimeJS
  Set edge = stdRuntimeJS.Create()
  
  'edge.Run "var a=1; a + 2"
End Sub
