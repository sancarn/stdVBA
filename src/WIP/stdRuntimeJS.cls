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
'
'Example:
'
' set js = stdRuntimeJS.Create(EngineEdge)
' js.addScript("function test(){console.log(""Hello World"")}")
' js.run("test()")

Private Const S_OK = 0

#If Win64 then
  Private Const NULL_PTR as LongLong = 0^
#Else
  Private Const NULL_PTR as Long = 0&
#End If

#if VBA7 then
  Private Declare PtrSafe Function EdgeCreateRuntime Lib "chakra.dll" Alias "JsCreateRuntime" (ByVal jsRuntimeAttributes As Integer, ByVal JsThreadServiceCallback As LongPtr, ByVal jsRuntimeHandle As LongPtr) As Long 'Note: ThreadServiceCallback seems pointless in VBA
  Private Declare PtrSafe Function EdgeCreateContext Lib "chakra.dll" Alias "JsCreateContext" (ByVal jsRuntimeHandle As LongPtr, ByVal jsContextRef As LongPtr) As Long
  Private Declare PtrSafe Function EdgeSetCurrentContext Lib "chakra.dll" Alias "JsSetCurrentContext" (ByVal jsContextRef As LongPtr) As Long
  Private Declare PtrSafe Function EdgeGetGlobalObject Lib "chakra.dll" Alias "JsGetGlobalObject" (ByVal jsGlobalValueRef As LongPtr) As Long
  Private Declare PtrSafe Function EdgeDisposeRuntime Lib "chakra.dll" Alias "JsDisposeRuntime" (ByVal jsRuntimeHandle As LongPtr) As Long
  Private Declare PtrSafe Function EdgeVariantToValue Lib "chakra.dll" Alias "JsVariantToValue" (ByVal src As LongPtr, ByVal dest As LongPtr) As Long
  Private Declare PtrSafe Function EdgeValueToVariant Lib "chakra.dll" Alias "JsValueToVariant" (ByVal src As LongPtr, ByVal dest As LongPtr) As Long
  Private Declare PtrSafe Function EdgeRunScript Lib "chakra.dll" Alias "JsRunScript" (ByVal code As String, ByVal JsSourceContext As Integer, ByVal JsSourceURL As String, ByVal JsValueRef As LongPtr) As Long
  Private Declare PtrSafe Function EdgeGetAndClearException Lib "chakra.dll" Alias "JsGetAndClearException" (ByVal JsValueRef As LongPtr) As Long
  Private Declare PtrSafe Function EdgeProjectWinRTNamespace Lib "chakra.dll" Alias "JsProjectWinRTNamespace" (ByVal sNamespace As String) As Long
  Private Declare PtrSafe Function EdgeConvertValueToString Lib "Chakra.dll" Alias "JsConvertValueToString" (ByVal m_JsValue As LongPtr, ByVal VARPTR_RESULstr As LongPtr) As Long
  Private Declare PtrSafe Function EdgeStringToPointer Lib "Chakra.dll" Alias "JsStringToPointer" (ByVal m_JsValue As LongPtr, stringValue As LongPtr, stringLength As Long) As Long
  Private Declare PtrSafe Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As LongPtr, Optional ByVal pszStrPtr As LongPtr) As Long
#Else
  Enum LongPtr
    [_]
  End Enum
  Private Declare Function EdgeCreateRuntime Lib "chakra.dll" Alias "JsCreateRuntime" (ByVal jsRuntimeAttributes As Integer, ByVal JsThreadServiceCallback As LongPtr, ByVal jsRuntimeHandle As LongPtr) As Long 'Note: ThreadServiceCallback seems pointless in VBA
  Private Declare Function EdgeCreateContext Lib "chakra.dll" Alias "JsCreateContext" (ByVal jsRuntimeHandle As LongPtr, ByVal jsContextRef As LongPtr) As Long
  Private Declare Function EdgeSetCurrentContext Lib "chakra.dll" Alias "JsSetCurrentContext" (ByVal jsContextRef As LongPtr) As Long
  Private Declare Function EdgeGetGlobalObject Lib "chakra.dll" Alias "JsGetGlobalObject" (ByVal jsGlobalValueRef As LongPtr) As Long
  Private Declare Function EdgeDisposeRuntime Lib "chakra.dll" Alias "JsDisposeRuntime" (ByVal jsRuntimeHandle As LongPtr) As Long
  Private Declare Function EdgeVariantToValue Lib "chakra.dll" Alias "JsVariantToValue" (ByVal src As LongPtr, ByVal dest As LongPtr) As Long
  Private Declare Function EdgeValueToVariant Lib "chakra.dll" Alias "JsValueToVariant" (ByVal src As LongPtr, ByVal dest As LongPtr) As Long
  Private Declare Function EdgeRunScript Lib "chakra.dll" Alias "JsRunScript" (ByVal code As String, ByVal JsSourceContext As Integer, ByVal JsSourceURL As String, ByVal JsValueRef As LongPtr) As Long
  Private Declare Function EdgeGetAndClearException Lib "chakra.dll" Alias "JsGetAndClearException" (ByVal JsValueRef As LongPtr) As Long
  Private Declare Function EdgeProjectWinRTNamespace Lib "chakra.dll" Alias "JsProjectWinRTNamespace" (ByVal sNamespace As String) As Long
  Private Declare Function EdgeConvertValueToString Lib "Chakra.dll" Alias "JsConvertValueToString" (ByVal m_JsValue As LongPtr, ByVal VARPTR_RESULstr As LongPtr) As Long
  Private Declare Function EdgeStringToPointer Lib "Chakra.dll" Alias "JsStringToPointer" (ByVal m_JsValue As LongPtr, stringValue As LongPtr, stringLength As Long) As Long
  Private Declare Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As LongPtr, Optional ByVal pszStrPtr As LongPtr) As Long
#End If

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

Public Enum EJavascriptEngineType
  EngineInternetExplorer
  EngineEdge
End Enum

Private Type TThis
  EngineType as EJavascriptEngineType
  hRuntime As LongPtr
  hContext As LongPtr
  hGlobal as LongPtr
  Error as String
End Type
Private This as TThis


'select case This.EngineType
'  case EJavascriptEngineType.EngineEdge
'  case EJavascriptEngineType.EngineInternetExplorer
'end select



Public Sub Test
  Dim rt as stdRuntimeJS: set rt = Create()
  Debug.assert false
End Sub

Public Function Create(Optional ByVal iEngineType As EJavascriptEngineType = EJavascriptEngineType.EngineEdge) As stdRuntimeJS
  Set Create = New stdRuntimeJS
  Call Create.protInit(iEngineType)
End Function


Public Sub protInit(iEngineType As EJavascriptEngineType)
  This.EngineType = iEngineType
  select case This.EngineType
    case EJavascriptEngineType.EngineEdge
      'Create the javascript runtime...
      Call ErrorGuard("EdgeCreateRuntime",     EdgeCreateRuntime(0, 0, VarPtr(This.hRuntime))) 
      Call ErrorGuard("EdgeCreateContext",     EdgeCreateContext(This.hRuntime, VarPtr(This.hContext))) 
      Call ErrorGuard("EdgeSetCurrentContext", EdgeSetCurrentContext(This.hContext))
      Call ErrorGuard("EdgeGetGlobalObject",   EdgeGetGlobalObject(VarPtr(This.hGlobal)))
    case EJavascriptEngineType.EngineInternetExplorer
      'TODO:
      Err.Raise 1, "stdRuntimeJS#protInit", "Internet Explorer mode is planned but not implemented yet. Use edge for now!"
    case else
      Err.Raise 1, "stdRuntimeJS#protInit", "Engine type not defined"
  end select
End Sub

Public Sub ErrorGuard(ByVal sFuncName as string, ByVal ErrorCode as Long)
  If ErrorCode <> S_OK then
    Err.Raise 1, "stdRuntimeJS#" & sFuncName, "Runtime error occurred in stdRuntimeJS. " & sFuncName & " function returned error code: " & ErrorCode & "."
  end if
End Sub

'Loads the file given
Public Sub Require(ByVal sPath As String)
  Dim sContent as string
  Dim ff as Long: ff = FreeFile()
  Open sPath For Input as #ff
    Input #ff, sContent
  Close #ff
  Call Run(sContent)
End Sub


Public Function Run(ByVal script As String)
  Dim v As LongPtr, sError As String
  select case This.EngineType
    Case EJavascriptEngineType.EngineInternetExplorer
      'Call IERunScript(script, 0, "source.js", VarPtr(v))
    Case EJavascriptEngineType.EngineEdge
      Call EdgeRunScript(script, 0, "source.js", VarPtr(v))
  End Select
End Function

Public Function Eval()
  select case This.EngineType
    case EJavascriptEngineType.EngineEdge
    case EJavascriptEngineType.EngineInternetExplorer
  end select
End Function

Public Sub AddObject(ByVal sName As String, ByRef obj As Object)
  select case This.EngineType
    case EJavascriptEngineType.EngineEdge
    case EJavascriptEngineType.EngineInternetExplorer
  end select
End Sub

Public Sub AddFunction(ByVal sName as string, ByVal func as LongPtr)
  select case This.EngineType
    case EJavascriptEngineType.EngineEdge
    case EJavascriptEngineType.EngineInternetExplorer
  end select
End Sub

Public Sub ProjectWinRTNamespace(sNamespace As String)
  select case This.EngineType
    Case EJavascriptEngineType.EngineInternetExplorer
      Err.Raise 1, "", "Projecting WinRT namespaces is not possible in IE9"
    Case EJavascriptEngineType.EngineEdge
      Call EdgeProjectWinRTNamespace(sNamespace)
  End Select
End Sub



Private Sub Class_Terminate()
  Call Terminate
End Sub

Public Sub Terminate()
  select case This.EngineType
    case EJavascriptEngineType.EngineEdge
      Call EdgeDisposeRuntime(VarPtr(pRuntime))
    case EJavascriptEngineType.EngineInternetExplorer
      'Call IEDisposeRuntime(VarPtr(pRuntime))
  end select
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

Private Function JsValueToSTR(m_JsValue As LongPtr) As String
  Dim StringLen As Long
  Dim JsStringPtr As LongPtr: Call ErrorGuard("JsConvertValueToString", JsConvertValueToString(m_JsValue, VarPtr(JsStringPtr)))
  Dim VbStringPtr As LongPtr: Call ErrorGuard("JsStringToPointer",      JsStringToPointer(JsStringPtr, VbStringPtr, StringLen))
  JsValueToSTR = GetStrFromPtrw(VbStringPtr)
End Function

Private Function GetStrFromPtrw(ByVal Ptr As LongPtr) As String
  if SysReAllocString(VarPtr(GetStrFromPtrw), Ptr) = 0 then
    Err.Raise 1, "stdRuntimeJS#SysReAllocString", "Runtime error occurred in stdRuntimeJS.SysReAllocString function."
  end if
End Function



















'JsCreateFunction(&setTimeoutCallback, nullptr, &function);
'JsCreatePropertyId("setTimeout", 10, &propertyId);
'JsSetProperty(windowObject, propertyId, function, true);








''FOR IE
''----------------------------------------
''Not sure what the enumeration here is, but AHK uses -1 for this param so maybe -1 is latest?
'Private Enum InternetExplorerJSRuntimeType
'  InternetExplorerJSRT_10
'  InternetExplorerJSRT_11 = -1
'  InternetExplorerJSRT_Edge 'Deprecated
'End Enum
''For initialisation:
''Private Declare Function IECreateRuntime Lib "jscript9.dll" Alias "JsCreateRuntime" (jsRuntimeAttributes As Integer, InternetExplorerJSRuntimeType As Long, JsThreadServiceCallback As LongPtr, jsRuntimeHandle As LongPtr) As Long
'Private Declare PtrSafe Function IECreateRuntime Lib "jscript9.dll" Alias "JsCreateRuntime" (ByVal jsRuntimeAttributes As Integer, ByVal InternetExplorerJSRuntimeType As Long, ByVal JsThreadServiceCallback As LongPtr, ByVal jsRuntimeHandle As LongPtr) As Long
'Private Declare PtrSafe Function IECreateContext Lib "jscript9.dll" Alias "JsCreateContext" (ByVal jsRuntimeHandle As LongPtr, ByVal IDebugApplication As LongPtr, ByVal jsContextRef As LongPtr) As Long
'Private Declare PtrSafe Function IESetCurrentContext Lib "jscript9.dll" Alias "JsSetCurrentContext" (ByVal jsContextRef As LongPtr) As Long
'Private Declare PtrSafe Function IEGetGlobalObject Lib "jscript9.dll" Alias "JsGetGlobalObject" (ByVal jsGlobalValueRef As LongPtr) As Long
'
''For uninitialisation:
'Private Declare PtrSafe Function IEDisposeRuntime Lib "jscript9.dll" Alias "JsDisposeRuntime" (ByVal jsRuntimeHandle As LongPtr) As Long
'
''Getting/Setting values:
'Private Declare PtrSafe Function IEVariantToValue Lib "jscript9.dll" Alias "JsVariantToValue" (ByVal src As LongPtr, ByVal dest As LongPtr) As Long
'Private Declare PtrSafe Function IEValueToVariant Lib "jscript9.dll" Alias "JsValueToVariant" (ByVal src As LongPtr, ByVal dest As LongPtr) As Long
'
''Executing JavaScript
'Private Declare PtrSafe Function IERunScript Lib "jscript9.dll" Alias "JsRunScript" (ByVal code As String, ByVal JsSourceContext As Integer, ByVal JsSourceURL As String, ByVal JsValueRef As LongPtr) As Long
'
''Getting exception information
'Private Declare PtrSafe Function IEGetAndClearException Lib "jscript9.dll" Alias "JsGetAndClearException" (ByVal JsValueRef As LongPtr) As Long