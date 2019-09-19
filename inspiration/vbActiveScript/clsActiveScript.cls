VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsActiveScriptSite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CopyRight 2003 David Zimmer
'Author:   David Zimmer <dzzie@yahoo.com>
'Site:     http://sandsprite.com
'
'License: You are free to use this code and type library in any commercial or
'         non commercial applications as you wish so long as this header remains
'         intact.
'
'Note do not delete the empty subs or functions, they need to be there because they
'are part fo the interface we have claimed to have implemented...


Implements IActiveScriptSite
Implements IActiveScriptSiteWindow 'this allows script to have a ui (like messageboxes)

Dim IScript As IActiveScript
Dim IParse As IActiveScriptParse

Private cObjects As New Collection

Event Error(pscripterror As Long)



Private Sub Class_Initialize()

    Dim hRes As Long
    Dim oUnk As vbActiveScript3.IUnknown
    Dim clsidVBS As vbActiveScript3.UUID
    Dim uuidActScr As vbActiveScript3.UUID
    
    'Create instance of the language engine
    CLSIDFromProgID "VBScript", clsidVBS
    CLSIDFromString IID_IActiveScript, uuidActScr
    
    'Start inproc script engine, VBSCRIPT.DLL
    hRes = CoCreateInstance(clsidVBS, Nothing, CLSCTX_INPROC_SERVER, uuidActScr, oUnk)
    
    'Get the IActiveScript interface
    Set IScript = oUnk
    
    IScript.SetScriptSite Me
        
    'get the iparse interface
    Set IParse = IScript
    
    IParse.InitNew
    
End Sub

Sub RunCode(sCode As String)
    
    Dim exep As vbActiveScript3.EXCEPINFO
     
    IParse.ParseScriptText sCode, _
                           Empty, _
                           Nothing, _
                           Empty, _
                           0, _
                           0, _
                           0, _
                           Null, _
                           exep
End Sub

Function Eval(sExpression As String) As Variant

    Dim exep As vbActiveScript3.EXCEPINFO
    Dim vAnswer As Variant

    IParse.ParseScriptText sExpression, _
                            Empty, _
                            Nothing, _
                            "", _
                            0, _
                            0, _
                            SCRIPTTEXT_ISEXPRESSION Or SCRIPTTEXT_ISVISIBLE, _
                            vAnswer, _
                            exep
                            
    Eval = vAnswer

End Function


Sub SetScriptState(scrState As SCRIPTSTATE)
    IScript.SetScriptState scrState
End Sub

Sub AddObject(sName As String, obj As Object)
    
    cObjects.Add obj, sName
    IScript.AddNamedItem StrPtr(sName), SCRIPTITEM_ISVISIBLE Or SCRIPTITEM_GLOBALMEMBERS

End Sub


Private Sub IActiveScriptSite_GetItemInfo(ByVal pstrName As String, ByVal dwReturnMask As vbActiveScript3.SCRIPTINFO, ppiunkItem As Long, ppti As Long)
     Dim IDisp As vbActiveScript3.IDispatch
     Dim obj As Object
     
     On Error Resume Next
     
     
     Set obj = cObjects(pstrName)
     If Err.Number = 0 Then
          If dwReturnMask = SCRIPTINFO_IUNKNOWN Then
                Set IDisp = obj
                Debug.Print "Asking for an Iunknown ** Using Ptr to Idispatch: " & ObjPtr(IDisp)
                ppiunkItem = ObjPtr(IDisp)
          Else
              Debug.Print "Looking for * to ITypeInfo, not supported"
          End If
          Set obj = Nothing
     Else
        Debug.Print "Cant find object named: " & pstrName
     End If
  
     
End Sub

Private Function IActiveScriptSite_GetLCID() As Long
    IActiveScriptSite_GetLCID = E_NOTIMPL
End Function

Private Function IActiveScriptSite_GetDocVersionString() As String

End Function

Private Sub IActiveScriptSite_OnEnterScript()

End Sub

Private Sub IActiveScriptSite_OnLeaveScript()

End Sub

Private Sub IActiveScriptSite_OnScriptError(pscripterror As Long)
   RaiseEvent Error(pscripterror)
End Sub


Private Sub IActiveScriptSite_OnScriptTerminate(pvarResult As Variant, pexcepinfo As vbActiveScript3.EXCEPINFO)

End Sub

Private Sub IActiveScriptSite_OnStateChange(ByVal ssScriptState As vbActiveScript3.SCRIPTSTATE)

End Sub

Private Sub IActiveScriptSiteWindow_EnableModeless(ByVal fEnable As vbActiveScript3.BOOL)
    
End Sub

Private Function IActiveScriptSiteWindow_GetWindow() As Long
    'without this our script could not show any kind of UI like messagboxes
    IActiveScriptSiteWindow_GetWindow = 0
End Function







'to execute a script you need to:
'Implement the IActiveScriptSite interface.
'Call CoCreateInstance and pass in the CLSID of the script engine.
'    Get an IActiveScript interface from the engine.
'Call IActiveScript::SetScriptSite and pass in the interface of your
'    IActiveScriptSite implementation.
'Call IActiveScript::QueryInterface and get an IActiveScriptParse interface.
'Call IActiveScriptParse::InitNew.
'Call IActiveScriptParse::ParseScriptText and pass in the script text along with flags SCRIPTTEXT_ISPERSISTENT|SCRIPTTEXT_ISVISIBLE.
'Call IActiveScript::AddNamedItem for each named object within your script.
'Call IActiveScript::SetScriptState(SCRIPTSTATE_CONNECTED).

'If the script engine needs you to create an instance of an object within the
'script, it will call IActiveScriptSite::GetItemInfo. You need to be able to
'create and return the appropriate object or type library.
