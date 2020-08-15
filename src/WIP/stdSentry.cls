VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdSentry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'In numerous cases using excel we like to set standard options to optimise performance:
'
'    Dim iEnableEvents as long: iEnableEvents = Application.EnableEvents
'    Dim iScreenUpdating as long: iScreenUpdating = Application.ScreenUpdating
'    Application.EnableEvents = false
'    Application.ScreenUpdating = false
'    
'      ... our code ...
'    
'    Application.EnableEvents = iEnableEvents
'    Application.ScreenUpdating = iScreenUpdating
'
'Similarly with std libraries, we have to add and pop from the stack:
'
'    stdError.addStack "myMethod"
'      ... our code ...
'    stdError.popStack
'    stdPerformance.monitor "myMethod"
'  
'    stdPerformance.stopMonitor
'
'This method could even lead to errors where popStack/stopMonitor being forgotton.
'
'A better pattern is as follows:
'
'    Sub test()
'      With stdError.getSentry("test")
'        With Performance.getSentry() 'sets Application.EnableEvents to false on initialise and to true on destroy
'              
'              '... our code ...
'    
'        End With
'      End With
'    End Sub
'
'Learnt this pattern from GSerg here https://stackoverflow.com/a/20718181/6302131
'
'Generic sentry pattern:
'
'   stdSentry.Create(ByVal OnDestroy as ICallable, ParamArray args() as variant)
'
'So for stdError:
'
'    Public Function getSentry(ByVal sSubName as string) as stdSentry
'      Call stdError.addStack(sSubName)
'      set getSentry = stdSentry.Create(stdCallback.CreateFromObjectMethod(stdError,"popStack"))
'    End Function
'
'The stdSentry is implemented very simply:
Private pOnDestroy as stdICallable
Private params as variant

Public Function Create(ByVal OnDestroy as stdICallable, Optional args as variant = Empty)
  set Create = new stdSentry
  Call Create.Init(OnDestroy,params)
End Function

Public Sub Init(ByVal OnDestroy as stdICallable, args as variant)
  set pOnDestroy = OnDestroy
  If IsObject(args) then 
    set params = args
  Else
    params = args
  End if
End Sub

Private Sub Class_Terminate()
  if not OnDestroy is nothing then
    Call OnDestroy.RunEx(args)
  End if
End Sub



