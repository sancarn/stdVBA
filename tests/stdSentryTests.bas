Attribute VB_Name = "stdSentryTests"

Private temp as boolean

Sub testAll()
  test.Topic "stdSentry"

  'Test CreateFromObjectProperty
  Test.Assert "Before CreateFromObjectProperty", Application.DisplayAlerts = True
  With stdSentry.CreateFromObjectProperty(Application, "DisplayAlerts", False)
    Test.Assert "During CreateFromObjectProperty", Application.DisplayAlerts = False
  End With
  Test.Assert "After CreateFromObjectProperty", Application.DisplayAlerts = True
  
  'Test CreateFromObjectMethod
  Dim c As Collection: Set c = New Collection
  Test.Assert "Before CreateFromObjectMethod", c.count = 0
  With stdSentry.CreateFromObjectMethod(c, "Add", Array(""), "Remove", Array(1))
    Test.Assert "During CreateFromObjectMethod", c.count = 1
  End With
  Test.Assert "After CreateFromObjectMethod", c.count = 0

  'Test CreateOptimiser
  Dim t1 as Boolean: t1 = Application.EnableEvents
  Dim t2 as Boolean: t2 = Application.ScreenUpdating
  Dim t3 as XlCalculation: t3 = Application.Calculation
  Application.EnableEvents = True
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  With stdSentry.CreateOptimiser(false,false,xlCalculationManual)
    Test.Assert "During Optimiser 1", Application.EnableEvents = False
    Test.Assert "During Optimiser 2", Application.ScreenUpdating = False
    Test.Assert "During Optimiser 3", Application.Calculation = xlCalculationManual
  End With
  Test.Assert "After Optimiser 1", Application.EnableEvents = True
  Test.Assert "After Optimiser 2", Application.ScreenUpdating = True
  Test.Assert "After Optimiser 3", Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = t1
  Application.ScreenUpdating = t2
  Application.Calculation = t3
  
  'Ensure optimiser doesn't change values by default
  With stdSentry.CreateOptimiser()
    Test.Assert "During Optimiser - unchanged by default 1", Application.EnableEvents = t1
    Test.Assert "During Optimiser - unchanged by default 2", Application.ScreenUpdating = t2
    Test.Assert "During Optimiser - unchanged by default 3", Application.Calculation = t3
  End With

  'Test embedded change
  Test.Assert "Embedded change - Before ScreenUpdating 1", Application.ScreenUpdating = t2
  With stdSentry.CreateOptimiser(ScreenUpdating:=false)
    Test.Assert "Embedded change - During ScreenUpdating 1", Application.ScreenUpdating = false
    With stdSentry.CreateOptimiser(ScreenUpdating:=true)
      Test.Assert "Embedded change - During ScreenUpdating 1.1", Application.ScreenUpdating = true
    End With
    Test.Assert "Embedded change - After ScreenUpdating 1.1", Application.ScreenUpdating = false
  End With
  Test.Assert "Embedded change - After ScreenUpdating 1", Application.ScreenUpdating = t2

  'Test custom change
  
  Dim cbSetTemp as stdCallback: set cbSetTemp = stdCallback.CreateFromModule("stdSentryTests","setTemp")
  Dim cbResetTemp as stdCallback: set cbResetTemp = stdCallback.CreateFromModule("stdSentryTests","resetTemp")
  Dim TempSetter as stdSentry: set TempSetter = stdSentry.Create(cbSetTemp, cbResetTemp, False)
  Test.Assert "Before Temp change", temp = false
  With TempSetter()
    Test.Assert "During Temp change", temp = true
  End With
  Test.Assert "After Temp change", temp = false 
End Sub

Public Sub setTemp()
  temp = true
End Sub
Public Sub resetTemp()
  temp = false
End Sub