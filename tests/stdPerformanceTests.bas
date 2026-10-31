Attribute VB_Name = "stdPerformanceTests"
'@lang VBA

Const C_MAX as long = 1000
Sub testAll()
    Test.Topic "stdPerformance"
    
    stdPerformance.MeasuresClear
    stdPerformance.AddMeasurement "Fractional", 0.25, 1
    stdPerformance.AddMeasurement "Fractional", 0.75, 1
    Test.Assert "Fractional mean", Abs(stdPerformance.Measurement("Fractional") - 0.5) < 0.0000001
    stdPerformance.AddMeasurement "Fractional", 0.125, 1
    Test.Assert "Fractional running mean", Abs(stdPerformance.Measurement("Fractional") - 0.375) < 0.0000001
    stdPerformance.AddMeasurement "Total", 12.5, 5
    Test.Assert "Operation count preserves total milliseconds", stdPerformance.Measurement("Total") = 12.5
    stdPerformance.MeasuresClear

#If Not Mac Then
    'Short scopes should retain fractions rather than whole GetTickCount intervals.
    Dim sample As Long, elapsed As Double, hasFraction As Boolean
    For sample = 1 To 100
        With stdPerformance.CreateMeasure("Resolution" & sample)
        End With
        elapsed = stdPerformance.Measurement("Resolution" & sample)
        Test.Assert "Nonnegative duration", elapsed >= 0
        If elapsed <> Fix(elapsed) Then hasFraction = True
    Next
    Test.Assert "Sub-millisecond timer precision", hasFraction
    stdPerformance.MeasuresClear
#End If

    'Measurement keys 1
    With stdPerformance.CreateMeasure("Test1")
    End With
    Test.Assert "MeasuresKeys 1", Join(stdPerformance.MeasuresKeys,";") = "Test1"
    
    'Measurement keys 2
    With stdPerformance.CreateMeasure("Test2")
    End With
    Test.Assert "MeasuresKeys 2", Join(stdPerformance.MeasuresKeys,";") = "Test1;Test2"

    'Only run in a full test
    if Test.FullTest then
        Test.Range.Parent.Activate
        DoEvents
        With stdPerformance.CreateMeasure("#1 Select and set")
            For i = 1 to C_MAX
                Test.Range.cells(1,1).select
                selection.value = "hello"
            Next
        End With

        With stdPerformance.CreateMeasure("#2 Set directly")
            For i = 1 to C_MAX
                Test.Range.cells(1,1).value = "hello"
            next
        End With
        DoEvents
        Test.Activate

        'GetMeasurement
        Test.Assert "[Get] Measurement", stdPerformance.Measurement("#1 Select and set") > 0
    end if

    'Clear all measurements
    stdPerformance.MeasuresClear

    'Remove items removes all keys
    Test.Assert "Clear removes items", ubound(stdPerformance.MeasuresKeys) - lbound(stdPerformance.MeasuresKeys) + 1 = 0

    'If key not found measurment is empty?
    Test.Assert "GetMeasurement null", stdPerformance.Measurement("#1 Select and set") = Empty

    'Optimise testing
    Application.ScreenUpdating = true
    Application.EnableEvents = true
    With stdPerformance.CreateOptimiser()
        Test.Assert "Optimise - 1a. ScreenUpdating=False", Not Application.ScreenUpdating
        Test.Assert "Optimise - 2a. EnableEvents=False", Not Application.EnableEvents
    End With
    Test.Assert "Optimise - 1b. ScreenUpdating restored", Application.ScreenUpdating
    Test.Assert "Optimise - 2b. EnableEvents restored", Application.EnableEvents
end sub