Attribute VB_Name = "stdLambdaTests"
Sub arithmetic()
    '(3*(2+5)+5*8/2^(2+1))/26
    '=(3*7+5*8/2^3)/26
    '=(21+5*8/8)/26
    '=(21+5*1)/26
    '=(21+5)/26
    '=26/26
    '=1
    Debug.Print stdLambda.Create("(3*(2+5)+5*8/2^(2+1))/26").Run()
End Sub

Sub logicAndComparison()
    Debug.Print stdLambda.Create("5<3 or 5>3").Run()
End Sub

Sub arguments()
    Debug.Print stdLambda.Create("$1 + $2").Run(5, 9)
End Sub

Sub objects()
    'Evaluate property access
    Debug.Print stdLambda.Create("$1.Range(""A1"")").Run(Sheets(1)).Address(True, True, xlA1, True)
    
    'Evaluate method access
    Range("A1").value = 1
    Range("A2").value = 2
    Range("A3").value = 3
    Range("A4").value = 4
    Debug.Print stdLambda.Create("$1#Find(3)").Run(Range("A:A")).Address(True, True, xlA1, True)
End Sub

Sub inlineif()
    Dim lambda As Variant
    Set lambda = stdLambda.Create("if $1 then 0 else if $2 then 1 else 1 + 1")
    Debug.Print lambda.Run(True, True)
    Debug.Print lambda.Run(False, True)
    Debug.Print lambda.Run(False, False)
End Sub

Sub funcs()
    Debug.Print stdLambda.Create("uCase(trim(""          oranges        "")) & len(""potatoes"")").Run()
End Sub

Sub performanceTest1()
    'Evaluate method access
    Range("A1").value = 1
    Range("A2").value = 2
    Range("A3").value = 3
    Range("A4").value = 4
    Dim lambda As Variant

    iStart = Timer
    Set lambda = stdLambda.Create("$1#Find(3)")
    For i = 1 To 10 ^ 3
        Call lambda.Run(Range("A:A"))
    Next
    Debug.Print "StdLambda: " & (Timer - iStart)

    iStart = Timer
    Set lambda = stdLambdaOld.Create("$1#Find(3)")
    For i = 1 To 10 ^ 3
        Call lambda.Run(Range("A:A"))
    Next
    Debug.Print "StdLambdaOld: " & (Timer - iStart)
End Sub

Sub performanceTest2()
    'Evaluate method access
    Range("A1").value = 1
    Range("A2").value = 2
    Range("A3").value = 3
    Range("A4").value = 4
    Dim lambda As Variant
    
    Formula = "0+(3*(2+5)+5*8/2^(2+1))/26-1+(3*(2+5)+5*8/2^(2+1))/26-1+(3*(2+5)+5*8/2^(2+1))/26-1+(3*(2+5)+5*8/2^(2+1))/26-1+(3*(2+5)+5*8/2^(2+1))/26-1+(3*(2+5)+5*8/2^(2+1))/26-1+(3*(2+5)+5*8/2^(2+1))/26-1+(3*(2+5)+5*8/2^(2+1))/26-1"
    Debug.Print "10^3 * """ & Formula & """"
    
    iStart = Timer
    Set lambda = stdLambda.Create(Formula)
    With lambda
        For i = 1 To 10 ^ 3
            Call .Run
        Next
    End With
    Debug.Print "StdLambda: " & (Timer - iStart)
    
    iStart = Timer
    Set lambda = stdLambdaOld.Create(Formula)
    For i = 1 To 10 ^ 3
        Call lambda.Run
    Next
    Debug.Print "StdLambdaOld: " & (Timer - iStart)
End Sub
