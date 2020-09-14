Attribute VB_Name = "Main"

Sub MainTestAll()
    Test.refresh

    On Error Resume Next
    Call stdLambdaTests.testAll
    Call stdArrayTests.testAll
    
End Sub