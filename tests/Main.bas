<<<<<<< HEAD
Attribute VB_Name = "Main"

Sub MainTestAll()
    Test.refresh

    On Error Resume Next
    Call stdLambdaTests.testAll
    Call stdArrayTests.testAll
    Call stdCallbackTests.testAll
    Call stdEnumeratorTests.testAll
=======
Attribute VB_Name = "Main"

Sub MainTestAll()
    Test.refresh

    On Error Resume Next
    Call stdLambdaTests.testAll
    Call stdArrayTests.testAll
    Call stdCallbackTests.testAll
    Call stdAccTests.testAll
    Call stdEnumeratorTests.testAll
    Call stdClipboardTests.testAll
    Call stdRegexTests.testAll
    Call stdWindowTests.testAll
    Call stdWebSocketTests.testAll
    Call stdProcessTests.testAll
>>>>>>> master
End Sub