Attribute VB_Name = "stdAccTests"
'@lang VBA

Sub testAll()
  Test.Topic "stdAcc"

  Dim acc As stdAcc
  Set acc = stdAcc.CreateFromApplication()
  
  Dim o As stdAccTestHelper
  Dim tmp1 As stdAcc, tmp2 As stdAcc, tmpc As Collection
  
  'Only run on full test
  if Test.FullTest then
    'Continue searching but nothing found
    With stdLambda.Create("$1.Add(): EAccFindResult.NoMatchFound")
      Set o = New stdAccTestHelper
      Set tmp1 = acc.FindFirst(.Bind(o))
      Test.Assert "FindFirst & EAccFindResult.NoMatchFound - Returned value is nothing" , (tmp1 Is Nothing)
      Test.Assert "FindFirst & EAccFindResult.NoMatchFound - Scanned more than 1 element" , (o.Count > 1)
      
      Set o = New stdAccTestHelper
      Set tmpc = acc.FindAll(.Bind(o))
      Test.Assert "FindAll & EAccFindResult.NoMatchFound - Returned count is 0" , (tmpc.Count = 0)
      Test.Assert "FindAll & EAccFindResult.NoMatchFound - Scanned more than 1 element" , (o.Count > 1)
    End With
  end if
  
  'Return any item
  With stdLambda.Create("$1.Add(): EAccFindResult.MatchFound")
    Set o = New stdAccTestHelper
    Set tmp1 = acc.FindFirst(.Bind(o))
    Test.Assert "FindFirst & EAccFindResult.MatchFound - FindFirst does test itself for condition" , (tmp1.hwnd & "-" & tmp1.Role = acc.hwnd & "-" & acc.Role)
    Test.Assert "FindFirst & EAccFindResult.MatchFound - Should only scan 1 element" , (o.Count = 1)
  End With
  
  'Nothing found, but cancel function
  With stdLambda.Create("$1.Add(): EAccFindResult.NoMatchCancelSearch")
    Set o = New stdAccTestHelper
    Set tmp1 = acc.FindFirst(.Bind(o))
    Test.Assert "FindFirst & EAccFindResult.NoMatchCancelSearch - Ensure nothing returned when search is cancelled." , (tmp1 Is Nothing)
    Test.Assert "FindFirst & EAccFindResult.NoMatchCancelSearch - Ensure cancellation causes immediate cancellation." , (o.Count = 1)
    
    Set o = New stdAccTestHelper
    Set tmpc = acc.FindAll(.Bind(o))
    Test.Assert "FindAll & EAccFindResult.NoMatchCancelSearch - Ensure empty collection returned when search is cancelled." , (tmpc.Count = 0)
    Test.Assert "FindAll & EAccFindResult.NoMatchCancelSearch - Ensure cancellation causes immediate cancellation." , (o.Count = 1)
  End With
  
  'Nothing found, continue search, but don't search descendents
  With stdLambda.Create("$1.Add(): if $3 > 0 then EAccFindResult.NoMatchSkipDescendents else EAccFindResult.NoMatchFound")
    Set o = New stdAccTestHelper
    Set tmp1 = acc.FindFirst(.Bind(o))
    Test.Assert "FindFirst & EAccFindResult.NoMatchSkipDescendents + NoMatchFound - Ensure nothing returned regardless of descendent skipping" , (tmp1 Is Nothing)
    Test.Assert "FindFirst & EAccFindResult.NoMatchSkipDescendents + NoMatchFound - Ensure more items scanned, but not the entire tree of descendents." , (o.Count = 8)
    
    Set o = New stdAccTestHelper
    Set tmpc = acc.FindAll(.Bind(o))
    Test.Assert "FindAll & EAccFindResult.NoMatchSkipDescendents + NoMatchFound - Ensure nothing returned regardless of descendent skipping" , (tmpc.Count = 0)
    Test.Assert "FindAll & EAccFindResult.NoMatchSkipDescendents + NoMatchFound - Ensure more items scanned, but not the entire tree of descendents" , (o.Count = 8)
  End With
  
  'Only run on full test
  If Test.FullTest then
    Dim checkInButton As stdAcc
    Set checkInButton = acc.FindFirst(stdLambda.Create("$1.name = ""Check In..."""))
    Test.Assert "Test that acc role works" , (checkInButton.Role = "ROLE_MENUITEM")
  end if

  'TODO: Check remaining function
  'TODO: acc.PrintDesc
End Sub
