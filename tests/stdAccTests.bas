Attribute VB_Name = "stdAccTests"

Sub testAll()
  Test.Topic "stdAcc"

  Dim acc As stdAcc
  Set acc = stdAcc.CreateFromApplication()
  
  Dim o As stdAccTestHelper
  Dim tmp1 As stdAcc, tmp2 As stdAcc, tmpc As Collection
  
  'Continue searching but nothing found
  With stdLambda.Create("$1#Add(): 0")
    Set o = New stdAccTestHelper
    Set tmp1 = acc.FindFirst(.Bind(o))
    Test.Assert "Return 2 a" , (tmp1 Is Nothing)
    Test.Assert "Return 2 b" , (o.Count > 1)
    
    Set o = New stdAccTestHelper
    Set tmpc = acc.FindAll(.Bind(o))
    Test.Assert "Return 2 b" , (tmpc.Count = 0)
    Test.Assert "Return 2 b" , (o.Count > 1)
  End With
  
  
  'Return any item
  With stdLambda.Create("$1#Add(): 1")
    Set o = New stdAccTestHelper
    Set tmp1 = acc.FindFirst(.Bind(o))
    Test.Assert "Return 1 a" , (tmp1.hwnd & "-" & tmp1.Role = acc.hwnd & "-" & acc.Role)
    Test.Assert "Return 1 b" , (o.Count = 1)
  End With
  
  'Nothing found, but cancel function
  With stdLambda.Create("$1#Add(): 2")
    Set o = New stdAccTestHelper
    Set tmp1 = acc.FindFirst(.Bind(o))
    Test.Assert "Return 2 a" , (tmp1 Is Nothing)
    Test.Assert "Return 2 b" , (o.Count = 1)
    
    Set o = New stdAccTestHelper
    Set tmpc = acc.FindAll(.Bind(o))
    Test.Assert "Return 2 c" , (tmpc.Count = 0)
    Test.Assert "Return 2 d" , (o.Count = 1)
  End With
  
  'Nothing found, continue search, but don't search descendents
  With stdLambda.Create("$1#Add(): if $3 > 0 then 3 else 0")
    Set o = New stdAccTestHelper
    Set tmp1 = acc.FindFirst(.Bind(o))
    Test.Assert "Return 3 a" , (tmp1 Is Nothing)
    Test.Assert "Return 3 b" , (o.Count = 8)
    
    Set o = New stdAccTestHelper
    Set tmpc = acc.FindAll(.Bind(o))
    Test.Assert "Return 3 c" , (tmpc.Count = 0)
    Test.Assert "Return 3 d" , (o.Count = 8)
  End With
  
  Dim checkInButton As stdAcc
  Set checkInButton = acc.FindFirst(stdLambda.Create("$1.name = ""Check In..."""))
  Test.Assert "Button Check In" , (checkInButton.Role = "ROLE_MENUITEM")
  
  'TODO: Check remaining function
  'TODO: acc.PrintDesc
End Sub
