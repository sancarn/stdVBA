# Ideas:

## Pointer Call

Code by TheTrick: http://www.vbforums.com/showthread.php?788413-VB6-Calling-functions-by-pointer&p=4942769&viewfull=1#post4942769

```vb

Option Explicit

Private Declare Function GetMem4 Lib "msvbvm60" ( _
                         ByRef src As Any, _
                         ByRef dst As Any) As Long
Private Declare Function VirtualProtect Lib "kernel32" ( _
                         ByVal lpAddress As Long, _
                         ByVal dwSize As Long, _
                         ByVal flNewProtect As Long, _
                         ByRef lpflOldProtect As Long) As Long
 
Private Const PAGE_EXECUTE_READWRITE = &H40

Private Type Vector2D
    posX As Single
    posY As Single
End Type

Private Declare Sub memcpy Lib "kernel32" _
                    Alias "RtlMoveMemory" ( _
                    ByRef Destination As Any, _
                    ByRef Source As Any, _
                    ByVal Length As Long)
                    
' // Buffer for exchanging
Dim buffer()    As Byte
Dim isInit      As Boolean


' // Helpers functions
Public Sub PatchFunc(ByVal Addr As Long)
    Dim InIDE As Boolean
 
    Debug.Assert MakeTrue(InIDE)
 
    If InIDE Then
        GetMem4 ByVal Addr + &H16, Addr
    Else
        VirtualProtect Addr, 8, PAGE_EXECUTE_READWRITE, 0
    End If

    GetMem4 &HFF505958, ByVal Addr
    GetMem4 &HE1, ByVal Addr + 4
End Sub
 
Public Function MakeTrue(ByRef bvar As Boolean) As Boolean
    bvar = True: MakeTrue = True
End Function


' // Calling of the standard functions using the pointers
Public Sub Main()
    Dim lngArray()  As Long
    Dim index       As Long
    
    ' // We're testing the function that sorts the long-array
    ReDim lngArray(99)
    
    For index = 0 To UBound(lngArray)
        lngArray(index) = Rnd * 100
    Next
    
    ' // Magic of the function pointers
    QuickSort VarPtr(lngArray(0)), UBound(lngArray) + 1, Len(lngArray(0)), AddressOf ComparatorLong
    
    ' // Now we're testing the function that sorts the string-array
    Dim strArray()  As String
    
    ReDim strArray(5)
    
    strArray(0) = "Calling"
    strArray(1) = "of the standard functions"
    strArray(2) = "using the pointers"
    strArray(3) = "on VB6"
    strArray(4) = "by The trick"
    strArray(5) = "2015"
    
    ' // We're calling same function using the magic of pointers
    QuickSort VarPtr(strArray(0)), UBound(strArray) + 1, 4, AddressOf ComparatorString
    
    ' // Now we're testing the function that sorts the UDT-array (2D-vectors)
    ' // For example we'll sorting the array by vector length
    Dim vecArray() As Vector2D
    
    ReDim vecArray(99)
    
    For index = 0 To UBound(vecArray)
        vecArray(index).posX = Rnd * 10
        vecArray(index).posY = Rnd * 10
    Next
    
    ' // We're calling same function for the sorting of the UDT-array
    QuickSort VarPtr(vecArray(0)), UBound(vecArray) + 1, LenB(vecArray(0)), AddressOf ComparatorVector2D
    
    ' // Test length
    For index = 0 To UBound(vecArray)
        Debug.Print Sqr(vecArray(index).posX ^ 2 + vecArray(index).posY ^ 2)
    Next
    
End Sub

' // This callback function which sorts two long values
Public Function ComparatorLong( _
                ByRef lItem1 As Long, _
                ByRef lItem2 As Long) As Long
    ComparatorLong = Sgn(lItem1 - lItem2)
End Function

' // This callback function which sorts two string values
Public Function ComparatorString( _
                ByRef lItem1 As String, _
                ByRef lItem2 As String) As Long
    ComparatorString = StrComp(lItem1, lItem2, vbTextCompare)
End Function

' // This callback function which sorts two 2D-vectors values by length
Public Function ComparatorVector2D( _
                ByRef lItem1 As Vector2D, _
                ByRef lItem2 As Vector2D) As Long
    ' // Optimize sqr
    ComparatorVector2D = Sgn((lItem1.posX * lItem1.posX + lItem1.posY * lItem1.posY) - _
                             (lItem2.posX * lItem2.posX + lItem2.posY * lItem2.posY))
End Function

' // Quick-sort using the callback function for a comparing
' // This function uses callback function (lpfnComparator)
Public Sub QuickSort( _
           ByVal lpFirstPtr As Long, _
           ByVal lNumOfItems As Long, _
           ByVal lSizeElement As Long, _
           ByVal lpfnComparator As Long)
           
    Dim lpI     As Long
    Dim lpJ     As Long
    Dim lpM     As Long
    Dim lpLast  As Long
    
    If Not isInit Then
        ' // Initialize patching and buffer for exchanging
        ReDim buffer(lSizeElement - 1)
        PatchFunc AddressOf MainComparator
        isInit = True
        
    End If
    
    lpLast = lpFirstPtr + (lNumOfItems - 1) * lSizeElement
    lpI = lpFirstPtr
    lpJ = lpLast
    lpM = lpFirstPtr + ((lNumOfItems - 1) \ 2) * lSizeElement

    Do Until lpI > lpJ
        
        ' // Call function that being passed into the lpfnComparator parameter
        Do While MainComparator(lpfnComparator, lpI, lpM) = -1
            lpI = lpI + lSizeElement
        Loop
        
        ' // Call function that being passed into the lpfnComparator parameter
        Do While MainComparator(lpfnComparator, lpJ, lpM) = 1
            lpJ = lpJ - lSizeElement
        Loop
        
        ' // Exchanging
        If (lpI <= lpJ) Then
            
            If lpI = lpM Then
                lpM = lpJ
            ElseIf lpJ = lpM Then
                lpM = lpI
            End If
            
            If lSizeElement > UBound(buffer) + 1 Then
                ReDim buffer(lSizeElement - 1)
            End If
            
            memcpy buffer(0), ByVal lpI, lSizeElement
            memcpy ByVal lpI, ByVal lpJ, lSizeElement
            memcpy ByVal lpJ, buffer(0), lSizeElement
  
            lpI = lpI + lSizeElement
            lpJ = lpJ - lSizeElement
            
        End If
        
    Loop

    If lpFirstPtr < lpJ Then
        QuickSort lpFirstPtr, (lpJ - lpFirstPtr) \ lSizeElement + 1, lSizeElement, lpfnComparator
    End If
    
    If lpI < lpLast Then
        QuickSort lpI, (lpLast - lpI) \ lSizeElement + 1, lSizeElement, lpfnComparator
    End If
    
End Sub

' // Prototype for comparator function
' // If lpItem1 > lpItem2 then function return 1
' // If lpItem1 = lpItem2 then function return 0
' // If lpItem1 < lpItem2 then function return -1
Public Function MainComparator( _
                ByVal lpAddressOfFunction As Long, _
                ByVal lpItem1 As Long, _
                ByVal lpItem2 As Long) As Long
End Function
```

### Core pattern:

```
sub Main
  Debug.Print doSomething(AddressOf theCallback)
end sub

Function theCallback(arg1, arg2, arg3, ...)

End Function

Function doSomething(ByVal callback as long)
  Call PatchFunc(AddressOf CallCallback)
  '...
  doSomething = CallCallback(callback, arg1, arg2, arg3, ...)
End Function

Function CallCallback(ptr as long, arg1, arg2, arg3, ...)

End Function
```


