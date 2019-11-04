VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdEnumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Spec:
'For some enumerable set ["foo","bar","baz"] let stdEnumerator implement functions:
'* next
'* each
'* map
'* with_index
'...
'Can stack each call as follows:
'map.with_index

Private Enum BiState
    FromForEach
    FromCallback
    FromIEnumVariant
End Enum

Private oRetArray As New stdArray
Private mode As BiState
Private oEnumVariant As IEnumVARIANT
Private Iteration As Long
Private lowerBound As Long


Public Function FeedItem(v As Variant)
    Call oRetArray.push(v)
End Function

'Incompatible with oEnumVariant
Public Function NextItem() As Variant
    mode = iMode
    If mode = FromForEach Then
        
    ElseIf mode = FromCallback Then
        
    ElseIf mode = FromIEnumVariant Then
        Set NextItem = oEnumVariant.Next()
    End If
End Function
Public Function NextItems() As stdArray
    mode = iMode
    Set NextItems = stdArray.Create()
    If mode = FromForEach Then
        
    ElseIf mode = FromCallback Then
        
    ElseIf mode = FromIEnumVariant Then
        Dim x As IEnumVARIANT
        Set x = oEnumVariant.Clone
        
        Z = x.Next
        While Z
            Call NextItems.push(Z)
            Z = x.Next
        Wend
    End If
End Function
Public Function PeekItem() As Variant
    If mode = FromForEach Then
        
    ElseIf mode = FromCallback Then
        
    ElseIf mode = FromIEnumVariant Then
        Dim x As IEnumVARIANT
        Set x = oEnumVariant.Clone
        Set PeekItem = x.Next
    End If
End Function
Public Function PeekItems() As stdArray
    
End Function
Public Function Rewind() As stdEnumerator
    
End Function
Public Function Size() As Long

End Function




'Public Static stdEnumerator Create(IDispatch obj, string method, *args)
Public Function Create(ByVal cb As stdCallback) As stdEnumerator
    Set Create = New stdEnumerator
    Call Create.Init(FromForEach, cb)
End Function

'Implements each like:
'  For i = 1 to size: cb(size,metadata...): next
Public Function CreateFromCallback(ByVal Size As Long, ByVal cb As stdCallback) As stdEnumerator
    Set FromCallback = New stdEnumerator
    Call CreateFromCallback.Init(FromCallback, Size, cb)
End Function

Public Function CreateFromIEnumVariant(ByVal obj As Object) As stdEnumerator
    Set CreateFromIEnumVariant = New stdEnumerator
    Call CreateFromIEnumVariant.Init(FromIEnumVariant, obj)
    
End Function


Public Sub Init(iMode As BiState, ParamArray v() As Variant)
    mode = iMode
    If mode = FromForEach Then
        
    ElseIf mode = FromCallback Then
        
    ElseIf mode = FromIEnumVariant Then
        Set oEnumVariant = v(0)
    End If
End Sub



Public Function GetArray() As stdArray
    Set GetArray = oRetArray
End Function
Public Sub ForEach(Optional ByVal cb As stdCallback)
    Dim v As Variant
    If mode = FromForEach Then
        
    ElseIf mode = FromCallback Then
        
    ElseIf mode = FromIEnumVariant Then
        For Each v In oEnumVariant
            Call cb(v)
        Next
    End If
End Sub
Public Function Map(Optional ByVal cb As stdCallback) As stdArray
    Dim v As Variant
    If mode = FromForEach Then
        
    ElseIf mode = FromCallback Then
        
    ElseIf mode = FromIEnumVariant Then
        For Each v In oEnumVariant
            oRetArray.push (cb(v))
        Next v
    End If
End Function
Public Sub WithIndex(Optional ByVal cb As stdCallback)
    If mode = FromForEach Then
        
    ElseIf mode = FromCallback Then
        
    ElseIf mode = FromIEnumVariant Then
        
    End If
End Sub
Public Sub EachWithIndex(Optional ByVal cb As stdCallback)
    If mode = FromForEach Then
        
    ElseIf mode = FromCallback Then
        
    ElseIf mode = FromIEnumVariant Then
        
    End If
End Sub
Public Sub WithObject(Optional ByVal cb As stdCallback)
    Call EachWithObject(cb)
End Sub
Public Sub EachWithObject(Optional ByVal cb As stdCallback)
    If mode = FromForEach Then
        
    ElseIf mode = FromCallback Then
        
    ElseIf mode = FromIEnumVariant Then
        
    End If
End Sub



