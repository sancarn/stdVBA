VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = true

'----------------------------------------------
' HASHTABLE class module
'
' This class implements a hashtable, a structure that offers many
' of the features of a collection or dictionary, and is often
' even faster than the built-in collection.
'
' NOTE: must make Item the default member, using the Tools | Procedure 
' Attributes dialog
'
' Usage:
'   Dim ht As New HashTable
'   ht.protSetSize 10000           ' initial number of slots (the higher,
'  the better)
'
'   ' enforce case-insensitive key search
'   ht.IgnoreCase = True
'   ' add values
'   ht.Add "key", value        ' add a value associated to a key
'   ' count how many values are in the table
'   Print ht.Count
'   ' read/write a value
'   Print ht("key")
'   ht("key") = newValue
'
'   ' remove a value
'   ht.Remove "key"
'   ' remove all values
'   ht.RemoveAll
'   ' check whether a value exists
'   If ht.Exists("key") Then ...
'
'   ' get the array of keys and values
'   Dim keys() As String, values() As Variant
'   keys() = ht.Keys
'   values() = ht.Values
'
'----------------------------------------------

' default values
Private Const DEFAULT_HASHSIZE = 10
Private Const DEFAULT_CAPACITY = 10 'The number of slots expected to be required typically. Note: You can override these using the CreateAdvanced() function.

Private Const DEFAULT_CHUNKSIZE = 1024

Private Type Bucket
    isRoot as boolean
    Key As String
    Value As Variant
    nextItem As Long      ' 0 if last item
End Type

'hashTbl() indices (of which are indices generated from hashFunction) points to the index of an element
'in slotTable() array.
Private hashTbl() As Long

'Array of buckets
Private slotTable() As Bucket

' pointer to first free slot
Private FreeNdx As Long


Private pHashSize As Long  'size of hash table
Private pListSize As Long  'size of slot table
Private pChunkSize As Long 'chunk size
Private pCount As Long     'items in the slot table

' member variable for IgnoreCase property
' True if keys are searched in case-unsensitive mode
' this can be assigned to only when the hash table is empty
Private pIgnoreCase As Boolean




Public Function Create(ParamArray keyItems() as variant) as stdDictionary
  set Create = new stdDictionary
  Call Create.protInit()
  For i = 0 to ubound(keyItems)
    Call Create.Add(keyItems(i),keyItems(i+1))
  next
End Function
Public Function CreateAdvanced(HashSize,ListSize,ChunkSize,ParamArray keyItems() as variant) as stdDictionary
  set Create = new stdDictionary
  Call Create.protInit()
  For i = 0 to ubound(keyItems)
    Call Create.Add(keyItems(i),keyItems(i+1))
  next
End Function



Property Get IgnoreCase() As Boolean
    IgnoreCase = pIgnoreCase
End Property

Property Let IgnoreCase(ByVal newValue As Boolean)
    If pCount Then
        Err.Raise 1001, , "Dictionary isn't empty"
    End If
    pIgnoreCase = newValue
End Property


' check whether an item is in the hash table
Function Exists(Key As String) As Boolean
    Exists = GetSlotIndex(Key) <> 0
End Function

' add a new element to the hash table

Sub Add(Key As String, Value As Variant)
    Dim ndx As Long, Create As Boolean
    
    ' get the index to the slot where the value is
    ' (allocate a new slot if necessary)
    Create = True
    ndx = GetSlotIndex(Key, Create)
    
    If Create Then
        ' the item was actually added
        If IsObject(Value) Then
            Set slotTable(ndx).Value = Value
        Else
            slotTable(ndx).Value = Value
        End If
    Else
        ' raise error "This key is already associated with an item of this 
        ' collection"
        Err.Raise 457
    End If
End Sub

' the value associated to a key
' (empty if not found)

Property Get Item(Key As String) As Variant
    Dim ndx As Long
    ' get the index to the slot where the value is
    ndx = GetSlotIndex(Key)
    If ndx = 0 Then
        ' return Empty if not found
    ElseIf IsObject(slotTable(ndx).Value) Then
        Set Item = slotTable(ndx).Value
    Else
        Item = slotTable(ndx).Value
    End If
End Property

Property Let Item(Key As String, Value As Variant)
    Dim ndx As Long
    ' get the index to the slot where the value is
    ' (allocate a new slot if necessary)
    ndx = GetSlotIndex(Key, True)
    ' store the value
    slotTable(ndx).Value = Value
End Property

Property Set Item(Key As String, Value As Object)
    Dim ndx As Long
    ' get the index to the slot where the value is
    ' (allocate a new slot if necessary)
    ndx = GetSlotIndex(Key, True)
    ' store the value
    Set slotTable(ndx).Value = Value
End Property

' remove an item from the hash table

Sub Remove(Key As String)
    Dim ndx As Long, HCode As Long, LastNdx As Long
    ndx = GetSlotIndex(Key, False, HCode, LastNdx)
    ' raise error if no such element
    If ndx = 0 Then Err.Raise 5
    
    If LastNdx Then
        ' this isn't the first item in the slotTable() array
        slotTable(LastNdx).nextItem = slotTable(ndx).nextItem
    ElseIf slotTable(ndx).nextItem Then
        ' this is the first item in the slotTable() array
        ' and is followed by one or more items
        hashTbl(HCode) = slotTable(ndx).nextItem
    Else
        ' this is the only item in the slotTable() array
        ' for this hash code
        hashTbl(HCode) = 0
    End If
    
    ' put the element back in the free list
    slotTable(ndx).nextItem = FreeNdx
    FreeNdx = ndx
    ' we have deleted an item
    pCount = pCount - 1
    
End Sub

' remove all items from the hash table

Sub RemoveAll()
    protSetSize pHashSize, pListSize, pChunkSize
End Sub

' the number of items in the hash table

Property Get Count() As Long
    Count = pCount
End Property

' the array of all keys
' (VB5 users: convert return type to Variant)

Property Get Keys() As Variant()
    Dim i As Long, ndx As Long
    Dim n As Long
    ReDim res(0 To pCount - 1) As Variant
    
    For i = 0 To pHashSize - 1
        ' take the pointer from the hash table
        ndx = hashTbl(i)
        ' walk the slottable() array
        Do While ndx
            res(n) = slotTable(ndx).Key
            n = n + 1
            ndx = slotTable(ndx).nextItem
        Loop
    Next
        
    ' assign to the result
    Keys = res()
End Property

' the array of all values
' (VB5 users: convert return type to Variant)

Property Get Values() As Variant()
    Dim i As Long, ndx As Long
    Dim n As Long
    ReDim res(0 To pCount - 1) As Variant
    
    For i = 0 To pHashSize - 1
        ' take the pointer from the hash table
        ndx = hashTbl(i)
        ' walk the slottable() array
        Do While ndx
            res(n) = slotTable(ndx).Value
            n = n + 1
            ndx = slotTable(ndx).nextItem
        Loop
    Next
        
    ' assign to the result
    Values = res()
End Property


'*********************
'* PROTECTED METHODS *
'*********************
'initialize the hash table
Friend Sub protSetSize(ByVal HashSize As Long, Optional ByVal ListSize As Long, Optional ByVal ChunkSize As Long)
    ' provide defaults
    If ListSize <= 0 Then ListSize = pListSize
    If ChunkSize <= 0 Then ChunkSize = pChunkSize
    ' save size values
    pHashSize = HashSize
    pListSize = ListSize
    pChunkSize = ChunkSize
    pCount = 0
    ' rebuild tables
    FreeNdx = 0
    ReDim hashTbl(0 To HashSize - 1) As Long
    ReDim slotTable(0) As Bucket
    ExpandSlotTable pListSize
End Sub

Friend Sub protInit(Optional ByVal HashSize as Long = DEFAULT_HASHSIZE, Optional ByVal ListSize as Long = DEFAULT_LISTSIZE, Optional ByVal ChunkSize as Long = DEFAULT_CHUNKSIZE)
  ' initialize the tables at default size
  Call protSetSize(HashSize, ListSize, ChunkSize)
End Sub

'*******************
'* Private methods *
'*******************

' expand the slotTable() array
Private Sub ExpandSlotTable(ByVal numEls As Long)
    Dim newFreeNdx As Long, i As Long
    newFreeNdx = UBound(slotTable) + 1
    
    ReDim Preserve slotTable(0 To UBound(slotTable) + numEls) As Bucket
    ' create the linked list of free items
    For i = newFreeNdx To UBound(slotTable)
        slotTable(i).nextItem = i + 1
    Next
    ' overwrite the last (wrong) value
    slotTable(UBound(slotTable)).nextItem = FreeNdx
    ' we now know where to pick the first free item
    FreeNdx = newFreeNdx
End Sub



'Given some data, obtain the hash of the data supplied. Keys in this case can be of any type:
'String, LongLong, Double, Float, Currency or Object. If an object is supplied then it is tested whether
'stdIHashable is implemented. If it isn't then the object's pointer is used instead. The Hash produced by this 
'function varies based on pHashSize. I.E. if pHashSize is 3 then only numbers between 0 and 2 are produced.
'However if pHashSize is 10 then only numbers between 0 and 9 will be produced.
'@param {ByRef Variant} Key to be hashed
'@returns {Long} Hashed value
Private Function HashFunction(Key as variant) as Long
  Dim iRet as long: iRet = 0
  select case VarType(Key)
    Case vbString
      'Implementation taken from Java:
      'Returns a hash code for this string. The hash code for a String is computed as `s[0]*31^(n-1) + s[1]*31^(n-2) + ... + s[n-1]`
      'using int arithmetic, where `s[i]` is the `i`th character of the string, `n` is the length of the string, and ^ indicates exponentiation.
      '(The hash balue of the empty string is zero.)
      Dim b() as Byte: b = Key
      Dim i as long: for i = 0 to ubound(b)
        'Note: Diversion from Java algorithm, VBA is limited to 32-bit long in most cases
        'therefore we do the following weird modulo stuff.
        iRet = -31 * (((iRet > 69273666) * iRet Mod 507) + (iRet <= 69273666) * iRet) + b(i)
      next
    #if VBA7 then
    case vbLongLong
      HashFunction = HashFunction(CStr(key))
    #end if
    case vbDouble, vbSingle, vbCurrency
      'There appears to be no way (without declares) to copy the bytes of a double into a VBA byte array. Instead we go with a string, which isn't ideal but hopefully
      'will be good enough
      HashFunction = HashFunction(Cstr(Key))
    case vbDate
      HashFunction = HashFunction(Cstr(CDbl(key)))
    case vbInteger, vbLong, vbBoolean
      HashFunction = Abs(Key)
    case vbEmpty, vbNull, vbVariant
      HashFunction = 0
    case vbError
      HashFunction = Key.Number
    case vbObject
      if typeof key is stdIHashable then
        HashFunction = HashFunction(Key.hash)
      else
        HashFunction = HashFunction(objptr(Key))
      end if
  end select

  'Ensure iRet is within pHashSize
  HashFunction = iRet mod pHashSize
End Function




' get the index where an item is stored or 0 if not found
' if Create = True the item is created
'
' on exit Create=True only if a slot has been actually created
Private Function GetSlotIndex(ByVal Key As String, Optional Create As Boolean, _
    Optional HCode As Long, Optional LastNdx As Long) As Long
    Dim ndx As Long
    
    ' raise error if invalid key
    If Len(Key) = 0 Then Err.Raise 1001, , "Invalid key"
    
    ' keep case-unsensitiveness into account
    If pIgnoreCase Then Key = UCase$(Key)
    ' get the index in the hashTbl() array
    HCode = HashFunction(Key) Mod pHashSize
    ' get the pointer to the slotTable() array
    ndx = hashTbl(HCode)
    
    ' exit if there is no item with that hash code
    Do While ndx
        ' compare key with actual value
        If slotTable(ndx).Key = Key Then Exit Do
        ' remember last pointer
        LastNdx = ndx
        ' check the next item
        ndx = slotTable(ndx).nextItem
    Loop
    
    ' create a new item if not there
    If ndx = 0 And Create Then
        ndx = GetFreeSlot()
        PrepareSlot ndx, Key, HCode, LastNdx
    Else
        ' signal that no item has been created
        Create = False
    End If
    ' this is the return value
    GetSlotIndex = ndx

End Function

' return the first free slot

Private Function GetFreeSlot() As Long
    ' allocate new memory if necessary
    If FreeNdx = 0 Then ExpandSlotTable pChunkSize
    ' use the first slot
    GetFreeSlot = FreeNdx
    ' update the pointer to the first slot
    FreeNdx = slotTable(GetFreeSlot).nextItem
    ' signal this as the end of the linked list
    slotTable(GetFreeSlot).nextItem = 0
    ' we have one more item
    pCount = pCount + 1
End Function

' assign a key and value to a given slot

Private Sub PrepareSlot(ByVal Index As Long, ByVal Key As String, _
    ByVal HCode As Long, ByVal LastNdx As Long)
    ' assign the key
    ' keep case-sensitiveness into account
    If pIgnoreCase Then Key = UCase$(Key)
    slotTable(Index).Key = Key
    
    If LastNdx Then
        ' this is the successor of another slot
        slotTable(LastNdx).nextItem = Index
    Else
        ' this is the first slot for a given hash code
        hashTbl(HCode) = Index
    End If
End Sub