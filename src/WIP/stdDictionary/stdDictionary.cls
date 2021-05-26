VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Hash size typically we are anticipating only 10 hashes being required. In reality the larger this number is the better in terms of
'performance, however the more memory is consumed. If you're generating hundreds of dictionaries you want to keep this number as low as possible.
'on the other hand if you are making 1 huge hash, then this number should likely be a lot larger.
Private Const DEFAULT_HASHSIZE = 10 

'The number of slots expected to be required typically. Typically this would be the typical number of items we would expect to be
'added to our dictionary.
'Note: You can override this capacity using `CreateAdvanced()` function.
Private Const DEFAULT_CAPACITY = 10

'Max 
Private Const MAX_CAPACITY = 2147483647

Private Const DEFAULT_CHUNKSIZE = 1024

Private Type Bucket
    bIsRoot as boolean
    vKey As Variant
    vValue As Variant
    iNextItem As Long      ' 0 if last item
End Type

Private pHashes() as Long
Private pBuckets() as Bucket
Private pHashSize as long
Private pListSize as Long
Private pIndex as Long




Public Function Create(ParamArray keyItems() as variant) as stdDictionary
  set Create = new stdDictionary
  Call Create.protInit()
  For i = 0 to ubound(keyItems)
    Call Create.Add(keyItems(i),keyItems(i+1))
  next
End Function
Public Function CreateAdvanced(HashSize, ListSize,ChunkSize,ParamArray keyItems() as variant) as stdDictionary
  set Create = new stdDictionary
  Call Create.protInit()
  For i = 0 to ubound(keyItems)
    Call Create.Add(keyItems(i),keyItems(i+1))
  next
End Function

'Property Get/Set/Let Item
Attribute Item.VB_UserMemId = 0
Public Property Get Item(ByVal vKey as variant) as Variant
  Dim index as Long: index = HashFunction(vKey)
  Dim b as Bucket: b = pBuckets(index)
  Dim bContinue as Boolean: bContinue = true
  Dim bFound as boolean: bFound = false
  While bContinue
    if testEqual(b.vKey,vKey) then
      bContinue = false
      bFound = true
    else
      if b.iNextItem then
        b = pBuckets(b.iNextItem)
      else
        bContinue = false
      end if
    end if
  Wend

  If bFound then
    if isObject(b.vValue) then
      set Item = b.vValue
    else
      let Item = b.vValue
    end if
  else
    Item = Empty
  end if
End Function
Public Property Let Item(ByVal vKey as variant, ByVal vValue as variant)
  
End Function
Public Property Set Item(ByVal vKey as variant, ByVal vValue as variant)

End Function

Public Sub Add(ByVal vKey as Variant, ByVal vValue as variant)
  '
End Sub
Public Sub Remove(ByVal vKey as Variant)
  '
End Sub
Public Sub RemoveAll()
  '
End Sub


'Initialise the hash
Friend Sub protInit(Optional ByVal HashSize as Long = DEFAULT_HASHSIZE, Optional ByVal ListSize as Long = DEFAULT_LISTSIZE, Optional ByVal ChunkSize as Long = DEFAULT_CHUNKSIZE)
  Redim pHashes(1 to HashSize)
  Redim pBuckets(1 to ListSize)
  pHashSize = HashSize
  pListSize = ListSize
  pIndex = 0
End Sub

'Ensure the buckets array is big enough to take additional elements / to save memory.
'@bIncrement {Boolean} True if incrementing size (add), False if decrementing size (remove)
Private Sub EnsureCapacity(Byval bIncrement as boolean)
  Dim iMinCapacityReq as Long: iMinCapacityReq = pIndex + iif(bIncrement,1,-1)
  if pListSize = MAX_CAPACITY then
    Err.Raise 1, "stdDictionary::EnsureCapacity", "Out of memory. Cannot expand array"
    Exit Sub
  end if
  if iMinCapacityReq - pListSize > 0 then
    Dim iNewCapacity as Long
    if MAX_CAPACITY - pListSize - pListSize \ 2 > 0 then
      iNewCapacity = pListSize + pListSize \ 2
    else
      iNewCapacity = MAX_CAPACITY
    end if
    pListSize = iNewCapacity 
    Redim Preserve pBuckets(1 to iNewCapacity)
  end if 
End Sub

'Test whether 2 variants are equal. If they are return true.
'@param {ByRef Variant} Value to be compared with    (typically a key)
'@param {ByRef Variant} Value to be compared against (typically a key)
'@return {Boolean} If equal return true
Private Function testEqual(ByRef v1 as variant, ByRef v2 as variant) as Boolean
  if vartype(v1) = vartype(v2) then
    if isObject(v1) then
      if typeof v1 is stdIHashable then
        Dim h as stdIHashable: set h = v1
        testEqual = h.isEqual(v2)
      else
        testEqual = v1 is v2
      end if
    else
      testEqual = v1 = v2
    end if
  end if
End Function

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
      iRet = HashFunction(CStr(key))
    #end if
    case vbDouble, vbSingle, vbCurrency
      'There appears to be no way (without declares) to copy the bytes of a double into a VBA byte array. Instead we go with a string, which isn't ideal but hopefully
      'will be good enough
      iRet = HashFunction(Cstr(Key))
    case vbDate
      iRet = HashFunction(Cstr(CDbl(key)))
    case vbInteger, vbLong, vbBoolean
      iRet = Abs(Key)
    case vbEmpty, vbNull, vbVariant
      iRet = 0
    case vbError
      iRet = Key.Number
    case vbObject
      if typeof key is stdIHashable then
        Dim hKey as stdIHashable: set hKey = key
        iRet = HashFunction(hKey.hash)
      else
        iRet = HashFunction(objptr(Key))
      end if
  end select

  'Ensure iRet is within pHashSize
  HashFunction = iRet mod pHashSize
End Function