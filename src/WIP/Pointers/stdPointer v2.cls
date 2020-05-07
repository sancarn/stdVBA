'Pre requisites:

'Revolution in pointers:
'http://www.vbforums.com/showthread.php?886203-vb6-Getting-AddressOf-for-VB-Class-Object-Modules
'

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (src As Any, dst As Any, ByVal cbLen As Long)
Private Declare Function VirtualQuery Lib "kernel32.dll" (ByVal addr As Long, pMBI As Any, ByVal lenMBI As Long) As Long
Private Enum CodePageTypeEnumn
    cptUnknown = 0
    cptVbClass = 0
    cptVbDataReport = 1
    cptVbFormOrMDI = 2
    cptVbPropertyPage = 3
    cptVbUserControl = 4
End Enum









Private Enum iTypeEnum
  Unknown
  Integer
  String
  Single
  Float
  Double
  Currency
  Variant
  ArrayBase
  Object
  Function
  'Anything else...?
End Enum

Private iType as iTypeEnum
Private iPtrVal as LongPtr


Public Property Get Size(Optional piType as iTypeEnum = iTypeEnum.Unknown) as Integer
  if piType = iTypeEnum.Unknown then
    piType = iType
  end if
  
  select case piType
    case iTypeEnum.Object
      Size = 4
    'case ...
      'Size = ...
    case else
      '--> Raise Error
  end select
End Function


Public Function From(ByRef x as variant) as stdPointer
  'If typename(x) = ...
  
  'end if
end function
Public Function FromPtr(ByVal piPtrVal as LongPtr, piType as iTypeEnum) as stdPointer
  Dim ret as stdPointer
  set ret = new stdPointer
  ret.init(piPtrVal, piType)
end function


Public Sub init(ByVal piPtrVal as LongPtr, piType as iTypeEnum)
  if not initialised then
    iPtrVal = piPtrVal
    iType = piType
  else
    '--> Raise error
  end if
End Sub

Public Funcion DeRef() as variant
  if iType = iTypeEnum.Object
    set DeRef = DeRefAsObject()
  else
    '--> Raise error
  end if
End Function

'Public Function Invoke() ?





'Object deref example:
'  Note: object size is 4 bytes
'    Returns an object given its pointer
'    This function reverses the effect of the ObjPtr function
Public Function DeRefAsObject() As Object
    Dim obj As Object
    ' force the value of the pointer into the temporary object variable
    CopyMemory obj, iPtrVal, Size(iTypeEnum.Object)
    ' assign to the result (this increments the ref counter)
    Set DeRefAsObject = obj
    ' manually destroy the temporary object variable 
    ' (if you omit this step you'll get a GPF!)
    CopyMemory obj, 0&, Size(iTypeEnum.Object)
End Function



'   'Gets a pointer to the last private method of an object stored in memory
'   '@param obj Object to get the pointer form         
'   '@param [iScanLimit] Limit to scan the object to. In general if your class has less than 512 methods/fields this number will not need to be altered.
'   Public Function GetLastPrivateMethod(ByRef obj as object, optional byval iScanLimit as long = 512) as Long
'     Dim v as Long, n as Long, nAddr As Long
'     Dim b as Byte, m As Byte
'     
'     GetMem4 ObjPtr(Me), nAddr                       'get address of the form's vtable
'     nAddr = nAddr + &H6F8                           'bump to the user part of the form's vtable
'     GetMem4 nAddr, n                                'read the address of the first entry point
'     GetMem1 n, m                                    'read the jump opcode at the first entry point [&H33 for psuedo code, &HE9 for native code]
'     For v = 1 To iScanLimit                         'scan a number of vtable entries (specified by user, default is 512)
'       nAddr = nAddr + 4                             'next entry address
'       GetMem4 nAddr, n                              'read the address of the entry point
'       If IsBadCodePtr(n) Then GoTo vTableEnd        'is the entry point address valid code?
'       GetMem1 n, b                                  'read the jump opcode at the entry point
'       If b <> m Then GoTo vTableEnd                 'does the jump opcode match that of the first vtable entry?
'     Next v
'     Exit Function                                   'last vtable entry not found... increase the For limit?
'   vTableEnd:
'     GetMem4 nAddr - 4, nAddr                        'back one entry to the last private method
'     GetLastPrivateMethod = nAddr
'   End Function


Public Function GetAddressOfEx(ByRef VbCodePage As Object, ByVal nOrdinal As Long, _
                            Optional ByVal iCodePageType As Long = CodePageTypeEnumn.cptUnknown, _
                            Optional ByRef nMethodCount As Long, Optional ByRef nLastMethodOffset As Long) As Long
    'Intellisense
    Dim CodePageType as CodePageTypeEnumn
    CodePageType = iCodePageType

    ' Routine is basically an AddressOf function for VB code pages (forms, classes, etc)
    
    If nOrdinal < 1 Then Exit Function
    If VbCodePage Is Nothing Then Exit Function
    
    ' redesigned but based on Paul Caton's zAddressOf method that can find function
    '   pointers within VB forms, classes, etc. Redesign includes merging 2 routines,
    '   using known VB offsets, and use of VirtualQuery over IsBadCodePtr API. This
    '   revised logic is slower than Caton's latest versions, but will prevent
    '   unintended page-guard activation and will not fail to return results in
    '   cases where Caton's zAddressOf would due to his built-in loop restrictions.
    
    ' Modify the routine for large-address-aware application, as needed, i.e., pointer-safe math.
    
    ' Parameters
    '   :: VbCodePage is the VB class module (form, class, etc) containing the method ordinal
    '   :: nOrdinal is the ordinal whose function pointer/address is to be returned
    '       ordinals are always one-bound and counted from the bottom of the code page
    '       the last method is ordinal #1, second to last is #2, etc
    '       keep public methods near top of code page & private/friend near bottom because
    '       VB will move public ones closer to top during runtime, offsetting your ordinals.
    '   :: CodePageType when passed can help the function scan the code page more efficiently
    '   :: nMethodCount is returned with the number of user-defined methods in the code page
    '   :: nLastMethodOffset is returned with the address after the last user-defined method
    ' Return value
    '   If success, the function pointer will be returned, else zero is returned
    '   If zero is returned, nMethodCount and nLastMethodOffset may not be updated
    
    ' How this method works...
    ' With known offsets and function signatures, finding what we want is pretty easy.
    ' The function signature is simply the 1st byte of the function's code
    ' 1) If a function pointer is zero, then this is expected
    '       Seen typically when a code page uses Implements keyword
    ' Otherwise, there are four byte values we are interested in (the signature)
    ' 2) Byte &H33  start of XOR instruction in native code (always when in IDE)
    ' 3) Byte &HE9  start of XOR instruction in P-Code (only when compiled in P-Code)
    ' 4) Byte &H81  start of ADD instruction, regardless of P-Code usage
    ' 5) Byte &H58  start of POP instruction, regardless of P-Code usage
    
    Dim bSig As Byte, bVal As Byte
    Dim nAddr&, vOffset&, nFirst&
    Dim nMethod&, nAttempts&, n&
    Dim minAddrV&, maxAddrV&, minAddrM&, maxAddrM&
    Dim MBI&(0 To 6)          ' faux MEMORY_BASIC_INFORMATION structure
    ' (0) BaseAddress member    minimum range of committed memory (same protection)
    ' (3) Range member          maximum range BaseAddress+Range
    ' (5) Protect member
    ' This structure is key to not crashing while probing memory addresses.
    ' The Protect member of the structure is examined after each call. If it
    ' contains &H101 (mask), then the address is a page-guard or has no-access.
    ' Otherwise, if it contains &HFE (mask) then the address is readable.
    
    ' Step 1. Probe the passed code page to find the first user-defined method.
    ' The probe is quite fast. The outer For:Next loop helps to quickly filter the
    ' passed code page via the known offsets. The inner DO loop will execute up to
    ' four times to find the right code page offset as needed. After found, it will
    ' execute as little as one time or several times, depending on Implements usage
    ' and number of Public variables declared within the code page. That inner loop
    ' has a fudge-factor built in should some signature exist that is not known yet.
    ' However, no others have been found, to date, after the known offsets of the
    ' correct code page.
    
    If CodePageType <= cptUnknown Or CodePageType > cptVbUserControl Then
        n = 0: nAttempts = 4
    Else
        n = CodePageType: nAttempts = n
    End If
    CopyMemory nAddr, ByVal ObjPtr(VbCodePage), 4 ' host VTable
    
    For n = n To nAttempts                      ' search in ascending order of offsets
        Select Case n
            Case 0: vOffset = nAddr + &H1C      ' known offset for VB Class,DataEnvironment,Add-in,DHTMLPage
            Case 1: vOffset = nAddr + &H9C      ' known offset for VB DataReport
            Case 2: vOffset = nAddr + &H6F8     ' known offset for VB Form, MDI
            Case 3: vOffset = nAddr + &H710     ' known offset for VB Property Page
            Case 4: vOffset = nAddr + &H7A4     ' known offset for VB UserControl
        End Select
        
        nAttempts = 4                           ' fudge-factor
        Do
            ' First validate the VTable slot address. If invalid, unsupported code page type
            If vOffset < minAddrV Or vOffset > maxAddrV Then
                MBI(5) = 0: VirtualQuery vOffset, MBI(0), 28
                If (MBI(5) And &HFE) = 0 Or (MBI(5) And &H101) <> 0 Then Exit For
                minAddrV = MBI(0): maxAddrV = minAddrV + MBI(3)  ' set min/max range
            End If
            CopyMemory nMethod, ByVal vOffset, 4 ' get function address at VTable slot
            If nMethod <> 0 Then                ' zero = implemented, skip
            
                ' Next validate the function pointer. If invalid, unsupported code page type
                If nMethod < minAddrM Or nMethod > maxAddrM Then
                    MBI(5) = 0: VirtualQuery nMethod, MBI(0), 28
                    If (MBI(5) And &HFE) = 0 Or (MBI(5) And &H101) <> 0 Then Exit For
                    minAddrM = MBI(0): maxAddrM = minAddrM + MBI(3)  ' set min/max range
                End If
            
                CopyMemory bVal, ByVal nMethod, 1 ' get the 1st byte of that method
                If bVal = &H33 Or bVal = &HE9 Then
                    nFirst = vOffset            ' cache the location of first user-defined method
                    bSig = bVal: Exit For       ' cache the function signature & done
                ElseIf bVal <> &H81 Then        ' if not one of these 4 signatures, decrement attempts
                    If bVal <> &H58 Then nAttempts = nAttempts - 1
                End If
            End If
            vOffset = vOffset + 4               ' look at next VTable slot
        Loop Until nAttempts = 0
    Next
    
    If nFirst = 0 Then Exit Function           ' failure
    ' If failure, then likely one of two reasons:
    ' 1) Unsupported code page
    ' 2) Code page has no user-defined methods
    
    ' Step 2. Find the last user-defined method.
    ' VB stacks user-defined methods contiguously, back to back. So, to find the last method,
    ' we simply need to keep looking until a signature no longer matches or we hit end of page.
    Do
        ' Validate the VTable slot address. If invalid, end of code page & done
        vOffset = vOffset + 4
        If vOffset < minAddrV Or vOffset > maxAddrV Then
            MBI(5) = 0: VirtualQuery vOffset, MBI(0), 28
            If (MBI(5) And &HFE) = 0 Or (MBI(5) And &H101) <> 0 Then Exit Do
            minAddrV = MBI(0): maxAddrV = minAddrV + MBI(3)  ' set min/max range
        End If
        
        CopyMemory nMethod, ByVal vOffset, 4    ' get function pointer at VTable slot
        If nMethod = 0 Then Exit Do             ' if zero, done because doesn't match our signature
        
        ' Validate the function pointer. If invalid, end of code page & done
        If nMethod < minAddrM Or nMethod > maxAddrM Then
            MBI(5) = 0: VirtualQuery nMethod, MBI(0), 28
            If (MBI(5) And &HFE) = 0 Or (MBI(5) And &H101) <> 0 Then Exit Do
            minAddrM = MBI(0): maxAddrM = minAddrM + MBI(3)  ' set min/max range
        End If
        
        CopyMemory bVal, ByVal nMethod, 1       ' get function's signature
    Loop Until bVal <> bSig                     ' done when doesn't match our signature
    
    ' Now set the optional parameter values
    nMethodCount = (vOffset - nFirst) \ 4
    nLastMethodOffset = vOffset
    
    ' Return the function pointer for requested ordinal, if a valid ordinal
    If nOrdinal <= nMethodCount Then
        CopyMemory GetAddressOfEx, ByVal vOffset - (nOrdinal * 4), 4
    End If

End Function
