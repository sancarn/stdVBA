'
' MEnumerator.bas
'
' Implementation of IEnumVARIANT to support For Each in VB6
'
' Original source:  http://www.vbforums.com/showthread.php?854963-VB6-IEnumVARIANT-For-Each-support-without-a-typelib
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type TENUMERATOR
    VTablePtr   As Long
    References  As Long
    Enumerable  As Object
    Index       As Long
    Upper       As Long
    Lower       As Long
End Type

Private Enum API
    NULL_ = 0
    S_OK = 0
    S_FALSE = 1
    E_NOTIMPL = &H80004001
    E_NOINTERFACE = &H80004002
    E_POINTER = &H80004003
#If False Then
    Dim NULL_, S_OK, S_FALSE, E_NOTIMPL, E_NOINTERFACE, E_POINTER
#End If
End Enum

Private Declare Function FncPtr Lib "msvbvm60" Alias "VarPtr" (ByVal Address As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60" (Src As Any, Dst As Any) As Long
Private Declare Function CopyBytesZero Lib "msvbvm60" Alias "__vbaCopyBytesZero" (ByVal Length As Long, Dst As Any, Src As Any) As Long
Private Declare Function CoTaskMemAlloc Lib "ole32" (ByVal cb As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Function IIDFromString Lib "ole32" (ByVal lpsz As Long, ByVal lpiid As Long) As Long
Private Declare Function SysAllocStringByteLen Lib "oleaut32" (ByVal psz As Long, ByVal cblen As Long) As Long
Private Declare Function VariantCopyToPtr Lib "oleaut32" Alias "VariantCopy" (ByVal pvargDest As Long, ByRef pvargSrc As Variant) As Long
Private Declare Function InterlockedIncrement Lib "kernel32" (ByRef Addend As Long) As Long
Private Declare Function InterlockedDecrement Lib "kernel32" (ByRef Addend As Long) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewEnumerator(ByRef Enumerable As Object, _
                              ByVal Upper As Long, _
                              Optional ByVal Lower As Long _
                              ) As IEnumVARIANT
' Class Factory
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Static VTable(6) As Long
    If VTable(0) = NULL_ Then
        ' Setup the COM object's virtual table
        VTable(0) = FncPtr(AddressOf IUnknown_QueryInterface)
        VTable(1) = FncPtr(AddressOf IUnknown_AddRef)
        VTable(2) = FncPtr(AddressOf IUnknown_Release)
        VTable(3) = FncPtr(AddressOf IEnumVARIANT_Next)
        VTable(4) = FncPtr(AddressOf IEnumVARIANT_Skip)
        VTable(5) = FncPtr(AddressOf IEnumVARIANT_Reset)
        VTable(6) = FncPtr(AddressOf IEnumVARIANT_Clone)
    End If
    
    Dim This As TENUMERATOR
    With This
        ' Setup the COM object
        .VTablePtr = VarPtr(VTable(0))
        .References = 1
        Set .Enumerable = Enumerable
        .Lower = Lower
        .Index = Lower
        .Upper = Upper
    End With
    
    ' Allocate a spot for it on the heap
    Dim pThis As Long
    pThis = CoTaskMemAlloc(LenB(This))
    If pThis Then
        ' CopyBytesZero is used to zero out the original
        ' .Enumerable reference, so that VB doesn't mess up the
        ' reference count, and free our enumerator out from under us
        CopyBytesZero LenB(This), ByVal pThis, This
        DeRef(VarPtr(NewEnumerator)) = pThis
    End If
End Function

Private Function RefToIID$(ByVal riid As Long)
    ' copies an IID referenced into a binary string
    Const IID_CB As Long = 16&  ' GUID/IID size in bytes
    DeRef(VarPtr(RefToIID)) = SysAllocStringByteLen(riid, IID_CB)
End Function

Private Function StrToIID$(ByRef iid As String)
    ' converts a string to an IID
    StrToIID = RefToIID$(NULL_)
    IIDFromString StrPtr(iid), StrPtr(StrToIID)
End Function

Private Function IID_IUnknown() As String
    Static iid As String
    If StrPtr(iid) = NULL_ Then _
        iid = StrToIID$("{00000000-0000-0000-C000-000000000046}")
    IID_IUnknown = iid
End Function

Private Function IID_IEnumVARIANT() As String
    Static iid As String
    If StrPtr(iid) = NULL_ Then _
        iid = StrToIID$("{00020404-0000-0000-C000-000000000046}")
    IID_IEnumVARIANT = iid
End Function

Private Function IUnknown_QueryInterface(ByRef This As TENUMERATOR, _
                                         ByVal riid As Long, _
                                         ByVal ppvObject As Long _
                                         ) As Long
    If ppvObject = NULL_ Then
        IUnknown_QueryInterface = E_POINTER
        Exit Function
    End If

    Select Case RefToIID$(riid)
        Case IID_IUnknown, IID_IEnumVARIANT
            DeRef(ppvObject) = VarPtr(This)
            IUnknown_AddRef This
            IUnknown_QueryInterface = S_OK
        Case Else
            IUnknown_QueryInterface = E_NOINTERFACE
    End Select
End Function

Private Function IUnknown_AddRef(ByRef This As TENUMERATOR) As Long
    IUnknown_AddRef = InterlockedIncrement(This.References)
End Function

Private Function IUnknown_Release(ByRef This As TENUMERATOR) As Long
    IUnknown_Release = InterlockedDecrement(This.References)
    If IUnknown_Release = 0& Then
        Set This.Enumerable = Nothing
        CoTaskMemFree VarPtr(This)
    End If
End Function

Private Function IEnumVARIANT_Next(ByRef This As TENUMERATOR, _
                                   ByVal celt As Long, _
                                   ByVal rgVar As Long, _
                                   ByRef pceltFetched As Long _
                                   ) As Long
    
    Const VARIANT_CB As Long = 16 ' VARIANT size in bytes
    
    If rgVar = NULL_ Then
        IEnumVARIANT_Next = E_POINTER
        Exit Function
    End If
    
    Dim Fetched As Long
    With This
        Do Until .Index > .Upper
            VariantCopyToPtr rgVar, .Enumerable(.Index)
            .Index = .Index + 1&
            Fetched = Fetched + 1&
            If Fetched = celt Then Exit Do
            rgVar = PtrAdd(rgVar, VARIANT_CB)
        Loop
    End With
    
    If VarPtr(pceltFetched) Then pceltFetched = Fetched
    If Fetched < celt Then IEnumVARIANT_Next = S_FALSE
End Function

Private Function IEnumVARIANT_Skip(ByRef This As TENUMERATOR, ByVal celt As Long) As Long
    IEnumVARIANT_Skip = E_NOTIMPL
End Function

Private Function IEnumVARIANT_Reset(ByRef This As TENUMERATOR) As Long
    IEnumVARIANT_Reset = E_NOTIMPL
End Function

Private Function IEnumVARIANT_Clone(ByRef This As TENUMERATOR, ByVal ppEnum As Long) As Long
    IEnumVARIANT_Clone = E_NOTIMPL
End Function

Private Function PtrAdd(ByVal Pointer As Long, ByVal Offset As Long) As Long
    Const SIGN_BIT As Long = &H80000000
    PtrAdd = (Pointer Xor SIGN_BIT) + Offset Xor SIGN_BIT
End Function

Private Property Let DeRef(ByVal Address As Long, ByVal Value As Long)
    GetMem4 Value, ByVal Address
End Property
