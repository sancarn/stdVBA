' //
' // IDispatch implementation of light-weight COM object
' // By The trick 2018
' //

Option Explicit

Private Const E_INVALIDARG            As Long = &H80070057
Private Const E_NOINTERFACE           As Long = &H80004002

Private Const CC_STDCALL              As Long = 4
Private Const VT_BYREF                As Long = &H4000&

Private Const DISPATCH_PROPERTYPUT    As Long = 4
Private Const DISPATCH_METHOD         As Long = 1
Private Const DISPATCH_PROPERTYGET    As Long = 2
Private Const LOCALE_SYSTEM_DEFAULT   As Long = &H800
Private Const HEAP_ZERO_MEMORY        As Long = &H8
Private Const FADF_AUTO               As Long = 1

Private Type SAFEARRAYBOUND
    cElements   As Long
    lLBound     As Long
End Type

Private Type SAFEARRAY
    cDims       As Integer
    fFeatures   As Integer
    cbElements  As Long
    cLocks      As Long
    pvData      As Long
    Bounds      As SAFEARRAYBOUND
End Type

Private Type UUID
    Data1       As Long
    Data2       As Integer
    Data3       As Integer
    Data4(7)    As Byte
End Type

Private Type PARAMDATA
    pszName     As String
    vt          As Integer
End Type

Private Type METHODDATA
    pszName     As String
    ppdata      As Long
    dispid      As Long
    iMeth       As Long
    cc          As Long
    cArgs       As Long
    wFlags      As Integer
    vtReturn    As Integer
End Type

Private Type INTERFACEDATA
    pmethdata   As Long
    cMembers    As Long
End Type

Private Enum eIUnknownMethods
    METH_IUNKNOWN_QI
    METH_IUNKNOWN_ADDREF
    METH_IUNKNOWN_RELEASE
    METH_IUNKNOWN_COUNT
End Enum

Private Enum eISumDiffMethods
    METH_ISUMDIFF_SETVAL
    METH_ISUMDIFF_GETVAL
    METH_ISUMDIFF_SUM
    METH_ISUMDIFF_DIFF
    METH_ISUMDIFF_COUNT
End Enum

Private Type tIUnknownVTable
    pfn(METH_IUNKNOWN_COUNT - 1)    As Long
End Type

Private Type tISumDiffVtable
    tUnk                            As tIUnknownVTable
    pfn(METH_ISUMDIFF_COUNT - 1)    As Long
End Type

Private Type CBaseClass
    pVtbl                           As Long
    lRefCounter                     As Long
    pDisp                           As Long            ' // Inner object
End Type

' // CSumDiff class
Private Type CSumDiff
    CBase                           As CBaseClass
    lVal1                           As Long
    lVal2                           As Long
End Type

Private Const SIZEOF_CSUMDIFF       As Long = &H14

Private Declare Function CreateDispTypeInfo Lib "OleAut32" ( _
                         ByRef pidata As INTERFACEDATA, _
                         ByVal lcid As Long, _
                         ByRef pptinfo As IUnknown) As Long
Private Declare Function CreateStdDispatch Lib "OleAut32" ( _
                         ByRef punkOuter As Any, _
                         ByRef pvThis As Any, _
                         ByVal ptinfo As IUnknown, _
                         ByRef ppunkStdDisp As Any) As Long
Private Declare Function IsEqualGUID Lib "ole32" ( _
                         ByRef rguid1 As UUID, _
                         ByRef rguid2 As UUID) As Long
Private Declare Function vbaObjSetAddref Lib "MSVBVM60.DLL" _
                         Alias "__vbaObjSetAddref" ( _
                         ByRef dstObject As Any, _
                         ByRef srcObjPtr As Any) As Long
Private Declare Function vbaObjSet Lib "MSVBVM60.DLL" _
                         Alias "__vbaObjSet" ( _
                         ByRef dstObject As Any, _
                         ByRef srcObjPtr As Any) As Long
Private Declare Function vbaCastObj Lib "MSVBVM60.DLL" _
                         Alias "__vbaCastObj" ( _
                         ByRef dstObject As Any, _
                         ByRef pIID As UUID) As Long
Private Declare Function HeapAlloc Lib "kernel32" ( _
                         ByVal hHeap As Long, _
                         ByVal dwFlags As Long, _
                         ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32" ( _
                         ByVal hHeap As Long, _
                         ByVal dwFlags As Long, _
                         ByVal lpMem As Long) As Long
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function GetMem4 Lib "msvbvm60" ( _
                         ByRef src As Any, _
                         ByRef dst As Any) As Long
Private Declare Sub MoveArray Lib "msvbvm60" _
                    Alias "__vbaAryMove" ( _
                    ByRef Destination() As Any, _
                    ByRef Source As Any)
Private Declare Sub IIDFromString Lib "ole32" ( _
                    ByVal lpsz As Long, _
                    ByRef lpiid As UUID)

Sub Main()
    Dim cObj    As Object
    
    ' // Restrictions:
    ' // You should destroy the object before ending
    
    Set cObj = CreateSumDiffObject(10, 20)
    
    cObj(0) = 54
    cObj(1) = 20
    
    Debug.Print cObj.Sum
    Debug.Print cObj.Diff
    
    cObj.Value(0) = 1231
    cObj.Value(1) = 2000
    
    Debug.Print cObj(0)
    Debug.Print cObj.Value(1)
    
    Set cObj = Nothing

End Sub
   
' // Create SumDiff object
Public Function CreateSumDiffObject( _
                ByVal lVal1 As Long, _
                ByVal lVal2 As Long) As IUnknown
    Static tInterfaceInfo   As INTERFACEDATA    ' // Global data
    Static cTypeInfo        As IUnknown
    
    Dim tObj()      As CSumDiff
    Dim pObj        As Long
    Dim tArrDesk    As SAFEARRAY
    Dim hr          As Long
    Dim cUnkDisp    As IUnknown
    
    ' // Get the interface data and create TypeInfo
    If tInterfaceInfo.pmethdata = 0 Then
        
        tInterfaceInfo = CSumDiff_InterfaceData()
        
        hr = CreateDispTypeInfo(tInterfaceInfo, LOCALE_SYSTEM_DEFAULT, cTypeInfo)
        If hr < 0 Then Exit Function
    
    End If
    
    ' // Alloc memory for the object and map it to the array item
    pObj = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, SIZEOF_CSUMDIFF)
    
    tArrDesk.cDims = 1
    tArrDesk.Bounds.cElements = 1
    tArrDesk.fFeatures = FADF_AUTO          ' // Don't free memory after destroying of array
    tArrDesk.pvData = pObj
    tArrDesk.cbElements = SIZEOF_CSUMDIFF
    
    MoveArray tObj, VarPtr(tArrDesk)
    
    ' // Call constructor
    If CSumDiff_CSumDiff(tObj(0)) = 0 Then
        HeapFree GetProcessHeap(), 0, ByVal pObj
        Exit Function
    End If
    
    ' // Initial values
    tObj(0).lVal1 = lVal1
    tObj(0).lVal2 = lVal2
    
    ' // Create aggregated IDispatch
    hr = CreateStdDispatch(tObj(0), tObj(0), cTypeInfo, tObj(0).CBase.pDisp)
    
    If hr < 0 Then
        HeapFree GetProcessHeap(), 0, ByVal pObj
        Exit Function
    End If

    vbaObjSetAddref CreateSumDiffObject, ByVal pObj
    
    Debug.Print "Object CSumDiff created at 0x" & Hex$(pObj)
    
End Function

Private Function FAR_PROC( _
                 ByVal lValue As Long) As Long
    Dim bIsInIDE    As Boolean
    
    ' // AddressOf statement actually returns the pointer to the small thunk which checks if code is running or not
    ' // When you see an object variable in the Watch window the code is stopped. We need to avoid that behavior
    ' // because some interfaces returns the HRESULT value and it'll cause unexpected behavior (return S_OK)
    
    Debug.Assert MakeTrue(bIsInIDE)
    
    If bIsInIDE Then
        GetMem4 ByVal lValue + &H16, FAR_PROC   ' // Skip thunk
    Else
        FAR_PROC = lValue
    End If
    
End Function

Private Function MakeTrue( _
                 ByRef bValue As Boolean) As Boolean
    MakeTrue = True
    bValue = True
End Function

' // Base class constructor
Private Function CBaseClass_CBaseClass( _
                 ByRef tObj As CBaseClass) As Long
    Static tVtable  As tIUnknownVTable
    
    tObj.lRefCounter = 0
    tObj.pVtbl = VarPtr(tVtable)
    
    If tVtable.pfn(METH_IUNKNOWN_QI) = 0 Then
        
        tVtable.pfn(METH_IUNKNOWN_QI) = FAR_PROC(AddressOf CBaseClass_QueryInterface)
        tVtable.pfn(METH_IUNKNOWN_ADDREF) = FAR_PROC(AddressOf CBaseClass_AddRef)
        tVtable.pfn(METH_IUNKNOWN_RELEASE) = FAR_PROC(AddressOf CBaseClass_Release)
        
    End If
    
    CBaseClass_CBaseClass = 1
    
End Function

' // Base class methods
Private Function CBaseClass_QueryInterface( _
                 ByRef tObj As CBaseClass, _
                 ByRef tiid As UUID, _
                 ByRef pOut As Long) As Long
    Static tIUnk    As UUID, tIDisp     As UUID
    
    If tIUnk.Data1 = 0 Then
        
        IIDFromString StrPtr("{00000000-0000-0000-C000-000000000046}"), tIUnk
        IIDFromString StrPtr("{00020400-0000-0000-C000-000000000046}"), tIDisp

    End If
        
    Select Case True
    Case IsEqualGUID(tiid, tIUnk)
        
        pOut = VarPtr(tObj)
        
        CBaseClass_AddRef tObj
        
    Case IsEqualGUID(tiid, tIDisp)
        
        ' // Return aggregable object
        pOut = vbaCastObj(ByVal tObj.pDisp, tIDisp)

    Case Else
    
        pOut = 0:   CBaseClass_QueryInterface = E_NOINTERFACE
        Exit Function
        
    End Select
    
End Function

Private Function CBaseClass_AddRef( _
                 ByRef tObj As CBaseClass) As Long
    
    tObj.lRefCounter = tObj.lRefCounter + 1
    CBaseClass_AddRef = tObj.lRefCounter
    
End Function

Private Function CBaseClass_Release( _
                 ByRef tObj As CBaseClass) As Long
    
    tObj.lRefCounter = tObj.lRefCounter - 1
    CBaseClass_Release = tObj.lRefCounter
    
    If CBaseClass_Release = 0 Then
        
        vbaObjSet tObj.pDisp, ByVal 0&
        
        ' // Destructor
        HeapFree GetProcessHeap(), 0, VarPtr(tObj)
        
        Debug.Print "Object was destroyed at 0x" & Hex$(VarPtr(tObj))
        
    End If
    
End Function

' // CSumDiff constructor
Private Function CSumDiff_CSumDiff( _
                 ByRef tObj As CSumDiff) As Long
    Static tVtable  As tISumDiffVtable
    
    If CBaseClass_CBaseClass(tObj.CBase) = 0 Then Exit Function
    
    tObj.CBase.pVtbl = VarPtr(tVtable)
    tObj.lVal1 = 0
    tObj.lVal2 = 0
    
    If tVtable.pfn(METH_ISUMDIFF_SETVAL) = 0 Then
        
        tVtable.tUnk.pfn(METH_IUNKNOWN_QI) = FAR_PROC(AddressOf CBaseClass_QueryInterface)
        tVtable.tUnk.pfn(METH_IUNKNOWN_ADDREF) = FAR_PROC(AddressOf CBaseClass_AddRef)
        tVtable.tUnk.pfn(METH_IUNKNOWN_RELEASE) = FAR_PROC(AddressOf CBaseClass_Release)
        tVtable.pfn(METH_ISUMDIFF_SETVAL) = FAR_PROC(AddressOf CSumDiff_SetVal)
        tVtable.pfn(METH_ISUMDIFF_GETVAL) = FAR_PROC(AddressOf CSumDiff_GetVal)
        tVtable.pfn(METH_ISUMDIFF_SUM) = FAR_PROC(AddressOf CSumDiff_Sum)
        tVtable.pfn(METH_ISUMDIFF_DIFF) = FAR_PROC(AddressOf CSumDiff_Diff)
        
    End If
                     
    CSumDiff_CSumDiff = 1
                     
End Function

' // CSumDiff methods
Private Function CSumDiff_SetVal( _
                 ByRef tObj As CSumDiff, _
                 ByVal lIndex As Long, _
                 ByVal lValue As Long) As Long
    
    Select Case lIndex
    Case 0:     tObj.lVal1 = lValue
    Case 1:     tObj.lVal2 = lValue
    Case Else:  CSumDiff_SetVal = E_INVALIDARG
    End Select
    
End Function

Private Function CSumDiff_GetVal( _
                 ByRef tObj As CSumDiff, _
                 ByVal lIndex As Long) As Long
    
    Select Case lIndex
    Case 0:     CSumDiff_GetVal = tObj.lVal1
    Case 1:     CSumDiff_GetVal = tObj.lVal2
    Case Else:  Err.Raise 5
    End Select
    
End Function

' // Calculate sum
Private Function CSumDiff_Sum( _
                 ByRef tObj As CSumDiff) As Long
    
    CSumDiff_Sum = tObj.lVal1 + tObj.lVal2
    
End Function

' // Calculate difference
Private Function CSumDiff_Diff( _
                 ByRef tObj As CSumDiff) As Long
    
    CSumDiff_Diff = tObj.lVal1 - tObj.lVal2
    
End Function

' // CSumDiff interface data
Private Function CSumDiff_InterfaceData() As INTERFACEDATA
    Static tData        As INTERFACEDATA
    Static tMembers()   As METHODDATA
    Static tParams0()   As PARAMDATA
    Static tParams1()   As PARAMDATA
    Static tParams2()   As PARAMDATA
    
    If tData.pmethdata = 0 Then
        
        ReDim tMembers(METH_ISUMDIFF_COUNT - 1)

        tData.cMembers = METH_ISUMDIFF_COUNT
        tData.pmethdata = VarPtr(tMembers(0))
        
        ' // [DEFAULT] HRESULT CSumDiff.Value (Byval Index As Long, Byval Value as Long)
        tParams0 = SetParamData("Index", vbLong, "Value", vbLong)
        tMembers(METH_ISUMDIFF_SETVAL) = SetMethodData("Value", 0, 3, vbError, DISPATCH_PROPERTYPUT, tParams0())
        
        ' // [DEFAULT] Long CSumDiff.Value (Byval Index As Long)
        tParams1 = SetParamData("Index", vbLong)
        tMembers(METH_ISUMDIFF_GETVAL) = SetMethodData("Value", 0, 4, vbLong, DISPATCH_PROPERTYGET Or DISPATCH_METHOD, tParams1())
    
        ' // Long CSumDiff.Sum ()
        tMembers(METH_ISUMDIFF_SUM) = SetMethodData("Sum", 3, 5, vbLong, DISPATCH_METHOD, tParams2())
        
        ' // Long CSumDiff.Diff ()
        tMembers(METH_ISUMDIFF_DIFF) = SetMethodData("Diff", 4, 6, vbLong, DISPATCH_METHOD, tParams2())
        
    End If
    
    CSumDiff_InterfaceData = tData
    
End Function

' // Fast creation
Private Function SetParamData( _
                 ParamArray vData() As Variant) As PARAMDATA()
    Dim vIndex  As Variant
    Dim lIndex  As Long
    Dim tRet()  As PARAMDATA
    
    ReDim tRet(UBound(vData) \ 2)
    
    For Each vIndex In vData
        
        If VarType(vIndex) = vbString Then
            tRet(lIndex).pszName = vIndex
        Else
            tRet(lIndex).vt = vIndex
            lIndex = lIndex + 1
        End If
        
    Next
    
    SetParamData = tRet
    
End Function

' // Fast creation
Private Function SetMethodData( _
                 ByRef sName As String, _
                 ByVal lDispID As Long, _
                 ByVal lMethIndex As Long, _
                 ByVal lRetValType As Long, _
                 ByVal lFlags As Long, _
                 ByRef tParams() As PARAMDATA) As METHODDATA
    
    With SetMethodData
        
        .cc = CC_STDCALL
        .dispid = lDispID
        .iMeth = lMethIndex
        .pszName = sName
        .vtReturn = lRetValType
        .wFlags = lFlags
        
        If Not Not tParams Then
            
            .ppdata = VarPtr(tParams(0))
            .cArgs = UBound(tParams) + 1
            
        End If
        
    End With
    
End Function