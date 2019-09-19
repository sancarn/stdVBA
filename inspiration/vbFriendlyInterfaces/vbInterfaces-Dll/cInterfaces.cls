VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cInterfaces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum HRESULT
  S_OK = 0
  S_FALSE = 1
  E_NOTIMPL = &H80004001
  E_NOINTERFACE = &H80004002
  E_POINTER = &H80004003
  E_ABORT = &H80004004
  E_FAIL = &H80004005
  E_ACCESSDENIED = &H80070005
  E_HANDLE = &H80070006
  E_OUTOFMEMORY = &H8007000E
  E_INVALIDARG = &H80070057
  E_UNEXPECTED = &H8000FFFF
 
  DISP_E_MEMBERNOTFOUND = &H80020003
  DISP_E_UNKNOWNNAME = &H80020006
  DISP_E_NONAMEDARGS = &H80020007
  DISP_E_EXCEPTION = &H80020009
  DISP_E_BADINDEX = &H8002000B
End Enum

Public Enum eAvailableVTables
  vtbl_IUnknown
  vtbl_IDispatch
  vtbl_IEnumVariant
  vtbl_IPicture
End Enum

Public Sub NewInstance(pVTableToUse As Long, ImplementingCallbackObj As vbIUnknown, pVarPtrNewInstance As Long, _
                       Optional ByVal UserData As Long = -1, Optional ByVal pExtUserData As Long, Optional ByVal ExtByteLen As Long)
  If pVTableToUse = 0 Then Err.Raise vbObjectError, , "pVTableToUse can't be a Null-Pointer"
  If pVarPtrNewInstance = 0 Then Err.Raise vbObjectError, , "pVarPtrNewInstance can't be a Null-Pointer"
  If ImplementingCallbackObj Is Nothing Then Err.Raise vbObjectError, , "we need a valid Implementing-Instance to call-back-to"
  CreateBaseInstance pVTableToUse, ImplementingCallbackObj, pVarPtrNewInstance, UserData, pExtUserData, ExtByteLen
End Sub

Public Sub AddTo(Methods() As Long, ByVal Addr As Long)
  Dim i As Long
  Do: i = i + 1: Loop While Methods(i)
  Methods(i) = Addr
End Sub
 
Public Sub CopyMethods(SrcMethods() As Long, pVTableDst As Long)
  If pVTableDst = 0 Then Err.Raise vbObjectError, , "pVTableDst can't be a Null-Pointer"
  Dim MethodEntries As Long
      MethodEntries = UBound(SrcMethods) - LBound(SrcMethods) + 1
  Assign ByVal pVTableDst, SrcMethods(LBound(SrcMethods)), MethodEntries * 4
End Sub

'this is thought for direct Construction of Instances which don't try to be more than what's (currently) implemented in this lib - and listed below
Public Property Get pVT(VTableType As eAvailableVTables) As Long
  Select Case VTableType
    Case vtbl_IUnknown:     pVT = vtIUnknown.pVTable
    Case vtbl_IDispatch:    pVT = vtIDispatch.pVTable
    Case vtbl_IEnumVariant: pVT = vtIEnumVariant.pVTable
    Case vtbl_IPicture:     pVT = vtIPicture.pVTable
  End Select
End Property

'the next two Props are for inheriting the Method-Addresses of the builtin VTables into outside defined, other VTables of concrete Types
Public Property Get vtIUnknownMethods() As Long()
  vtIUnknownMethods = vtIUnknown.Methods
End Property
Public Property Get vtIDispatchMethods() As Long()
  vtIDispatchMethods = vtIDispatch.Methods
End Property

Public Property Get sIID_IUnknown() As String
  sIID_IUnknown = "{00000000-0000-0000-C000-000000000046}"
End Property
Public Property Get sIID_IDispatch() As String
  sIID_IDispatch = "{00020400-0000-0000-C000-000000000046}"
End Property
Public Property Get sIID_IEnumVariant() As String
  sIID_IEnumVariant = "{00020404-0000-0000-C000-000000000046}"
End Property
Public Property Get sIID_IPicture() As String
  sIID_IPicture = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"
End Property

Public Sub VariantCopyToPtr(ByVal pDst As Long, Src As Variant)
  VariantCopyToPtrAPI pDst, Src
End Sub

Public Sub MemCopyPtr(ByVal pDst As Long, ByVal pSrc As Long, ByVal Bytes As Long)
  Assign ByVal pDst, ByVal pSrc, Bytes
End Sub

Public Function IIDsEqual(sIID1 As String, sIID2 As String) As Boolean
  IIDsEqual = StrComp(sIID1, sIID2, vbTextCompare) = 0
End Function

Public Function IIDInList(sIIDtoCompareToList As String, ParamArray IIDList()) As Boolean
  Dim i As Long
  For i = 0 To UBound(IIDList)
    If StrComp(sIIDtoCompareToList, IIDList(i), vbTextCompare) = 0 Then IIDInList = True: Exit For
  Next
End Function

