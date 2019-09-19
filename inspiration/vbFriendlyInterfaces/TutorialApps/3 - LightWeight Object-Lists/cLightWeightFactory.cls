VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAddressLWeightFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'other than the lightweight-factory in Tutorial-Folder #2, the implementation here will Add (and hand out)
'only earlybound-instances, it acts more as a Parent-Class-Container for vbIAddress-instances (a List-Class)
'This class (in conjunction with the vbIAddress-implementation in vtIAddress.bas) offers the same functionality
'(only faster - especially in case of instance-removal) as the combination of: cAddressClassicFactory and cAddressClassic

Implements vbIUnknown
Implements vbIDispatch

Private sIID_IAddress As String, m As tmAddress, mZero As tmAddress
Private mCol As New Collection '<- we'll keep our amount of lightweight instances in a normal VB-Collection
 
Private Sub Class_Initialize()
  Dim Unk As stdole.IUnknown, ConcreteInstance As vbIAddress
  vbI.NewInstance vtIAddress.pVTable, Me, VarPtr(Unk)
  Set ConcreteInstance = Unk 'and now force Query-Interface the very first time (to fill our sIID_IAddress-Variable first)
End Sub
 
'the first of 5 Public methods in our "LightWeight-Factory" (handing out new lightweight-instances 'As vbIAddress' + adding it into our internal Collection)
Public Function Add(ByVal ID As Long, Name As String, LastName As String, ByVal BirthDay As Date) As vbIAddress
  m.ID = ID
  m.Name = Name
  m.BirthDay = BirthDay
  vbI.NewInstance vtIAddress.pVTable, Me, VarPtr(Add), , VarPtr(m), LenB(m)
  
  mCol.Add Add
End Function

Public Property Get Item(Key) As vbIAddress
  Set Item = mCol.Item(Key)
End Property

Public Property Get Count() As Long
  Count = mCol.Count
End Property

Public Function Enumerator() As stdole.IUnknown
Attribute Enumerator.VB_UserMemId = -4
  Set Enumerator = mCol.[_NewEnum]
End Function

Public Sub RemoveAll()
  Set mCol = Nothing
End Sub
'************* End of the Public Factory-Interface *************


'************* IUnknown-Implementation *****************
Private Sub vbIUnknown_QueryInterface(UserData As Long, ByVal pVTable As Long, RefCount As Long, sReqIID As String, Unk As stdole.IUnknown)
  If Len(sIID_IAddress) = 0 Then sIID_IAddress = sReqIID 'only done once, to gather the IID for early-binding to vbIAddress

  If vbI.IIDsEqual(sReqIID, sIID_IAddress) Then RefCount = RefCount + 1 'we only support Early-Binding here (no IDispatch or 'As Object' allowed on the outside)
End Sub

Private Sub vbIUnknown_Terminate(UserData As Long, ByVal pVTable As Long)
  If pVTable = vtIAddress.pVTable And UserData <> -1 Then 'it's a vbIAddress-Instance-Type with extended Data
    vbI.MemCopyPtr VarPtr(m), UserData, LenB(m)
    m = mZero 'let VB cleanup the member-content of 'm' for us, by applying an Empty struct of the same UDT
  End If
End Sub
 
 
'************* IDispatch-Implementation ****************
Private Function vbIDispatch_GetIDForMemberName(UserData As Long, ByVal pVTable As Long, MemberName As String) As Long
End Function
 
Private Function vbIDispatch_Invoke(UserData As Long, ByVal pVTable As Long, ByVal DispID As Long, ByVal CallType As VbCallType, VResult As Variant, ParamArray P() As Variant) As HRESULT
  vbIDispatch_Invoke = DISP_E_MEMBERNOTFOUND '<- let's return the appropriate HRESULT, in case someone tries
End Function
 
