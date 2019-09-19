VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbIDispatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GetIDForMemberName(UserData As Long, ByVal pVTable As Long, MemberName As String) As Long
End Function

Public Function Invoke(UserData As Long, ByVal pVTable As Long, ByVal DispID As Long, ByVal CallType As VbCallType, VResult, ParamArray P()) As HRESULT
End Function
