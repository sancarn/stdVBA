VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbIEnumVariant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
Public Function NextElm(Index As Long, ByVal cElements As Long, VariantArrayFirstElement As Variant, ElementsFetched As Long) As HRESULT
End Function

Public Function Skip(Index As Long, ByVal cElements As Long) As HRESULT
End Function

Public Function Reset(Index As Long) As HRESULT
End Function
