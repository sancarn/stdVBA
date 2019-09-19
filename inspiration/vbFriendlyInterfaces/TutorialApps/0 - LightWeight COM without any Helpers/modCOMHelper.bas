Attribute VB_Name = "modCOMHelper"
Option Explicit

Declare Function CoTaskMemAlloc& Lib "ole32" (ByVal sz&)
Declare Sub CoTaskMemFree Lib "ole32" (ByVal pMem&)
Declare Sub Assign Lib "kernel32" Alias "RtlMoveMemory" (Dst As Any, Src As Any, Optional ByVal CB& = 4)
 
Function FuncPtr(ByVal Addr As Long) As Long 'just a small Helper for the AddressOf KeyWord
  FuncPtr = Addr
End Function

