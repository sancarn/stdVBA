VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbIAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'a small Interface-Definition (this time 3 Properties and one method)
'the 3 Properties will expand into 2 methods each (Get- and Let-methods, in exactly this order)

Public ID As Long
Public Name As String
Public BirthDay As Date
Public Function BirthDayToday() As Boolean: End Function
 
