VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAddressClassic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Child-Class of cAddressClassicFactory: the usual VB6-approach - just for comparison (of coding-efforts and performance)

Private Type tm
  ID As Long
  Name As String
  BirthDay As String
End Type

Private m As tm

Public Property Get ID() As Long
  ID = m.ID
End Property
Public Property Let ID(ByVal RHS As Long)
  m.ID = RHS
End Property

Public Property Get Name() As String
  Name = m.Name
End Property
Public Property Let Name(RHS As String)
  m.Name = RHS
End Property

Public Property Get BirthDay() As Date
  BirthDay = m.BirthDay
End Property
Public Property Let BirthDay(RHS As Date)
  m.BirthDay = RHS
End Property

Public Function BirthDayToday() As Boolean
  Dim Today As Date
      Today = Date
  With m
    BirthDayToday = Day(.BirthDay) = Day(Today) And Month(.BirthDay) = Month(Today)
  End With
End Function
