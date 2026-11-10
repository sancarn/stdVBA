VERSION 1.0 CLASS
BEGIN
  MultiUse = -1 'True
END
Attribute VB_Name = "stdOutlookSyncFake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Test-only late-bound stand-in. Never attaches to the real Outlook application.
Public WindowCount As Long
Public Displays As Long
Public Requests As Long
Public ShowProgress As Boolean
Public FailSync As Boolean
Public FailInitialize As Boolean
Public InitializeCalls As Long
Public Order As String

Public Function GetNamespace(ByVal name As String) As Object
  If name <> "MAPI" Then Err.Raise 5, , "Wrong namespace"
  Order = Order & "Namespace;"
  Set GetNamespace = Me
End Function

Public Function GetDefaultFolder(ByVal kind As Long) As Object
  If kind <> 6 Then Err.Raise 5, , "Wrong initialization folder"
  InitializeCalls = InitializeCalls + 1
  If FailInitialize Then Err.Raise 5, , "Synthetic initialization failure"
  Order = Order & "Initialize;"
  Set GetDefaultFolder = Me
End Function

Public Property Get DefaultStore() As Object
  Order = Order & "Store;"
  Set DefaultStore = Me
End Property

Public Property Get StoreID() As String
  StoreID = "63000000"
End Property

Public Property Get Explorers() As Object
  Set Explorers = Me
End Property

Public Property Get Count() As Long
  Count = WindowCount
End Property

Public Function GetExplorer() As Object
  Set GetExplorer = Me
End Function

Public Sub Display()
  Displays = Displays + 1
  WindowCount = WindowCount + 1
  Order = Order & "Display;"
End Sub

Public Sub SendAndReceive(ByVal showProgressDialog As Boolean)
  Requests = Requests + 1
  ShowProgress = showProgressDialog
  Order = Order & "Sync;"
  stdOutlookSendCalendarTests.CheckSyncOrder
  If FailSync Then Err.Raise 5, , "Synthetic synchronization failure"
End Sub
