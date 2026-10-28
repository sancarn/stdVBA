VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Scratch-project adapter for the same Test.Assert API used by the full suite.
Public checks As Long
Public lastCheck As String

Public Sub Topic(ByVal name As String)
End Sub

Public Sub Assert(ByVal name As String, ByVal passed As Boolean)
  checks = checks + 1
  lastCheck = name
  If Not passed Then Err.Raise vbObjectError + 4000, "stdRegexTests", name
End Sub
