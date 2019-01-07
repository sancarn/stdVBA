VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "STD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'TODO:
'*****************************************************************************************
'STD_Types_Callback
'- Full implementation of Pointer
'STD_Automation_COM:
' - Make Predeclared id
' - Fix STD__NSAutomation accordingly
' - Implement raw COM functions and behaviour
'STD_Runtimes_CLR:
' - Make Predeclared id
' - Fix STD__NSRuntimes accordingly
'STD_Types_Regex
' - Implement. Currently implementation is found as a function in STD_Types_Date
'STD_Types_Enum
' - Full Implementation.
'*****************************************************************************************
'Interesting Finds:
'*****************************************************************************************
'Generally advanced VBA:
'  http://exceldevelopmentplatform.blogspot.com/
'DispIDs:
'  https://docs.microsoft.com/en-us/previous-versions/windows/embedded/ms864423(v%3Dmsdn.10)
'Inline Assembly and [] syntax:
'  http://www.vbforums.com/showthread.php?862105-FYI-How-to-impl-DISPID_EVALUATE-in-VB6
'Execute OPCodes (Machine code) in VB6:
'  http://www.freevbcode.com/ShowCode.asp?ID=1863
'VB6 Attributes:
'  https://stackoverflow.com/questions/33648764/what-does-the-attribute-keyword-do-in-vb6
'Analysis of type library
'  http://www.vbforums.com/showthread.php?800045-Undocumented-Function-using-VBA6-Dll&p=4902755&viewfull=1#post4902755    || or http://www.vbforums.com/showthread.php?799593
'Serialise COM object:
'  http://exceldevelopmentplatform.blogspot.com/2018/12/vba-persistence-use-lset-to-serialise.html


Public Automation As STD__NSAutomation
Public Runtimes As STD__NSRuntimes
Public Types As STD__NSTypes


Private Sub Class_Initialize()
  Set Automation = New STD__NSAutomation
  Set Runtimes = New STD__NSRuntimes
  Set Types = New STD__NSTypes
End Sub
