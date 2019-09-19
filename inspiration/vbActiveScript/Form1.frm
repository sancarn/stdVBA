VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "This is the caption of my form"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'--> MAKE SURE YOU REGISTER THE TLB FILE <--
'
'If you down know how..google for "tlb reggie register" reggie is a free
'tlb registration program you can find on the net..adds right click option
'for it.
'
'
'This is a VB only implementation of IActiveScript which lets you integrate
'scripting support in your apps without the need for the MSScript control.
'
'I started working on this because I wanted to work torwards adding full
'IActiveScript support including debugging because the MS Script control
'sometimes isnt enough.
'
'Anyway, this is the first step. It may not be 100% perfectly implemented
'but it runs and works with objects you pass to it.
'
'You are free to use it in any commercial or non commercial applications
'as you so desire. Only a brief line "This product contains software written by
'David Zimmer" is required.
'
'Enjoy
'
'http://sandsprite.com
'
'ps - if you want to see a C activex control that you can use from VB with
'debugging support check out ken fousts citrus debugger
'
'http://sandsprite.com/CodeStuff/CitrusDebugger.7z
'
'For the brave of heart, I also started to try implementing the iActiveScript
'debug interfaces directly in VB. You can get a copy here:
'
'http://sandsprite.com/CodeStuff/vbActiveScript_wDbg_incomplete.zip
'
'Currently I have ditched the MS Script engine completely and am now using
'DukTape javascript engine 1mb self contained complete with debug support.
'search my site or web for duk4vb.


Dim WithEvents scriptControl As clsActiveScriptSite
Attribute scriptControl.VB_VarHelpID = -1


Private Sub Form_Load()
    Set scriptControl = New clsActiveScriptSite
    
    Dim myScript As String
    
    myScript = "frm.visible = true" & vbCrLf & _
               "msgbox frm.caption" & vbCrLf & _
               "mTextbox.text = ""This text set through scripting""" & vbCrLf & _
               "frm.left =0: frm.top=0" & vbCrLf & _
               "myString = ""This is my string you see!""" & vbCrLf & _
               "function test(): test=""sample function return"":end function"
               
    With scriptControl
        .AddObject "frm", Me
        .AddObject "mTextbox", Text1
        .RunCode myScript
        MsgBox .Eval("myString")
        MsgBox .Eval("10+10")
        MsgBox .Eval("test()")
    End With
    
End Sub

Private Sub scriptControl_Error(pscripterror As Long)
    MsgBox "Script error number: " & pscripterror
End Sub
