VERSION 5.00
Begin VB.Form fDemoCalcService 
   BackColor       =   &H00FFFFFF&
   Caption         =   "A few MethodCalls over the SOAP-CalcService"
   ClientHeight    =   5235
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9210
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows-Standard
End
Attribute VB_Name = "fDemoCalcService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private mSOAP As cSOAPInterface

Friend Sub InitAndShow(SOAP As cSOAPInterface)
  Set mSOAP = SOAP
  
  AutoRedraw = True
  Print "Click Me!"
  
  Show vbModal
End Sub

Private Sub Form_Click()
  Cls
  
  Print vbLf; " Result of HelloWorld(""VB6-cSOAPInterface""):"
  Print "    "; mSOAP.Execute.HelloWorld("VB6-cSOAPInterface").Text
 
  Print vbLf; " Result of AddInteger(1, 2):"
  Print "    "; mSOAP.Execute.AddInteger(1, 2).Text
  
  Print vbLf; " Result of AddDouble(1.11, 2.22):"
  Print "    "; mSOAP.Execute.AddDouble(1.11, 2.22).Text
 
  Print vbLf; " Result of CalcPrimeFactors(""1234567890""):"
  Print "    "; mSOAP.Execute.CalcPrimeFactors("1234567890").Text

  Print vbLf; " Result of CalcPrimeFactors2(1234567890):"
  Print "    "; mSOAP.Execute.CalcPrimeFactors2(1234567890).Text
  
  Refresh 'just to put out the intermediate results fast, since the next call will deliberately take 1sec (at the ServerSide)
  
  Print vbLf; " Result of SlowWorld(1):"
  Print "    "; mSOAP.Execute.SlowWorld(1).Text
End Sub

