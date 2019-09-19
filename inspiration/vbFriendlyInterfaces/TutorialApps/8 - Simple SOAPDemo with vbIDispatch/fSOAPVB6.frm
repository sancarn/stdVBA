VERSION 5.00
Begin VB.Form fDemo 
   Caption         =   "VB6 SOAP Demo (using WinHttp 5.1 and MSXML LateBound, working from XP onwards)"
   ClientHeight    =   6765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   ScaleHeight     =   451
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   737
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame frmExamples 
      Caption         =   "Examples, which make use of the above Services in simple Forms"
      Height          =   855
      Left            =   180
      TabIndex        =   8
      Top             =   5760
      Width           =   10635
      Begin VB.CommandButton cmdShowCalcServ 
         Caption         =   "Show Calc-Service-Form"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   300
         Width           =   2355
      End
      Begin VB.CommandButton cmdShowCurrConv 
         Caption         =   "Show Currency-Converter-Form"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2700
         TabIndex        =   9
         Top             =   300
         Width           =   2415
      End
   End
   Begin VB.ListBox lstOutParams 
      Height          =   1230
      Left            =   2640
      TabIndex        =   6
      Top             =   4200
      Width           =   8175
   End
   Begin VB.ListBox lstInParams 
      Height          =   1230
      Left            =   2640
      TabIndex        =   4
      Top             =   2820
      Width           =   8175
   End
   Begin VB.ListBox lstMethods 
      Height          =   2010
      Left            =   2640
      TabIndex        =   3
      Top             =   660
      Width           =   8175
   End
   Begin VB.CommandButton cmdGetInterface 
      Caption         =   "Get Interface from WSDL -->"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   180
      Width           =   2355
   End
   Begin VB.ComboBox cmbWSDLs 
      Height          =   315
      ItemData        =   "fSOAPVB6.frx":0000
      Left            =   2640
      List            =   "fSOAPVB6.frx":000A
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   180
      Width           =   8175
   End
   Begin VB.Label lblOutParams 
      Alignment       =   1  'Rechts
      Caption         =   "Output-Parameters:"
      Height          =   255
      Left            =   660
      TabIndex        =   7
      Top             =   4230
      Width           =   1755
   End
   Begin VB.Label lblInParams 
      Alignment       =   1  'Rechts
      Caption         =   "Input-Parameters:"
      Height          =   255
      Left            =   780
      TabIndex        =   5
      Top             =   2850
      Width           =   1635
   End
   Begin VB.Label lblMethods 
      Alignment       =   1  'Rechts
      Caption         =   "SOAP-Methods:"
      Height          =   255
      Left            =   1020
      TabIndex        =   1
      Top             =   690
      Width           =   1395
   End
End
Attribute VB_Name = "fDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SOAP As New cSOAPInterface

Private Sub Form_Load()
  cmbWSDLs.ListIndex = 0
End Sub

Private Sub cmbWSDLs_Click()
  cmdGetInterface_Click
End Sub

Private Sub cmdGetInterface_Click()
Dim i As Long
On Error GoTo 1
  
  lstInParams.Clear
  lstOutParams.Clear
  lstMethods.Clear
  Refresh
  
  SOAP.InitAndParseWSDL cmbWSDLs.Text
  For i = 1 To SOAP.MethodCount
    lstMethods.AddItem SOAP.Method(i).Name
  Next i
  
  cmdShowCurrConv.Enabled = InStr(1, cmbWSDLs.Text, "CurrencyConvertor", 1) > 0 And SOAP.MethodCount > 0
  cmdShowCalcServ.Enabled = InStr(1, cmbWSDLs.Text, "CalcService", 1) > 0 And SOAP.MethodCount > 0
 
1 If Err Then MsgBox Err.Description
End Sub
 
Private Sub lstMethods_Click()
  FillParamList lstInParams, SOAP.Method(lstMethods.Text).InParams
  FillParamList lstOutParams, SOAP.Method(lstMethods.Text).OutParams
End Sub

Private Sub FillParamList(Lst As ListBox, Params As Collection)
Dim P As cSOAPParam
  Lst.Clear
  For Each P In Params
    Lst.AddItem "Name: " & P.Name & " (Type: " & P.TypeDef & _
                IIf(Len(P.MinOccurs), ", MinOccurs: " & P.MinOccurs, "") & _
                IIf(Len(P.MaxOccurs), ", MaxOccurs: " & P.MaxOccurs, "") & ")"
  Next
End Sub

Private Sub lstInParams_Click()
  If lstInParams.ListIndex = -1 Then Exit Sub Else lstOutParams.ListIndex = -1
  ShowExtendedTypeInfo SOAP.Method(lstMethods.Text).InParams(lstInParams.ListIndex + 1)
End Sub
Private Sub lstOutParams_Click()
  If lstOutParams.ListIndex = -1 Then Exit Sub Else lstInParams.ListIndex = -1
  ShowExtendedTypeInfo SOAP.Method(lstMethods.Text).OutParams(lstOutParams.ListIndex + 1)
End Sub
Private Sub ShowExtendedTypeInfo(P As cSOAPParam)
  fExtendedTypeInfo.ShowInfo P.ExtendedTypeInfo, Me
End Sub

Private Sub cmdShowCalcServ_Click()
  fDemoCalcService.InitAndShow SOAP
End Sub
Private Sub cmdShowCurrConv_Click()
  fDemoCurrencyConverter.InitAndShow SOAP
End Sub
 

