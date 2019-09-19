VERSION 5.00
Begin VB.Form fDemoCurrencyConverter 
   Caption         =   "Get Currency-Exchange-Rates"
   ClientHeight    =   1410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   ScaleHeight     =   1410
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   315
      Left            =   4620
      TabIndex        =   6
      Top             =   900
      Width           =   1215
   End
   Begin VB.TextBox txtRate 
      Height          =   315
      Left            =   4620
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.ComboBox cmbTo 
      Height          =   315
      ItemData        =   "fDemoCurrencyConverter.frx":0000
      Left            =   2220
      List            =   "fDemoCurrencyConverter.frx":01C9
      Sorted          =   -1  'True
      Style           =   2  'Dropdown-Liste
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.ComboBox cmbFrom 
      Height          =   315
      ItemData        =   "fDemoCurrencyConverter.frx":04C0
      Left            =   660
      List            =   "fDemoCurrencyConverter.frx":0689
      Sorted          =   -1  'True
      Style           =   2  'Dropdown-Liste
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "-> current Rate:"
      Height          =   195
      Left            =   3420
      TabIndex        =   5
      Top             =   420
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "To:"
      Height          =   195
      Left            =   1920
      TabIndex        =   4
      Top             =   420
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "From:"
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   420
      Width           =   495
   End
End
Attribute VB_Name = "fDemoCurrencyConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSOAP As cSOAPInterface

Friend Sub InitAndShow(SOAP As cSOAPInterface)
  Set mSOAP = SOAP
  'set the ComboBoxes to some initial string-Enum-Values, according to the WSDL-Param-description
  cmbFrom.Text = "GBP" 'Great-Britain-Pound
  cmbTo.Text = "EUR" 'Euro
  Show vbModal
End Sub

'GUI-related triggering...
Private Sub cmbFrom_Click()
  txtRate.Text = ConversionRate(cmbFrom.Text, cmbTo.Text)
End Sub
Private Sub cmbTo_Click()
  txtRate.Text = ConversionRate(cmbFrom.Text, cmbTo.Text)
End Sub
Private Sub cmdRefresh_Click()
  txtRate.Text = ConversionRate(cmbFrom.Text, cmbTo.Text)
End Sub

'...and a Function-Wrapper for the SOAP-based Remote-Call
Public Function ConversionRate(strEnumFrom As String, strEnumTo As String) As Double
  If Len(strEnumFrom) <> 3 Or Len(strEnumTo) <> 3 Then 'a simple check on the Input-Arguments
    ConversionRate = -1 'the -1 signifying: "no conversion-rate available"
  Else
    ConversionRate = Val(mSOAP.Execute.ConversionRate(strEnumFrom, strEnumTo).Text)
  End If
End Function

 
