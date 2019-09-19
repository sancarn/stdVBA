VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSOAPMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String, ServiceURL As String, TargetNamespace As String, SoapAction As String
Public InSequenceName As String, InParams As New Collection
Public OutSequenceName As String, OutParams As New Collection

Public Function Request(ParamArray P()) As Object
Dim sCont As String, Pref As String
    sCont = "<?xml version='1.0' encoding='utf-8'?><soap:Envelope"
    sCont = sCont & " xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'"
    sCont = sCont & " xmlns:xsd='http://www.w3.org/2001/XMLSchema'"
    sCont = sCont & " xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"
    sCont = sCont & "<soap:Body>" & P2Xml(P) & "</soap:Body></soap:Envelope>"
  
  With CreateObject("WinHttp.WinHttpRequest.5.1")
    .Open "POST", ServiceURL, False
    If Len(SoapAction) Then .setRequestHeader "SOAPAction", SoapAction
    .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    .send GetUTF8BytesFromString(sCont)
    If .Status <> 200 Then Err.Raise vbObjectError, , .statusText
    
    Set Request = CreateObject("MSXML2.DOMDocument")
        Request.loadXML .responseText
        Pref = Split(Request.XML, OutSequenceName)(0)
        If Right(Pref, 1) = ":" Then Pref = Mid(Pref, InStrRev(Pref, "<") + 1) Else Pref = ""
    Set Request = Request.getElementsByTagName(Pref & OutSequenceName)(0).firstChild
  End With
End Function

Private Function P2Xml(ByVal P) As String
Dim i As Long, sValue As String, MethodNode As Object, DOM As Object
  Set DOM = CreateObject("MSXML2.DOMDocument")
  Set MethodNode = DOM.createNode(1, Name, TargetNamespace)
    For i = 0 To UBound(P)
      Select Case VarType(P(i))
        Case vbDouble, vbSingle, vbCurrency: sValue = Replace(P(i), ",", ".")
        Case Else:                           sValue = P(i)
      End Select
      MethodNode.appendChild(DOM.createNode(1, InParams(i + 1).Name, TargetNamespace)).Text = sValue
    Next i
  P2Xml = MethodNode.XML
End Function
 
Function GetUTF8BytesFromString(S)
  With CreateObject("ADODB.Stream")
    .Type = 2 'text
    .Charset = "utf-8"
    .Open
    .WriteText S, 0
    .Position = 0
    .Type = 1 'binary
    .Position = 3 'skip the UTF8-BOM
    GetUTF8BytesFromString = .Read
  End With
End Function

