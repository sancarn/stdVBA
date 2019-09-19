VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSOAPInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbIUnknown
Implements vbIDispatch
 
Public ServiceURL As String, TargetNamespace As String, PortName As String

Private vbI As New vbInterfaces.cInterfaces, mCol As New Collection
Private DOM As Object, TypeList As Object, TypeListImp As Object, N As Object

Public Function InitAndParseWSDL(ByVal URL As String)
  Set mCol = Nothing 'clear the existing Methods-Collection
  Set DOM = GetDOMFromURL(URL)

  ServiceURL = GetAttr(DOM.getElementsByTagName("soap:address")(0), "location")
  PortName = GetAttr(DOM.getElementsByTagName("soap:binding")(0).parentNode, "type", True)
  TargetNamespace = GetAttr(DOM.documentElement, "targetNamespace")
 
  For Each N In DOM.documentElement.childNodes
    If N.baseName = "types" Then
      Set TypeList = N.firstChild.childNodes
      If N.firstChild.firstChild.baseName = "import" Then
        Set TypeListImp = GetDOMFromURL(GetAttr(N.firstChild.firstChild, "schemaLocation")).documentElement.childNodes
      End If
    End If
  Next N
 
  For Each N In DOM.getElementsByTagName("soap:operation")
    AddNewMethodFrom N
  Next
End Function

Private Function GetDOMFromURL(ByVal URL As String)
  If StrComp(Left(URL, 5), "http:", 1) Then URL = Left(ServiceURL, InStr(8, ServiceURL, "/")) & URL
  With CreateObject("WinHttp.WinHttpRequest.5.1")
       .open "GET", Trim$(URL), False
       .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
       .send
       Set GetDOMFromURL = CreateObject("MSXML2.DOMDocument")
           GetDOMFromURL.loadXML .responseText
   End With
End Function

Private Sub AddNewMethodFrom(N As Object)
Dim M As New cSOAPMethod, InMsg As String, OutMsg As String
    M.TargetNamespace = TargetNamespace
    M.ServiceURL = ServiceURL
    M.Name = GetAttr(N.parentNode, "name")
    M.SoapAction = GetAttr(N, "soapAction")
  
    For Each N In DOM.documentElement.childNodes
      If N.baseName = "portType" Then If GetAttr(N, "name") = PortName Then Exit For
    Next N
    For Each N In N.childNodes
      If N.baseName = "operation" Then If GetAttr(N, "name") = M.Name Then Exit For
    Next N
    For Each N In N.childNodes
      If N.baseName = "input" Then InMsg = GetAttr(N, "message", True)
      If N.baseName = "output" Then OutMsg = GetAttr(N, "message", True)
    Next N
    For Each N In DOM.documentElement.childNodes
      If N.baseName = "message" Then
        If GetAttr(N, "name") = InMsg Then M.InSequenceName = GetAttr(N.firstChild, "element", True)
        If GetAttr(N, "name") = OutMsg Then M.OutSequenceName = GetAttr(N.firstChild, "element", True)
      End If
    Next N
    ParseMethodParams M.InSequenceName, M.InParams
    ParseMethodParams M.OutSequenceName, M.OutParams
    
    mCol.Add M, M.Name
End Sub

Private Sub ParseMethodParams(ByVal Name As String, Col As Collection)
Dim Param As cSOAPParam
  For Each N In TypeList
    If N.baseName = "element" And GetAttr(N, "name") = Name And N.hasChildNodes Then
      If N.firstChild.baseName = "complexType" And N.firstChild.hasChildNodes Then Set N = N.firstChild: Exit For
    ElseIf N.baseName = "complexType" And GetAttr(N, "name") = Name Then
      If N.hasChildNodes Then Exit For
    ElseIf N.baseName = "element" And GetAttr(N, "name") = Name And Len(GetAttr(N, "type")) > 0 Then
      AddParamFromNode N, New cSOAPParam, Col: Exit Sub
    End If
  Next N
  If N Is Nothing And Not TypeListImp Is Nothing Then
    For Each N In TypeListImp
      If N.baseName = "element" And GetAttr(N, "name") = Name And N.hasChildNodes Then
        If N.firstChild.baseName = "complexType" And N.firstChild.hasChildNodes Then Set N = N.firstChild: Exit For
      ElseIf N.baseName = "complexType" And GetAttr(N, "name") = Name Then
        If N.hasChildNodes Then Exit For
      End If
    Next N
  End If
  If N Is Nothing Then Exit Sub
  If N.firstChild.hasChildNodes And N.baseName <> "element" Then Set N = N.firstChild
  For Each N In N.childNodes
    If N.baseName = "element" Then AddParamFromNode N, New cSOAPParam, Col
  Next
End Sub

Private Sub AddParamFromNode(N As Object, Param As cSOAPParam, Col As Collection)
  Param.Name = GetAttr(N, "name")
  Param.TypeDef = GetAttr(N, "type", True)
  Param.MinOccurs = GetAttr(N, "minOccurs")
  Param.MaxOccurs = GetAttr(N, "maxOccurs")
  Col.Add Param, Param.Name
  
  For Each N In TypeList
    If GetAttr(N, "name") = Param.TypeDef Then Exit For
  Next
  If N Is Nothing And Not TypeListImp Is Nothing Then
    For Each N In TypeListImp
      If GetAttr(N, "name") = Param.TypeDef Then Exit For
    Next
  End If
  If Not N Is Nothing Then Param.ExtendedTypeInfo = Replace(N.XML, vbTab, "    ")
End Sub

Private Function GetAttr(N, AttrName As String, Optional ByVal NoPrefix As Boolean) As String
  On Error GoTo 1
     GetAttr = N.Attributes.getNamedItem(AttrName).Text
1 If NoPrefix Then GetAttr = Mid$(GetAttr, InStr(GetAttr, ":") + 1)
End Function

Public Property Get MethodCount() As Long
  MethodCount = mCol.Count
End Property

Public Property Get MethodExists(MethodName As String) As Boolean
On Error GoTo NoSuchMethod
  Call mCol.Item(MethodName)
  MethodExists = True
NoSuchMethod:
End Property

Public Property Get Method(MethodNameOrOneBasedIndex) As cSOAPMethod
  Set Method = mCol(MethodNameOrOneBasedIndex)
End Property

Public Function Execute() As Object 'our IDispatch-supporting DispObject, which allows LateBound Method-Calling "per Dot"
  vbI.NewInstance vbI.pVT(vtbl_IDispatch), Me, VarPtr(Execute)
End Function
 
'************* IUnknown-Implementation *****************
Private Sub vbIUnknown_QueryInterface(UserData As Long, ByVal pVTable As Long, RefCount As Long, sReqIID As String, Unk As stdole.IUnknown)
  If vbI.IIDsEqual(sReqIID, vbI.sIID_IDispatch) Then RefCount = RefCount + 1
End Sub

Private Sub vbIUnknown_Terminate(UserData As Long, ByVal pVTable As Long)
End Sub


'************* IDispatch-Implementation ****************
Private Function vbIDispatch_GetIDForMemberName(UserData As Long, ByVal pVTable As Long, MemberName As String) As Long
  Dim i As Long
  For i = 1 To mCol.Count
    If StrComp(mCol(i).Name, MemberName, 1) = 0 Then vbIDispatch_GetIDForMemberName = i: Exit For
  Next
End Function

Private Function vbIDispatch_Invoke(UserData As Long, ByVal pVTable As Long, ByVal DispID As Long, ByVal CallType As VbCallType, VResult As Variant, ParamArray P() As Variant) As vbInterfaces.HRESULT
  If DispID < 1 Or DispID > mCol.Count Then vbIDispatch_Invoke = DISP_E_MEMBERNOTFOUND: Exit Function
  
  If CallType And (VbGet Or VbMethod) Then '
    Select Case UBound(P) + 1
      Case 0: Set VResult = Method(DispID).Request()
      Case 1: Set VResult = Method(DispID).Request(P(0))
      Case 2: Set VResult = Method(DispID).Request(P(0), P(1))
      Case 3: Set VResult = Method(DispID).Request(P(0), P(1), P(2))
      Case 4: Set VResult = Method(DispID).Request(P(0), P(1), P(2), P(3))
      Case 5: Set VResult = Method(DispID).Request(P(0), P(1), P(2), P(3), P(4))
      Case 6: Set VResult = Method(DispID).Request(P(0), P(1), P(2), P(3), P(4), P(5))
      Case 7: Set VResult = Method(DispID).Request(P(0), P(1), P(2), P(3), P(4), P(5), P(6))
      Case 8: Set VResult = Method(DispID).Request(P(0), P(1), P(2), P(3), P(4), P(5), P(6), P(7))
      Case 9: Set VResult = Method(DispID).Request(P(0), P(1), P(2), P(3), P(4), P(5), P(6), P(7), P(8))
    End Select
  End If
End Function
