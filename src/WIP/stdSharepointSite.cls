VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdSharepointSite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'P.S. Useful information about APIs can be found here:
'http://sharepoint/sub/site/_vti_bin/Lists.asmx

Const ListPath As String = "/_vti_bin/Lists.asmx"

'Private data
Private pPath As String
Private pUser As String
Private pPass As String
Private pHost As String

'Request object used for reference
Public LastReq As Object

'Constructor
'@param sSharepointHost:string - The host domain of sharepoint
'@param sSharepointPath:string - Path to sharepoint site/subsite
'//sharepoint/some/sub/site/
'E.G. For "http://teams/wholesale/shopping/leicester/SitePages/Home.aspx"   the site sharepoint path is "http://teams/wholesale/shopping/leicester/"
'@param @optional sUsername: string - The user's username (used for authentication)
'@param @optional sPassword: string - The user's password (used for authentication)
Public Function Create(ByVal sSharepointHost As String, ByVal sSharePointPath As String, Optional ByVal sUsername As String = Empty, Optional ByVal sPassword As String = Empty) As stdSharepointSite
    Set Create = New stdSharepointSite
    Create.Init sSharepointHost, sSharePointPath, sUsername, sPassword
End Function
Public Sub Init(ByVal sSharepointHost As String, ByVal sSharePointPath As String, ByVal sUsername As String, ByVal sPassword As String)
    pPath = IIf(Right(sSharePointPath, 1) = "/", Left(sSharePointPath, Len(sSharePointPath) - 1), sSharePointPath)
    pUser = sUsername
    pPass = sPassword
    pHost = sSharepointHost
End Sub



'@param pageURL: string - Locaiton of the file
'@param checkoutToLocal: string - Checkout for local editing or for remote editing? Usually false
'@note  Missing param lastModified in format:    Format(lastModified, "dd mmm yyyy hh:MM:ss") - no way of getting lastModified date yet. Will be wanted in the future.
Public Function CheckOutFile(ByVal pageURL As String, ByVal CheckoutToLocal As Boolean) As Boolean
    Set LastReq = DispatchSOAP("CheckOutFile", pPath & ListPath, Array( _ 
      "pageUrl", pageURL, _ 
      "checkoutToLocal", LCase(CStr(CheckoutToLocal)), _ 
      "lastmodified", "" _ 
    ))
    CheckOutFile = LastReq.Status = 200
End Function

'@param pageURL: string - The file location
'@param comment: string - The version comment to add
'@param CheckinType: Integer - 0 = MinorCheckIn (+0.1), 1 = MajorCheckIn (+1.0), 2 = OverwriteCheckIn (set v1.0)
Public Function CheckInFile(ByVal pageURL As String, ByVal comment As String, Optional ByVal CheckinType As Integer = 0) As Boolean
    Set LastReq = DispatchSOAP("CheckInFile", pPath & ListPath, Array( _ 
      "pageUrl", pageURL, _ 
      "comment", comment, _ 
      "CheckinType", CheckinType _ 
    ))
    CheckInFile = LastReq.Status = 200
End Function

'@param pageURL: string - The file location
Public Function CheckOutDiscard(ByVal pageURL As String) As Boolean
    Set LastReq = DispatchSOAP("UndoCheckOut", pPath & ListPath, Array( _ 
      "pageUrl", pageURL _ 
    ))
    CheckOutDiscard = LastReq.Status = 200
End Function


'GetListCollection
'@desc Returns all lists implemented on the site specified.
Public Function GetListCollection() As Boolean
    Set LastReq = DispatchSOAP("GetListCollection", pPath & ListPath, Array( _ 
    ))
    GetListCollection = LastReq.Status = 200
End Function

'@param listName: string     - List title or GUID
'@param query: string        - Query to filter on (only return the filtered items)
'@param viewName: string     - The GUID of the view to get items of
'@param viewFields: string   - The fields to return
'@param rowLimit: string     - The number of rows that should be returned. Default is 999999
'@param queryOptions: string - Additional query options (see docs)
'@param webID: string        - Original web site (see docs)
'@docs https://docs.microsoft.com/en-us/previous-versions/office/developer/sharepoint-services/ms772599(v=office.12)
Public Function GetListItems(ByVal listName As String, Optional ByVal query As String = "", Optional ByVal viewName As String = "", Optional ByVal viewFields As String = "", Optional ByVal rowLimit As Long = 999999, Optional ByVal queryOptions As String, Optional ByVal webID As String = "") As Object
    Set LastReq = DispatchSOAP("GetListItems", pPath & ListPath, Array( _
        "listName", listName, _
        "viewName", viewName, _
        "query", query, _
        "viewFields", viewFields, _
        "rowLimit", rowLimit, _
        "QueryOptions", queryOptions, _
        "webID", webID _
    ))
    GetListItems = LastReq.Status = 200
End Function

'@param listName: string - List title or GUID
Public Function GetList(ByVal listName As String) As Boolean
    Set LastReq = DispatchSOAP("GetList", pPath & ListPath, Array( _
        "listName", listName _
    ))
End Function

'@param listName: string - List title or GUID
'@param updates: string - A Batch Element containing updates to make
'@docs Main          - https://docs.microsoft.com/en-us/previous-versions/office/developer/sharepoint-services/ms772668(v=office.12)
'@docs BatchElement  - https://docs.microsoft.com/en-us/previous-versions/office/developer/sharepoint-services/ms437562(v=office.12)
'@docs MethodElement - https://docs.microsoft.com/en-us/previous-versions/office/developer/sharepoint-services/ms459050(v=office.12)
Public Function UpdateListItems(ByVal listName As String, ByVal updates As String) As Boolean
    Set LastReq = DispatchSOAP("UpdateListItems", pPath & ListPath, Array( _ 
      "listName", listName, _ 
      "updates", updates _ 
    ))
    UpdateListItems = LastReq.Status = 200
End Function

'@param col: Collection - A collection of Updates created with CreateUpdate() method.
'@param sErrorHandling: string - Either "Continue" or "Return". Changes error handling of the update
Public Function UpdateListItems_CreateBatch(ByVal col As Collection, Optional ByVal sErrorHandling As String = "Continue")
    Dim sData As String: sData = ""
    sData = sData & "<Batch OnError=""Continue"">" & vbCrLf
    
    Dim vMethod
    For Each vMethod In col
        sData = sData & vMethod & vbCrLf
    Next
    
    UpdateListItems_CreateBatch = sData & "</Batch>"
End Function

'@param sUpdateID: string - A free text used to create an update identifier. This will be useful when parsing the result of the Sharepoint call
'@param sCmd: string      - The update command. Can either be "Delete", "Update" or "New"
'@param params...: string - An array of FieldName,FieldValue pairs e.g. Array("ID",1,"Title","Car","Price",30.132)
Public Function UpdateListItems_CreateMethod(ByVal sUpdateID As String, ByVal sCmd As String, params() As Variant)
    Dim sMethod As String: sMethod = ""
    sMethod = sMethod & "<Method ID='" & sUpdateID & "' Cmd='" & sCmd & "'>" & vbCrLf
    
    Dim i As Long:
    For i = LBound(params) To UBound(params) Step 2
        sMethod = sMethod & "<Field Name='" & params(i) & "'>" & params(i + 1) & "</Field>" & vbCrLf
    Next
    
    UpdateListItems_CreateMethod = sMethod & "</Method>"
End Function








'@param sPrimaryTag: string      - Name of operation as per sharepoint specification
'@param sSitePath: string        - The path to the sharepoint class entry point
'@param params: Array(String*2n) - Array of String pairs:
'   * Class method parameter name
'   * Class method parameter value
'this is used to generate SOAP XML, see #getSoapXML()
Private Function DispatchSOAP(ByVal sPrimaryTag As String, ByVal sSitePath As String, params As Variant) As Object
    Dim sSOAP As String: sSOAP = getSoapXML(sPrimaryTag, params)
    Set DispatchSOAP = HTTPPost(GetSOAPSchema(sPrimaryTag), pPath & ListPath, sSOAP)
End Function

'@param sPrimaryTag: string      - Name of operation as per sharepoint specification
Private Function GetSOAPSchema(ByVal sPrimaryTag As String) As String
    GetSOAPSchema = "http://schemas.microsoft.com/sharepoint/soap/" & sPrimaryTag
End Function

'Generates SOAP XML for calls to SharePoint
'@param sPrimaryTag: string      - Name of operation as per sharepoint specification
'@param params: Array(String*2n) - Array of String pairs:
'   * Class method parameter name
'   * Class method parameter value
Private Function getSoapXML(ByVal sPrimaryTag As String, params As Variant) As String
    'SOAP header
    Dim sData As String: sData = ""
    sData = sData & "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf
    sData = sData & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & vbCrLf
    sData = sData & "  <soap:Body>" & vbCrLf
    
    'SOAP Primary wrapper header
    sData = sData & "    <" & sPrimaryTag & " xmlns=""http://schemas.microsoft.com/sharepoint/soap/"">" & vbCrLf
    
    'Loop through and apply params
    Dim i As Long
    For i = LBound(params) To UBound(params) Step 2
        sData = sData & "      <" & params(i) & ">" & params(i + 1) & "</" & params(i) & ">" & vbCrLf
    Next
    
    'SOAP Primary wrapper footer
    sData = sData & "    </" & sPrimaryTag & ">'" & vbCrLf 'random quote
    
    'SOAP footer
    sData = sData & "  </soap:Body>" & vbCrLf
    sData = sData & "</soap:Envelope>" & vbCrLf
    
    'Return data
    getSoapXML = sData
End Function


Private Function HTTPPost(ByVal sSOAPAction As String, ByVal sPath As String, ByVal sData As String, Optional ByVal sContentType As String = "text/xml; charset=utf-8") As Object
    Dim oHTTP As Object: Set oHTTP = CreateObject("MSXML2.serverXMLHTTP")
    With oHTTP
        If pUser = "" And pPass = "" Then
            .Open "POST", sPath, False
        Else
            .Open "POST", sPath, False, pUser, pPass
        End If
        .SetRequestHeader "Content-Type", sContentType
        .SetRequestHeader "Host", pHost
        .SetRequestHeader "Content-Length", Len(sData)
        .SetRequestHeader "SOAPAction", sSOAPAction
        .Send sData
    End With
    
    Set HTTPPost = oHTTP
End Function

