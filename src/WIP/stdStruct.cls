VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdCStruct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Spec:
'Constructors:
'    Create()
'    CreateEmpty()
'
'Instance Methods:
'  Type definition:
'    Sub typedef(sLineDef)
'  Type usage:
'    Get lines as Dictionary
'    Get pointer as LongPtr / Long
'    Get data as Byte()
'    Get item(sKeyName as string) as variant
'    Let item(sKeyName as string, RHS as variant)
'
'Examples
' Dim WSAData as stdCStruct: set WSAData = stdCStruct.Create("WSAData")
' Call WSAData.typedef("#define WSADESCRIPTION_LEN	256")
' Call WSAData.typedef("#define WSASYS_STATUS_LEN	128")
' Call WSAData.typedef("typedef struct WSAData {")
' Call WSAData.typedef("	WORD	wVersion;")
' Call WSAData.typedef("	WORD	wHighVersion;")
' Call WSAData.typedef("	char	szDescription[WSADESCRIPTION_LEN+1];")
' Call WSAData.typedef("	char	szSystemStatus[WSASYS_STATUS_LEN+1];")
' Call WSAData.typedef("	unsigned short	iMaxSockets;")
' Call WSAData.typedef("	unsigned short	iMaxUdpDg;")
' Call WSAData.typedef("	char * 	lpVendorInfo;")
' Call WSAData.typedef("}")
'
' if WSAStartup(MAKEWORD(2, 2), WSAData.pointer) = 0 then Err.Raise 1, "Winsock Failed " & Err.LastDllError
' '...

'stdTable myTable
' |- stdArray rows
'    |- stdTableRow row
'       |- stdDictionary data
'
' |
'
'
'
