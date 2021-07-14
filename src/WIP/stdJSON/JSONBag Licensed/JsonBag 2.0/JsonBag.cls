VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JsonBag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = true
Option Explicit

'Not a real (fractional) number, but Major.Minor integers:
Private Const CLASS_VERSION As String = "2.0"
'
'JsonBag Class (JsonBag.cls)
'
'Version 2.0
'
'A parser/serializer class for JSON data interchange written in Visual
'Basic 6.0 (some versions usable in Office VBA with little or no
'modification).
'
'
'Copyright 2013, 2014, 2015 Robert D. Riemersma, Jr.
'
'Licensed under the Apache License, Version 2.0 (the "License");
'you may not use this file except in compliance with the License.
'You may obtain a copy of the License at
'
'    http://www.apache.org/licenses/LICENSE-2.0
'
'Unless required by applicable law or agreed to in writing, software
'distributed under the License is distributed on an "AS IS" BASIS,
'WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'See the License for the specific language governing permissions and
'limitations under the License.

'Character constants.
Private Const LBRACE As String = "{"
Private Const RBRACE As String = "}"
Private Const LBRACKET As String = "["
Private Const RBRACKET As String = "]"
Private Const COLON As String = ":"
Private Const COMMA As String = ","
Private Const BLANKSPACE As String = " "
Private Const QUOTE As String = """"
Private Const PLUS As String = "+"
Private Const MINUS As String = "-"
Private Const RADIXPOINT As String = "." 'Always a period since we're locale-blind.
Private Const JSON_EXP As String = "e"
Private Const ZERO As String = "0"
Private Const NINE As String = "9"
Private Const REVSOLIDUS As String = "\"

Private Const WHITE_SPACE As String = vbTab & vbLf & vbCr & " "

'AscW() value constants.
Private Const LBRACE_W As Long = &H7B&
Private Const RBRACE_W As Long = &H7D&
Private Const LBRACKET_W As Long = &H5B&
Private Const RBRACKET_W As Long = &H5D&
Private Const COLON_W As Long = &H3A&
Private Const COMMA_W As Long = &H2C&
Private Const NULL_W As Long = 0
Private Const BLANKSPACE_W As Long = &H20&
Private Const QUOTE_W As Long = &H22&
Private Const PLUS_W As Long = &H2B&
Private Const MINUS_W As Long = &H2D&
Private Const RADIXPOINT_W As Long = &H2E& 'Always a period since we're locale-blind.
Private Const JSON_EXP_W As Long = &H65&
Private Const ZERO_W As Long = &H30&
Private Const NINE_W As Long = &H39&
Private Const REVSOLIDUS_W As Long = &H5C&

Private Const S_OK As Long = 0
Private Const VARIANT_ALPHABOOL As Long = &H2&
Private Const LOCALE_INVARIANT As Long = 127& 'Used to do VT conversions with the invariant locale.

#If Vba7 Then
'In VB6 these will all be marked in error by the IDE, but the compiler will never
'See them so you can ignore the error color (typically red):
    #If Win64 Then
    Private Declare PtrSafe Function HashData Lib "shlwapi" ( _
        ByVal pbData As LongPtr, _
        ByVal cbData As Long, _
        ByVal pbHash As LongPtr, _
        ByVal cbHash As Long) As Long
    
    Private Declare PtrSafe Function StrSpn Lib "shlwapi" Alias "StrSpnW" ( _
        ByVal psz As LongPtr, _
        ByVal pszSet As Long) As Long
    
    Private Declare PtrSafe Function VariantChangeTypeEx Lib "oleaut32" ( _
        ByRef vargDest As Variant, _
        ByRef varSrc As Variant, _
        ByVal lcid As Long, _
        ByVal wFlags As Integer, _
        ByVal vt As VbVarType) As Long
    #Else
    Private Declare PtrSafe Function HashData Lib "shlwapi" ( _
        ByVal pbData As Long, _
        ByVal cbData As Long, _
        ByVal pbHash As Long, _
        ByVal cbHash As Long) As Long
    
    Private Declare PtrSafe Function StrSpn Lib "shlwapi" Alias "StrSpnW" ( _
        ByVal psz As Long, _
        ByVal pszSet As Long) As Long
    
    Private Declare PtrSafe Function VariantChangeTypeEx Lib "oleaut32" ( _
        ByRef vargDest As Variant, _
        ByRef varSrc As Variant, _
        ByVal lcid As Long, _
        ByVal wFlags As Integer, _
        ByVal vt As VbVarType) As Long
    #End If
#Else
Private Declare Function HashData Lib "shlwapi" ( _
    ByVal pbData As Long, _
    ByVal cbData As Long, _
    ByVal pbHash As Long, _
    ByVal cbHash As Long) As Long

Private Declare Function StrSpn Lib "shlwapi" Alias "StrSpnW" ( _
    ByVal psz As Long, _
    ByVal pszSet As Long) As Long

Private Declare Function VariantChangeTypeEx Lib "oleaut32" ( _
    ByRef vargDest As Variant, _
    ByRef varSrc As Variant, _
    ByVal lcid As Long, _
    ByVal wFlags As Integer, _
    ByVal vt As VbVarType) As Long
#End If

Private TypeNameOfMe As String 'Used in raising exceptions.
Private Names As Collection
Private Values As Collection
Private CursorIn As Long 'Scan position within JSON input string.
Private LengthIn As Long 'Length of JSON input string.
Private TextOut As String 'Buffer to build JSON output string in.
Private CursorOut As Long 'Append position within JSON output string.
Private NumberType As VbVarType

Private mIsArray As Boolean
Private mDecimalMode As Boolean
Private mWhitespace As Boolean 'True to use indenting and newlines on JSON Get.
Private mWhitespaceIndent As Integer 'Number of spaces per level for whitespace indenting.
Private mWhitespaceNewLine As String

'=== Public Properties =================================================================

Public Property Get Count() As Long
Attribute Count.VB_Description = "Count of Items in the list."
    Count = Values.Count
End Property

Public Property Get DecimalMode() As Boolean
Attribute DecimalMode.VB_Description = "Causes numbers to be parsed and stored as Decimal type instead of Double."
    DecimalMode = mDecimalMode
End Property

Public Property Let DecimalMode(ByVal RHS As Boolean)
    Dim Item As Variant
    
    mDecimalMode = RHS
    If mDecimalMode Then
        NumberType = vbDecimal
    Else
        NumberType = vbDouble
    End If
    For Each Item In Values
        If TypeOf Item Is JsonBag Then
            Item.DecimalMode = mDecimalMode
        End If
    Next
End Property

Public Property Let IsArray(ByVal RHS As Boolean)
    If Values.Count > 0 Then
        Err.Raise &H80049900, TypeNameOfMe, "Cannot change IsArray setting after items have been added"
    Else
        mIsArray = RHS
    End If
End Property

Public Property Get IsArray() As Boolean
Attribute IsArray.VB_Description = "True if this object is a JSON array instead of a JSON object.  Must be set before first item is added."
    IsArray = mIsArray
End Property

'Default property.
Public Property Get Item(ByVal Key As Variant) As Variant
Attribute Item.VB_Description = "Retrieve an Item by Key or Index, change or add a new Item."
Attribute Item.VB_UserMemId = 0
    'Retrieval works either by key or index for "objects" but only
    'by index for "arrays."
    
    Dim PrefixedKey As String
    
    If IsNull(Key) Then Err.Raise &H80049904, TypeNameOfMe, "Key must be String or an index"
    If VarType(Key) = vbString Then
        If mIsArray Then
            Err.Raise &H80049908, TypeNameOfMe, "Array values can only be acessed by index"
        End If
        
        If Exists(Key) Then
            PrefixedKey = PrefixHash(Key)
            If IsObject(Values.Item(PrefixedKey)) Then
                Set Item = Values.Item(PrefixedKey)
            Else
                Item = Values.Item(PrefixedKey)
            End If
        Else
            Err.Raise &H8004990C, TypeNameOfMe, "Requested Item by key doesn't exist (case mismatch?)"
        End If
    Else
        If IsObject(Values.Item(Key)) Then
            Set Item = Values.Item(Key)
        Else
            Item = Values.Item(Key)
        End If
    End If
End Property

Public Property Let Item(Optional ByVal Key As Variant = Null, ByVal RHS As Variant)
    'Add new Item or change existing Item's value.
    '
    'When IsArray = True:
    '
    '   Pass a Null as Key to add a new item at the end of the "array."
    '
    '   Pass an index (Long) as Key to assign a new value to an
    '   existing Item.  However if the index is greater than .Count
    '   the value is added as a new entry at the end of the "array."
    '
    'When IsArray = False (i.e. a JSON "object"):
    '
    '   Pass a name (String) as Key.  If the named Item exists its
    '   value is updated.  If it does not exist a new Item is added.
    '
    'Item reassignment for existing items (assign new value) is
    'implemented as remove and re-add.  This means changing the value
    'of an "object's" Item moves it to the end of the list.
    
    Dim PrefixedKey As String
    
    With Values
        If mIsArray Then
            If VarType(Key) = vbString Then
                Err.Raise &H8004990E, TypeNameOfMe, "Array values can only be changed by index or added via Null"
            End If
            
            If IsNull(Key) Then
                .Add RHS 'Add at end.
                Names.Add .Count, CStr(.Count)
            Else
                If Key > .Count Then
                    .Add RHS 'Add at end.
                    Names.Add .Count, CStr(.Count)
                Else
                    .Remove Key
                    .Add RHS, , Key 'Insert into position.
                End If
            End If
        Else
            If VarType(Key) <> vbString Then
                Err.Raise &H80049910, TypeNameOfMe, "Object values can only be changed or added by key not by index"
            End If
            
            PrefixedKey = PrefixHash(Key)
            On Error Resume Next
            .Add RHS, PrefixedKey
            If Err Then
                On Error GoTo 0
                'Add failed, Key must already exist.  Remove/re-add.  Remove Name.
                .Remove PrefixedKey
                .Add RHS, PrefixedKey
                Names.Remove PrefixedKey
            Else
                On Error GoTo 0
            End If
            'Add Name.
            Names.Add Key, PrefixedKey
        End If
    End With
End Property

Public Property Set Item(ByVal Key As Variant, ByVal RHS As Variant)
    'This is just an alias for Let since we don't have to do anything
    'different.
    '
    'This allows either Let or Set to be used by client logic.
    
    Item(Key) = RHS
End Property

Public Property Get JSON() As String
Attribute JSON.VB_Description = "A string representing the serialized contents of the object."
    CursorOut = 1
    SerializeItem vbNullString, Me
    JSON = Left$(TextOut, CursorOut - 1)
    
    'Clear for next reuse.  Do it here to reclaim space.
    TextOut = vbNullString
End Property

Public Property Let JSON(ByRef RHS As String)
    Clear
    
    CursorIn = 1
    LengthIn = Len(RHS)
    
    SkipWhitespace RHS
    
    Select Case Mid$(RHS, CursorIn, 1)
        Case LBRACE
            CursorIn = CursorIn + 1
            mIsArray = False
            ParseObject RHS, CursorIn, Len(RHS)
        Case LBRACKET
            CursorIn = CursorIn + 1
            mIsArray = True
            ParseArray RHS, CursorIn, Len(RHS)
        Case Else
            Error99A0 "either " & LBRACE & " or " & LBRACKET, CursorIn
    End Select
End Property

Public Property Get Name(ByVal Index As Long) As Variant
Attribute Name.VB_Description = "Retrieves list Item Name by Index."
    If mIsArray Then
        Name = CLng(Names.Item(Index))
    Else
        Name = Names.Item(Index)
    End If
End Property

Public Property Get Version() As String()
    Version = Split(CLASS_VERSION, ".")
End Property

Public Property Get Whitespace() As Boolean
    Whitespace = mWhitespace
End Property

Public Property Let Whitespace(ByVal RHS As Boolean)
    Dim Item As Variant
    
    mWhitespace = RHS
    For Each Item In Values
        If TypeOf Item Is JsonBag Then
            Item.Whitespace = mWhitespace
        End If
    Next
End Property

Public Property Get WhitespaceIndent() As Integer
    WhitespaceIndent = mWhitespaceIndent
End Property

Public Property Let WhitespaceIndent(ByVal RHS As Integer)
    Dim Item As Variant
    
    If 1 > RHS Or RHS > 32 Then Err.Raise 380 'Invalid property value.
    
    mWhitespaceIndent = RHS
    For Each Item In Values
        If TypeOf Item Is JsonBag Then
            Item.WhitespaceIndent = mWhitespaceIndent
        End If
    Next
End Property

Public Property Get WhitespaceNewLine() As String
    WhitespaceNewLine = mWhitespaceNewLine
End Property

Public Property Let WhitespaceNewLine(ByVal RHS As String)
    Dim Item As Variant
    
    If Len(RHS) = 0 Then Err.Raise 380 'Invalid property value.
    
    mWhitespaceNewLine = RHS
    For Each Item In Values
        If TypeOf Item Is JsonBag Then
            Item.WhitespaceNewLine = mWhitespaceNewLine
        End If
    Next
End Property

'=== Public Methods ====================================================================

Public Function AddNewArray(Optional ByVal Key As Variant = Null) As JsonBag
Attribute AddNewArray.VB_Description = "Create new ""array"" type Item and add it to the list, returning a reference to it."
    Dim NewArray As JsonBag
    
    Set NewArray = New JsonBag
    NewArray.IsArray = True
    Set Item(Key) = NewArray
    Set AddNewArray = NewArray
End Function

Public Function AddNewObject(Optional ByVal Key As Variant = Null) As JsonBag
Attribute AddNewObject.VB_Description = "Create new ""object"" type Item and add it to the list, returning a reference to it."
    Dim NewObject As JsonBag
    
    Set NewObject = New JsonBag
    Set Item(Key) = NewObject
    Set AddNewObject = NewObject
End Function

Public Sub Clear()
Attribute Clear.VB_Description = "Clears all data and sets IsArray to False."
    Set Names = New Collection
    Set Values = New Collection
End Sub

Public Function Exists(ByVal Key As Variant) As Boolean
Attribute Exists.VB_Description = "Returns True if item specified by Key or Index is present."
    Dim Name As String
    
    If VarType(Key) = vbString Then
        On Error Resume Next
        Name = Names.Item(PrefixHash(Key))
    Else
        On Error Resume Next
        Name = Names.Item(Key)
    End If
    Exists = Err.Number = 0
    Err.Clear
End Function

'Marked as hidden and ProcedureID = -4
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "Iterates over the Item names."
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = Values.[_NewEnum]
End Function

Public Sub Remove(ByVal Key As Variant)
Attribute Remove.VB_Description = "Removes Item specified by Key or Index."
    'Allow remove by Key or Index (only by Index for arrays).  If the item
    'does not exist return silently.
    
    Dim PrefixedKey As String
    Dim I As Long
    
    If VarType(Key) = vbString Then
        If mIsArray Then Err.Raise &H8004991C, TypeNameOfMe, "Must remove by index for arrays"
        
        PrefixedKey = PrefixHash(Key)
        On Error Resume Next
        Names.Remove PrefixedKey
        If Err Then
            Exit Sub
        End If
        On Error GoTo 0
        Values.Remove PrefixedKey
    Else
        If 1 <= Key And Key <= Values.Count Then
            Values.Remove Key
            If IsArray Then
                For I = Names.Count To Key Step -1
                    Names.Remove I
                Next
                For I = Key To Values.Count
                    Names.Add I, CStr(I)
                Next
            Else
                Names.Remove Key
            End If
        End If
    End If
End Sub

'=== Friend Methods (do not call from client logic) ====================================

Friend Sub ParseArray( _
    ByRef Text As String, _
    ByRef StartCursor As Long, _
    ByVal TextLength As Long)
    'This call is made within the context of the instance at hand.
    
    Dim ArrayValue As Variant
    
    CursorIn = StartCursor
    LengthIn = TextLength

    Do
        SkipWhitespace Text
        Select Case AscW(Mid$(Text, CursorIn, 1))
            Case COMMA_W
                CursorIn = CursorIn + 1
            Case RBRACKET_W
                CursorIn = CursorIn + 1
                Exit Do
            Case Else
                ParseValue Text, ArrayValue
                Values.Add ArrayValue
                Names.Add Values.Count
        End Select
    Loop
    StartCursor = CursorIn
End Sub

Friend Sub ParseObject( _
    ByRef Text As String, _
    ByRef StartCursor As Long, _
    ByVal TextLength As Long)
    'This call is made within the context of the instance at hand.
    
    Dim ItemName As String
    Dim Value As Variant
    Dim FoundFirstItem As Boolean
    
    CursorIn = StartCursor
    LengthIn = TextLength
    
    Do
        SkipWhitespace Text
        Select Case AscW(Mid$(Text, CursorIn, 1))
            Case QUOTE_W
                CursorIn = CursorIn + 1
                ItemName = ParseName(Text)
                ParseValue Text, Value
                Item(ItemName) = Value
                FoundFirstItem = True
            Case COMMA_W
                CursorIn = CursorIn + 1
                If Not FoundFirstItem Then
                    Err.Raise &H80049920, TypeNameOfMe, "Found "","" before first property at character " & CStr(CursorIn - 1)
                End If
            Case RBRACE_W
                CursorIn = CursorIn + 1
                Exit Do
            Case Else
                Error99A0 ", or }", CursorIn - 1
        End Select
    Loop
    StartCursor = CursorIn
End Sub

'=== Private Methods ===================================================================

Private Sub Cat(ByRef NewText As String)
    Const TEXT_CHUNK As Long = 512 'Allocation size for destination buffer Text.
    Dim LenNew As Long
    
    LenNew = Len(NewText)
    If LenNew > 0 Then
        If CursorOut + LenNew - 1 > Len(TextOut) Then
            If LenNew > TEXT_CHUNK Then
                TextOut = TextOut & Space$(LenNew + TEXT_CHUNK)
            Else
                TextOut = TextOut & Space$(TEXT_CHUNK)
            End If
        End If
        Mid$(TextOut, CursorOut, LenNew) = NewText
        CursorOut = CursorOut + LenNew
    End If
End Sub

Private Sub Error99A0(ByVal Symbol As String, ByVal Position As Long)
    Err.Raise &H800499A0, TypeNameOfMe, "Expected " & Symbol & " at character " & CStr(Position)
End Sub

Private Sub Error99B0(ByVal Position As Long)
    Err.Raise &H800499B0, TypeNameOfMe, "Bad string character escape at character " & CStr(Position)
End Sub

Private Function ParseName(ByRef Text As String) As String
    ParseName = ParseString(Text)
    
    SkipWhitespace Text
    If Mid$(Text, CursorIn, 1) <> COLON Then
        Error99A0 COLON, CursorIn
    End If
    CursorIn = CursorIn + 1
End Function

Private Function ParseNumber(ByRef Text As String) As Variant
    Const BUILD_CHUNK As Long = 10
    Dim SaveCursor As Long
    Dim BuildString As String
    Dim BuildCursor As Long
    Dim Char As String
    Dim CharW As Long
    Dim GotDecPoint As Boolean
    Dim GotExpSign As Boolean
    
    SaveCursor = CursorIn 'Saved for "bad number format" error.
    BuildString = Space$(BUILD_CHUNK)
    
    'We know 1st char has been validated by the caller.
    BuildCursor = 1
    Mid$(BuildString, 1, 1) = Mid$(Text, CursorIn, 1)
    
    For CursorIn = CursorIn + 1 To LengthIn
        Char = LCase$(Mid$(Text, CursorIn, 1))
        CharW = AscW(Char)
        Select Case CharW
            Case ZERO_W To NINE_W
                'Do nothing.
            Case RADIXPOINT_W
                If GotDecPoint Then
                    Err.Raise &H80049924, TypeNameOfMe, "Second decimal point at character " & CStr(CursorIn)
                End If
                If Mid$(BuildString, BuildCursor, 1) = MINUS Then
                    Err.Raise &H80049928, TypeNameOfMe, "Digit expected at character " & CStr(CursorIn)
                End If
                GotDecPoint = True
            Case JSON_EXP_W
                CursorIn = CursorIn + 1
                Exit For
            Case Else
                Exit For
        End Select
        BuildCursor = BuildCursor + 1
        If BuildCursor > Len(BuildString) Then BuildString = BuildString & Space$(BUILD_CHUNK)
        Mid$(BuildString, BuildCursor, 1) = Char
    Next
    
    If CharW = JSON_EXP_W Then
        BuildCursor = BuildCursor + 1
        If BuildCursor > Len(BuildString) Then BuildString = BuildString & Space$(BUILD_CHUNK)
        Mid$(BuildString, BuildCursor, 1) = Char
        
        For CursorIn = CursorIn To LengthIn
            Char = Mid$(Text, CursorIn, 1)
            Select Case AscW(Char)
                Case ZERO_W To NINE_W
                    'Do nothing.
                Case PLUS_W, MINUS_W
                    If GotExpSign Then
                        Err.Raise &H8004992C, TypeNameOfMe, "Second exponent sign at character " & CStr(CursorIn)
                    End If
                    GotExpSign = True
                Case Else
                    Exit For
            End Select
            BuildCursor = BuildCursor + 1
            If BuildCursor > Len(BuildString) Then BuildString = BuildString & Space$(BUILD_CHUNK)
            Mid$(BuildString, BuildCursor, 1) = Char
        Next
    End If
    
    If CursorIn > LengthIn Then
        Err.Raise &H80049930, TypeNameOfMe, "Ran off end of string while parsing a number"
    End If
    
    ParseNumber = Left$(BuildString, BuildCursor)
    If VariantChangeTypeEx(ParseNumber, ParseNumber, LOCALE_INVARIANT, 0, NumberType) <> S_OK Then
        Err.Raise &H80049934, TypeNameOfMe, "Number overflow or parse error at character " & CStr(SaveCursor)
    End If
End Function

Private Function ParseString(ByRef Text As String) As String
    Const BUILD_CHUNK As Long = 32
    Dim BuildCursor As Long
    Dim Char As String
    
    ParseString = Space$(BUILD_CHUNK)
    
    For CursorIn = CursorIn To LengthIn
        Char = Mid$(Text, CursorIn, 1)
        Select Case AscW(Char)
            Case QUOTE_W
                CursorIn = CursorIn + 1
                ParseString = Left$(ParseString, BuildCursor)
                Exit Function 'Normal exit.
            Case REVSOLIDUS_W
                CursorIn = CursorIn + 1
                If CursorIn > LengthIn Then
                    Error99B0 CursorIn
                End If
                Char = LCase$(Mid$(Text, CursorIn, 1)) 'Accept uppercased escape symbols.
                Select Case Char
                    Case QUOTE, REVSOLIDUS, "/"
                        'Do nothing.
                    Case "b"
                        Char = vbBack
                    Case "f"
                        Char = vbFormFeed
                    Case "n"
                        Char = vbLf
                    Case "r"
                        Char = vbCr
                    Case "t"
                        Char = vbTab
                    Case "u"
                        CursorIn = CursorIn + 1
                        If LengthIn - CursorIn < 3 Then
                            Error99B0 CursorIn
                        End If
                        On Error Resume Next
                        Char = ChrW$(CLng("&H0" & Mid$(Text, CursorIn, 4)))
                        If Err Then
                            On Error GoTo 0
                            Error99B0 CursorIn
                        End If
                        On Error GoTo 0
                        CursorIn = CursorIn + 3 'Not + 4 because For loop will increment again.
                    Case Else
                        Error99B0 CursorIn
                End Select
            Case Is >= BLANKSPACE_W
                'Do nothing, i.e. fall through passing Char unchanged.
            Case Else
                Err.Raise &H80049938, TypeNameOfMe, "Invalid string character at " & CStr(CursorIn)
        End Select
        BuildCursor = BuildCursor + 1
        If BuildCursor > Len(ParseString) Then ParseString = ParseString & Space$(BUILD_CHUNK)
        Mid$(ParseString, BuildCursor, 1) = Char
    Next
    
    Error99A0 QUOTE, LengthIn + 1
End Function

Private Sub ParseValue(ByRef Text As String, ByRef Value As Variant)
    Dim SubBag As JsonBag
    Dim Token As String
    
    SkipWhitespace Text
    Select Case AscW(Mid$(Text, CursorIn, 1))
        Case QUOTE_W
            CursorIn = CursorIn + 1
            Value = ParseString(Text)
        Case MINUS_W, ZERO_W To NINE_W
            Value = ParseNumber(Text)
        Case LBRACE_W
            CursorIn = CursorIn + 1
            Set SubBag = New JsonBag
            With SubBag
                .DecimalMode = mDecimalMode
                .IsArray = False
                .ParseObject Text, CursorIn, LengthIn
                .Whitespace = mWhitespace
                .WhitespaceIndent = mWhitespaceIndent
            End With
            Set Value = SubBag
        Case LBRACKET_W
            CursorIn = CursorIn + 1
            Set SubBag = New JsonBag
            With SubBag
                .DecimalMode = mDecimalMode
                .IsArray = True
                .ParseArray Text, CursorIn, LengthIn
                .Whitespace = mWhitespace
                .WhitespaceIndent = mWhitespaceIndent
            End With
            Set Value = SubBag
        Case Else
            If Mid$(Text, CursorIn, 1) = COLON Then
                Err.Raise &H800499C0, TypeNameOfMe, "Unexpected "":"" at character " & CStr(CursorIn)
            Else
                'Special value tokens.
                Token = LCase$(Mid$(Text, CursorIn, 4))
                If Token = "null" Then
                    Value = Null
                    CursorIn = CursorIn + 4
                ElseIf Token = "true" Then
                    Value = True
                    CursorIn = CursorIn + 4
                Else
                    Token = LCase$(Mid$(Text, CursorIn, 5))
                    If Token = "false" Then
                        Value = False
                        CursorIn = CursorIn + 5
                    Else
                        Err.Raise &H8004993C, TypeNameOfMe, "Bad value at character " & CStr(CursorIn)
                    End If
                End If
            End If
    End Select
End Sub

Private Function PrefixHash(ByVal KeyString As String) As String
    'This is used to make Collection access by key case-sensitive.
    
    Dim Hash As Long
    
    HashData StrPtr(KeyString), 2 * Len(KeyString), VarPtr(Hash), 4
    PrefixHash = Right$("0000000" & Hex$(Hash), 8) & KeyString
End Function

Private Sub SerializeItem( _
    ByVal ItemName As String, _
    ByRef Item As Variant, _
    Optional ByVal Level As Integer)
    'For outer level call set CursorOut = 1 before calling.  For outer level call
    'or array calls pass vbNullString as ItemName for "anonymity."
    
    Const TEXT_CHUNK As Long = 64
    Dim Indent As String
    Dim Anonymous As Boolean
    Dim Name As Variant
    Dim ItemIndex As Long
    Dim TempItem As Variant
    Dim ItemBag As JsonBag
    Dim SubBag As JsonBag
    Dim ItemText As String
    Dim ArrayItem As Variant
    
    If Whitespace Then
        Indent = Space$(mWhitespaceIndent * Level)
    End If
    
    Anonymous = StrPtr(ItemName) = 0 'Check for vbNullString.
    If Not Anonymous Then
        'Not vbNullString so we have a named Item.
        If Whitespace Then Cat Indent
        Cat SerializeString(ItemName) & COLON
    End If
    
    Select Case VarType(Item)
        Case vbEmpty, vbNull 'vbEmpty case should actually never occur.
            If Whitespace And Anonymous Then Cat Indent
            Cat "null"
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte, vbBoolean
            If Whitespace And Anonymous Then Cat Indent
            If VariantChangeTypeEx(TempItem, _
                                   Item, _
                                   LOCALE_INVARIANT, _
                                   VARIANT_ALPHABOOL, _
                                   vbString) <> S_OK Then
                Err.Raise &H80049940, TypeNameOfMe, "Item """ & ItemName & """ value " & CStr(Item) & " failed to serialize"
            End If
            Cat LCase$(TempItem) 'Convert to lowercase "true" and "false" and "1.234e34" and such.
        Case vbString
            If Whitespace And Anonymous Then Cat Indent
            Cat SerializeString(Item)
        Case vbObject
            Set ItemBag = Item
            If ItemBag.IsArray Then
                If Whitespace And Anonymous Then Cat Indent
                Cat LBRACKET
                If ItemBag.Count < 1 Then
                    Cat RBRACKET
                Else
                    If Whitespace Then Cat mWhitespaceNewLine
                    With ItemBag
                        For ItemIndex = 1 To .Count
                            If IsObject(.Item(ItemIndex)) Then
                                Set TempItem = .Item(ItemIndex)
                            Else
                                TempItem = .Item(ItemIndex)
                            End If
                            SerializeItem vbNullString, TempItem, Level + 1
                            Cat COMMA
                            If Whitespace Then Cat mWhitespaceNewLine
                        Next
                    End With
                    If Whitespace Then
                        CursorOut = CursorOut - 3
                        Cat mWhitespaceNewLine & Indent & RBRACKET
                    Else
                        Mid$(TextOut, CursorOut - 1) = RBRACKET
                    End If
                End If
            Else
                If Whitespace And Anonymous Then Cat Indent
                Cat LBRACE
                If ItemBag.Count < 1 Then
                    Cat RBRACE
                Else
                    If Whitespace Then Cat mWhitespaceNewLine
                    With ItemBag
                        For ItemIndex = 1 To .Count
                            If IsObject(.Item(ItemIndex)) Then
                                Set TempItem = .Item(ItemIndex)
                            Else
                                TempItem = .Item(ItemIndex)
                            End If
                            SerializeItem .Name(ItemIndex), TempItem, Level + 1
                            Cat COMMA
                            If Whitespace Then Cat mWhitespaceNewLine
                        Next
                    End With
                    If Whitespace Then
                        CursorOut = CursorOut - 3
                        Cat mWhitespaceNewLine & Indent & RBRACE
                    Else
                        Mid$(TextOut, CursorOut - 1) = RBRACE
                    End If
                End If
            End If
        Case Else
            Err.Raise &H80049944, TypeNameOfMe, """Item " & ItemName & """ unknown/unsupported type = " & CStr(VarType(Item))
    End Select
End Sub

Private Function SerializeString(ByVal Text As String) As String
    Dim BuildString As String
    Dim BuildCursor As Long
    Dim TextCursor As Long
    Dim Char As String
    Dim intChar As Integer

    BuildString = Space$(3 * Len(Text) \ 2)
    BuildCursor = 1
    StringCat BuildString, BuildCursor, QUOTE
    For TextCursor = 1 To Len(Text)
        Char = Mid$(Text, TextCursor, 1)
        Select Case Char
            Case QUOTE, REVSOLIDUS
                StringCat BuildString, BuildCursor, REVSOLIDUS & Char
            Case vbBack
                StringCat BuildString, BuildCursor, REVSOLIDUS & "b"
            Case vbFormFeed
                StringCat BuildString, BuildCursor, REVSOLIDUS & "f"
            Case vbLf
                StringCat BuildString, BuildCursor, REVSOLIDUS & "n"
            Case vbCr
                StringCat BuildString, BuildCursor, REVSOLIDUS & "r"
            Case vbTab
                StringCat BuildString, BuildCursor, REVSOLIDUS & "t"
            Case " " To "!", "#" To LBRACKET, RBRACKET To "~"
                StringCat BuildString, BuildCursor, Char
            Case Else
                intChar = AscW(Char)
                Select Case intChar
                    Case 0 To &H1F, &H7F To &H9F, &H34F, &H200B To &H200F, _
                         &H2028 To &H202E, &H2060, &HFE01 To &HFE0F, _
                         &HFEFF, &HFFFD, &HD800 To &HDFFF
                        StringCat BuildString, BuildCursor, _
                               REVSOLIDUS & "u" & Right$("000" & Hex$(intChar), 4)
                    Case Else
                        StringCat BuildString, BuildCursor, Char
                End Select
        End Select
    Next
    StringCat BuildString, BuildCursor, QUOTE
    SerializeString = Left$(BuildString, BuildCursor - 1)
End Function

Private Sub SkipWhitespace(ByRef Text As String)
    CursorIn = CursorIn + StrSpn(StrPtr(Text) + 2 * (CursorIn - 1), StrPtr(WHITE_SPACE))
End Sub

Private Sub StringCat(ByRef TextOut As String, ByRef CursorOut, ByRef NewText As String)
    Const TEXT_CHUNK As Long = 64 'Allocation size for destination buffer Text.
    Dim LenNew As Long
    
    LenNew = Len(NewText)
    If LenNew > 0 Then
        If CursorOut + LenNew - 1 > Len(TextOut) Then
            If LenNew > TEXT_CHUNK Then
                TextOut = TextOut & Space$(LenNew + TEXT_CHUNK)
            Else
                TextOut = TextOut & Space$(TEXT_CHUNK)
            End If
        End If
        Mid$(TextOut, CursorOut, LenNew) = NewText
        CursorOut = CursorOut + LenNew
    End If
End Sub

'=== Private Events ====================================================================

Private Sub Class_Initialize()
    TypeNameOfMe = TypeName(Me)
    Clear 'Creates Collections.
    DecimalMode = False
    WhitespaceIndent = 4 'Default.
    WhitespaceNewLine = vbNewLine 'Default.
End Sub
