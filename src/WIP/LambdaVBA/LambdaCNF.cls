VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LambdaCNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private initialized As Boolean
Private grammar As Object
Private startSymbol As String
Private lookupTable As Object

'Creates a tokenizer object
'  @param oTokenizer - Dictionary<string,string>
'Example:
'  oTokenizer("SYMBOL") = "RegexToFind"
'  oTokenizer("ADD") = "\+"
Function Create(oGrammar As Object, sStartSymbol As String) As LambdaCNF
    'Dim o As LambdaCNF
    Set Create = New LambdaCNF
    Call Create.Init(oGrammar, sStartSymbol)
End Function
Sub Init(ByVal oGrammar As Object, ByVal sStartSymbol As String)
    If Not initialized Then
        Set grammar = oGrammar
        startSymbol = sStartSymbol
        Set lookupTable = createLookupTable(grammar)
        
        initialized = True
    Else
        Err.Raise 1, "LambdaTokenizer::Init", "Object already initialized"
    End If
End Sub



'/**
' * Given a CNF:
' * ```js
' * {
' *   A:[{left:"C", right:"D"}],
' *   B:[{left:"D", right:"D"}, {left:"D", right:"E"}]
' * }
' * ```
' * Creates a lookup table:
' * ```js
' * {
' *   C: {D: [{left:"C", right:"D", defSymbol:"A"}]},
' *   D: {D: [{left:"D", right:"D", defSymbol:"B"}],
' *       E: [{left:"D", right:"E", defSymbol:"B"}]}
' * }
' * ```
' */
' Ultimately of the form:
' {
'   leftSymbol: { rightSymbol: array of CNF elements
'   ...
' }
Private Function createLookupTable(grammar As Object) As Object
    Dim lookup As Object: Set lookup = CreateObject("Scripting.Dictionary")
    
    
    ' Go through all definitions
    Dim symbol As Variant, symbols As Variant
    symbols = grammar.keys() 'set a variable to grammar.keys() so keys() is only called once.
    For Each symbol In symbols
        Dim options As Object
        Set options = grammar(symbol)
        
        'Go through the options of every definition
        Dim oOption As Variant
        For Each oOption In options.arr
            'Make sure an object exists to store all patterns that contain the left symbol
            If Not lookup.exists(oOption("left")) Then Set lookup(oOption("left")) = CreateObject("Scripting.Dictionary")
            
            'Make sure an object exists to store all the patterns that contain the left and right symbol
            Dim leftMap As Object: Set leftMap = lookup(oOption("left"))
            If Not leftMap(oOption("right")) Then Set leftMap(oOption("right")) = stdArray.Create()
            
            'Add pattern to the map
            Dim leftAndRightMap As stdArray: Set leftAndRightMap = leftMap(oOption("right"))
            
            'leftAndRightMap.push({...option, defSymbol: symbol});
            Call leftAndRightMap.push(MergeDicts(oOption, CreateDict("defSymbol", symbol)))
        Next
    Next
    Set createLookupTable = lookup
End Function



'/**
' * Creates a CYK table given an input string
' * @param oInput The input tokens
' * @param approximate Whether to allow changing a right token to improve the match
' * @returns The table
' */
Public Function performCYK(oInput As stdArray) As Object()
    'usage of re represents the column or range end, and rs the row or range start (as index of the token)
    Dim iTableLength As Long
    iTableLength = oInput.Length - 1
    
    Dim table() As Object
    ReDim table(0 To iTableLength, 0 To iTableLength) As Object
    
    Dim re As Long, rs As Long
    For re = 0 To iTableLength Step 1
        For rs = re To 0 Step -1
            'Initialise cell
            Set table(rs, re) = CreateDict( _
                "symbols", stdArray.Create(), _
                "definitions", stdArray.Create(), _
                "range", CreateDict("start", Null, "end", "null") _
            )
            
            'Get cell object for easier referencing
            Dim cell As Object
            Set cell = table(rs, re)
            
            'Initialise the value if it's on the diagonal
            If rs = re Then
                Dim s As String: s = oInput.item(rs + 1)("symbol")
                Call cell("symbols").push(s)
                
                'Copy all the definitions and symbols that match
                Dim t As Object
                If lookupTable.exists(s) Then
                    Set t = lookupTable(s)
                    
                    If t.exists(Empty) Then
                        If Not t(Empty) Is Nothing Then
                            Dim def As Variant
                            
                            For Each def In t(Empty).arr
                                If Not cell("symbols").includes(def("defSymbol")) Then
                                    Call cell("symbols").push(def("defSymbol"))
                                End If
                                Call cell("definitions").push(MergeDicts(CreateDict("index", rs), def))
                            Next
                        End If
                    End If
                End If
                
                Set cell("range") = oInput.item(rs + 1)("range")
            Else
                cell("range")("start") = table(rs, rs)("range")("start")
                cell("range")("end") = table(re, re)("range")("end")
            End If
        Next
    Next
        
    
    'Fill the table (skipping the diagonal since that's already initialized)
    For re = 1 To iTableLength Step 1
        For rs = re - 1 To 0 Step -1 ' TODO: check that is re-1 is really required here?
            Set cell = table(rs, re)
            
            'Go through all the options for combining the subparts.
            Dim k As Long
            For k = rs To re - 1 Step 1
                'Get this option for left and right pair
                Dim left As Object: Set left = table(rs, k)
                Dim right As Object: Set right = table(k + 1, re)
                
                'Go through all combinations of left and right symbol options
                Dim lsI As Long, lsLen As Long
                lsLen = left("symbols").Length - 1
                For lsI = 0 To lsLen Step 1
                    Dim leftSymbol As String
                    leftSymbol = left("symbols").item(lsI + 1)
                    
                    'Check the lookup to see if this fits a definition
                    If lookupTable.exists(leftSymbol) Then
                        'Get left definitions
                        Dim leftDefs As Object
                        Set leftDefs = lookupTable(leftSymbol)
                        
                        'Combine it with all the right options
                        Dim rsI As Long, rsLen As Long
                        rsLen = right("symbols").Length - 1
                        For rsI = 0 To rsLen Step 1
                            Dim rightSymbol As String: rightSymbol = right("symbols").item(rsI + 1)
                            
                            'If definitions don't exist skip
                            If leftDefs.exists(rightSymbol) Then
                                Dim defs As stdArray
                                Set defs = leftDefs(rightSymbol)
                                
                                'Add all the definitions to the cell
                                Dim defI As Long, defsLen As Long
                                defsLen = defs.Length - 1
                                For defI = 0 To defsLen
                                    Dim defFin As Object
                                    Set defFin = defs.item(defI + 1)
                                    If Not cell("symbols").includes(defFin("defSymbol")) Then
                                        Call cell("symbols").push(defFin("defSymbol"))
                                    End If
                                    Call cell("definitions").push(MergeDicts(CreateDict("index", k), defFin))
                                Next
                            End If
                        Next
                    End If
                Next
            Next
        Next
    Next
    
    performCYK = table
End Function

'/**
' * Retrieves the binary tree of some input
' * @param oInput The input to construct a tree for
' * @returns The tree
' */
Public Function createBinaryTree(oInput As stdArray) As Object
    'Obtain the table using CYK
    Dim table As Variant
    table = performCYK(oInput)
    
    'Check if we matched the whole input with the required start symbol, and return an error otherwise
    
    MsgBox "TO BE IMPLEMENTED", vbCritical
    
End Function


















Private Function MergeDicts(ByVal o1 As Object, ByVal o2 As Object) As Object
    Dim vKey As Variant, vKeys As Variant
    Set MergeDicts = CreateObject("scripting.dictionary")
    
    'Add keys from o1
    vKeys = o1.keys()
    For Each vKey In vKeys
        MergeDicts.ADD vKey, o1(vKey)
    Next
    
    'Add keys from o2
    vKeys = o2.keys()
    For Each vKey In vKeys
        MergeDicts.ADD vKey, o2(vKey)
    Next
End Function


Public Sub Test()
    'TEST CreateDict()
    Dim o As Object
    Set o = CreateDict("a", 1, "b", "hello world", "c", CreateDict(1, 2))
    Debug.Assert o!a = 1
    Debug.Assert o!b = "hello world"
    Debug.Assert VarType(o!c) = vbObject
    Debug.Assert o!c.Count = 1
    Debug.Assert o("c")(1) = 2
    
    
    'TEST MergeDicts()
    Dim o1 As Object, o2 As Object, oc As Object
    Set o1 = CreateDict("a", 1, "b", 2)
    Set o2 = CreateDict("c", 2, "d", 3)
    Set oc = MergeDicts(o1, o2)
    Debug.Assert o1.Count = 2
    Debug.Assert o2.Count = 2
    Debug.Assert Not o1.exists("c")
    Debug.Assert Not o1.exists("d")
    Debug.Assert Not o2.exists("a")
    Debug.Assert Not o2.exists("b")
    Debug.Assert oc("a") = 1
    Debug.Assert oc("b") = 2
    Debug.Assert oc("c") = 2
    Debug.Assert oc("d") = 3
    
    
    'TEST CreateLookupTable:
    '/**
    ' * Given a CNF:
    ' * ```js
    ' * {
    ' *   A:[{left:"C", right:"D"}],
    ' *   B:[{left:"D", right:"D"}, {left:"D", right:"E"}]
    ' * }
    ' * ```
    ' * Creates a lookup table:
    ' * ```js
    ' * {
    ' *   C: {D: [{left:"C", right:"D", defSymbol:"A"}]},
    ' *   D: {D: [{left:"D", right:"D", defSymbol:"B"}],
    ' *       E: [{left:"D", right:"E", defSymbol:"B"}]}
    ' * }
    ' * ```
    ' */
    Dim grammar As Object
    Set grammar = CreateDict( _
      "A", stdArray.Create(CreateCNFP("C", "D")), _
      "B", stdArray.Create(CreateCNFP("D", "D"), CreateCNFP("D", "E")) _
    )
    Dim grammarLookup As Object
    Set grammarLookup = createLookupTable(grammar)
    Debug.Assert grammarLookup.exists("C")
    Debug.Assert grammarLookup("C").exists("D")
    Debug.Assert grammarLookup("C")("D").item(1)("left") = "C"
    Debug.Assert grammarLookup("C")("D").item(1)("right") = "D"
    Debug.Assert grammarLookup("C")("D").item(1)("defSymbol") = "A"
    Debug.Assert grammarLookup.exists("D")
    Debug.Assert grammarLookup("D").exists("D")
    Debug.Assert grammarLookup("D")("D").item(1)("left") = "D"
    Debug.Assert grammarLookup("D")("D").item(1)("right") = "D"
    Debug.Assert grammarLookup("D")("D").item(1)("defSymbol") = "B"
    Debug.Assert grammarLookup("D").exists("E")
    Debug.Assert grammarLookup("D")("E").item(1)("left") = "D"
    Debug.Assert grammarLookup("D")("E").item(1)("right") = "E"
    Debug.Assert grammarLookup("D")("E").item(1)("defSymbol") = "B"
    
    
    
    '# TEST CNF
    'import {CNF} from "CYK-BNF-CFG-AST-creator";
    'const cnf = new CNF(
    '    {
    '        EXP: [
    '            {left: "ADD", right: "EXP", metaData:5}, // We can attach any metadata to an expression under the metaData key
    '            {left: "SUB", right: "EXP"},
    '            {left: "DIV", right: "EXP"},
    '            {left: "MUL", right: "EXP"},
    '            {left: "Num"},
    '        ],
    '        ADD: [{left: "EXP", right: "Add"}],
    '        SUB: [{left: "EXP", right: "Sub"}],
    '        DIV: [{left: "EXP", right: "Div"}],
    '        MUL: [{left: "EXP", right: "Mul"}],
    '    },
    '    "EXP" // The start symbol
    ');
    'const tree = cnf.performCYK([
    '    {symbol:"Num", text:"34", range:{start:0, end:2}},
    '    {symbol:"Add", text:"+", range:{start:2, end:3}},
    '    {symbol:"Num", text:"2", range:{start:3, end:4}}
    ']);
    
    Dim cnf As LambdaCNF
    'Note "ADD" is different to "Add" here
    Set cnf = LambdaCNF.Create( _
        CreateDict( _
            "EXP", stdArray.Create( _
                CreateCNFP("ADD", "EXP", 5), _
                CreateCNFP("SUB", "EXP", 5), _
                CreateCNFP("DIV", "EXP", 5), _
                CreateCNFP("MUL", "EXP", 5), _
                CreateCNFP("Num") _
            ), _
            "ADD", stdArray.Create(CreateCNFP("EXP", "Add")), _
            "SUB", stdArray.Create(CreateCNFP("EXP", "Sub")), _
            "DIV", stdArray.Create(CreateCNFP("EXP", "Div")), _
            "MUL", stdArray.Create(CreateCNFP("EXP", "Mul")) _
        ), _
        "EXP" _
    )
    Dim tokens As stdArray
    Set tokens = stdArray.Create()
    tokens.push CreateDict("symbol", "Num", "text", "34", "range", CreateDict("start", 0, "end", 2))
    tokens.push CreateDict("symbol", "Add", "text", "+", "range", CreateDict("start", 2, "end", 3))
    tokens.push CreateDict("symbol", "Num", "text", "2", "range", CreateDict("start", 3, "end", 4))
    tokens.push CreateDict("symbol", "Add", "text", "+", "range", CreateDict("start", 2, "end", 3))
    tokens.push CreateDict("symbol", "Num", "text", "2", "range", CreateDict("start", 3, "end", 4))
    tokens.push CreateDict("symbol", "Add", "text", "+", "range", CreateDict("start", 2, "end", 3))
    tokens.push CreateDict("symbol", "Num", "text", "2", "range", CreateDict("start", 3, "end", 4))
    tokens.push CreateDict("symbol", "Add", "text", "+", "range", CreateDict("start", 2, "end", 3))
    tokens.push CreateDict("symbol", "Num", "text", "2", "range", CreateDict("start", 3, "end", 4))
    tokens.push CreateDict("symbol", "Add", "text", "+", "range", CreateDict("start", 2, "end", 3))
    tokens.push CreateDict("symbol", "Num", "text", "2", "range", CreateDict("start", 3, "end", 4))
    tokens.push CreateDict("symbol", "Add", "text", "+", "range", CreateDict("start", 2, "end", 3))
    tokens.push CreateDict("symbol", "Num", "text", "2", "range", CreateDict("start", 3, "end", 4))
    tokens.push CreateDict("symbol", "Add", "text", "+", "range", CreateDict("start", 2, "end", 3))
    tokens.push CreateDict("symbol", "Num", "text", "2", "range", CreateDict("start", 3, "end", 4))
    tokens.push CreateDict("symbol", "Add", "text", "+", "range", CreateDict("start", 2, "end", 3))
    tokens.push CreateDict("symbol", "Num", "text", "2", "range", CreateDict("start", 3, "end", 4))
    tokens.push CreateDict("symbol", "Add", "text", "+", "range", CreateDict("start", 2, "end", 3))
    tokens.push CreateDict("symbol", "Num", "text", "2", "range", CreateDict("start", 3, "end", 4))
    tokens.push CreateDict("symbol", "Add", "text", "+", "range", CreateDict("start", 2, "end", 3))
    tokens.push CreateDict("symbol", "Num", "text", "2", "range", CreateDict("start", 3, "end", 4))
    tokens.push CreateDict("symbol", "Add", "text", "+", "range", CreateDict("start", 2, "end", 3))
    tokens.push CreateDict("symbol", "Num", "text", "2", "range", CreateDict("start", 3, "end", 4))
    tokens.push CreateDict("symbol", "Add", "text", "+", "range", CreateDict("start", 2, "end", 3))
    tokens.push CreateDict("symbol", "Num", "text", "2", "range", CreateDict("start", 3, "end", 4))
    tokens.push CreateDict("symbol", "Add", "text", "+", "range", CreateDict("start", 2, "end", 3))
    tokens.push CreateDict("symbol", "Num", "text", "2", "range", CreateDict("start", 3, "end", 4))
    tokens.push CreateDict("symbol", "Add", "text", "+", "range", CreateDict("start", 2, "end", 3))
    tokens.push CreateDict("symbol", "Num", "text", "2", "range", CreateDict("start", 3, "end", 4))
    tokens.push CreateDict("symbol", "Add", "text", "+", "range", CreateDict("start", 2, "end", 3))
    tokens.push CreateDict("symbol", "Num", "text", "2", "range", CreateDict("start", 3, "end", 4))
    tokens.push CreateDict("symbol", "Add", "text", "+", "range", CreateDict("start", 2, "end", 3))
    tokens.push CreateDict("symbol", "Num", "text", "2", "range", CreateDict("start", 3, "end", 4))
    tokens.push CreateDict("symbol", "Add", "text", "+", "range", CreateDict("start", 2, "end", 3))
    tokens.push CreateDict("symbol", "Num", "text", "2", "range", CreateDict("start", 3, "end", 4))
    tokens.push CreateDict("symbol", "Add", "text", "+", "range", CreateDict("start", 2, "end", 3))
    tokens.push CreateDict("symbol", "Num", "text", "2", "range", CreateDict("start", 3, "end", 4))
    tokens.push CreateDict("symbol", "Add", "text", "+", "range", CreateDict("start", 2, "end", 3))
    tokens.push CreateDict("symbol", "Num", "text", "2", "range", CreateDict("start", 3, "end", 4))
    tokens.push CreateDict("symbol", "Add", "text", "+", "range", CreateDict("start", 2, "end", 3))
    tokens.push CreateDict("symbol", "Num", "text", "2", "range", CreateDict("start", 3, "end", 4))
    
    
    
    
    Debug.Print Now()
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    For i = 1 To 1
        table = cnf.performCYK(tokens)
    Next
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    Debug.Print Now()
    
    Debug.Print toJSON(table, "2D")
    
    
    
    
    
        
    
    
    
    
    
    Debug.Assert False
End Sub

Private Function toJSON(o As Variant, Optional sTypeName As String) As String
    Dim sOut As String
    Dim i As Long, j As Long
    If sTypeName = "2D" Then
        sOut = "["
        For i = LBound(o, 1) To UBound(o, 1)
            sOut = sOut & "["
            For j = LBound(o, 2) To UBound(o, 2)
                sOut = sOut & toJSON(o(i, j)) & IIf(j = UBound(o, 2), "", ",")
            Next
            sOut = sOut & "]" & IIf(i = UBound(o, 1), "", ",")
        Next
        sOut = sOut & "]"
    Else
        If TypeName(o) = "Dictionary" Then
            sOut = "{"
            Dim vKey As Variant
            Dim vKeys As Variant
            vKeys = o.keys()
            For i = LBound(vKeys) To UBound(vKeys)
                vKey = vKeys(i)
                sOut = sOut & """" & vKey & """:" & toJSON(o(vKey)) & IIf(i = UBound(vKeys), "", ",")
            Next
            sOut = sOut & "}"
        ElseIf VarType(o) = vbString Then
            sOut = """" & Replace(Replace(o, "\", "\\"), """", "\""") & """"
        ElseIf TypeName(o) = "stdArray" Then
            sOut = "["
            For i = 1 To o.Length
                sOut = sOut & toJSON(o.item(i)) & IIf(i = o.Length, "", ",")
            Next
            sOut = sOut & "]"
        ElseIf VarType(o) = vbObject Then
            If o Is Nothing Then
                sOut = sOut & "null"
            End If
        ElseIf o = Empty Then
            Select Case VarType(o)
                Case vbLong, vbInteger
                    sOut = "0"
                Case Else
                    sOut = "null"
            End Select
        Else
            sOut = o
        End If
    End If
    
    toJSON = sOut
End Function


Private Function CreateCNFP(left As String, Optional right As Variant = Empty, Optional metadata As Variant = Nothing) As Object
    Set CreateCNFP = CreateDict("left", left, "right", right, "metadata", metadata)
End Function
Private Function CreateDict(ParamArray v() As Variant) As Object
    Set CreateDict = CreateObject("Scripting.Dictionary")
    For i = LBound(v) To UBound(v) Step 2
        CreateDict.ADD v(i), v(i + 1)
    Next
End Function
