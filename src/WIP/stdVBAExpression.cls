VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBAExpression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'A VBA wrapper for Excel formulae
'Allows accessing of registered object data

'Example:
'  Dim f_x as VBAExpression
'  set f_x = VBAExpression.Create()
'
'  set f_x.RegisteredObjects["sheet"] = ActiveSheet
'  f_x.Formula = "#sheet.UsedRange.rows.count"
'
'  ==> Count of rows in used range of ActiveSheet
'
'  Public Functions in modules are also of course available within the VBA Expression.
'  As with formulae, the module name can also be included:
'
'    f_x.Formula = "VBAExpressionTesting.TestPublicFunc()"
'    ?f_x() '=> True
'
'  Calling functions on objects is not supported. E.G.
'
'    f_x.Formula = "#obj.SomeFunc()"
'
'  As a work around, use a Public module function:
'
'    Module MyModule
'      Function SomeFunc(obj as object)
'        SomeFunc = obj.SomeFunc()
'      End Function
'    End Module
'    ...
'    f_x.Formula = "MyModule.SomeFunc(#obj)"

Const ChildrenProperty = "Children"


Public formula As String
Public RegisteredObjects As Object
Private pInitialised As Boolean
Private isCompiled As Boolean
Private AST As Object
'Private Compiled_Parts As Collection

Public Function Create(Optional formula As String, Optional RegisteredObjects As Object) As VBAExpression
  Set Create = New VBAExpression
  Call Create.init(formula, RegisteredObjects)
End Function
Public Function Evaluate() As Variant
Attribute Evaluate.VB_UserMemId = 0
  If Not isCompiled Then
    Me.Compile False
  End If
  
  Dim part As Object, expression As String
  For Each part In AST("Children")
    Select Case part("Type")
      Case "Literal"
        expression = expression & part("value")
      Case "Macro"
        'get initial object from name
        Dim subPart As Object, obj As Object
        Dim childCount As Long: childCount = part("children").Count
        
        
        For i = 1 To childCount
          If i = 1 Then
            Set obj = RegisteredObjects(part("Children")(i))
          Else
            'if getReturnType(...) = vbObject then
            '  set obj = ...
            'else
            '  obj = ...
            'end if
          End If
        Next
        
        expression = expression & value
    End Select
  Next
  
  'Evaluate equation:
  Evaluate = Application.Evaluate(expression)
End Function


Public Sub Compile(Optional setCompiled As Boolean = True)
  Set AST = CompileToAST(Me.formula)
  isCompiled = setCompiled
End Sub

'New compile is going to do the following:
'    Imagine the following formula:
'        IF("hel""lo #world""",True,#do.diddle,#do.something.cool("a#sdf",1))
'    In this case we want to scan all the way up to #do
'    Parse previous into a literal
'    Continue parsed #do until we find a "," (or function boundary) while at iBracketScope = 0
'    Pass this section off to be parsed seperately (by ParseVBAExpression), and then continue onwards as a literal.
'Result will look something like this:
'  {
'    Type: "AST",
'    Children: [
'      {Type:"Literal",Value:"IF("},
'      {Type:"Literal",Value:"\"hel\"\"lo #world\"\""},
'      {Type:"Literal",Value:",True,"},
'      {Type:"Macro",  Children:[
'        {Type:"EntryPoint", Value:"do"},
'        {Type:"Getter", Value:"diddle", Params:[]}
'      ]},
'      {Type:"Literal",Value:","},
'      {Type:"Macro",  Children:[
'        {Type:"EntryPoint", Value:"do"},
'        {Type:"Getter", Value:"something", Params:[]},
'        {Type:"Getter", Value:"cool", Params:[
'          {Type:"Literal", Value:"\"a#sdf\",1"}
'        ]}
'      ]}
'      {Type:"Literal",Value:")"},
'    ]
'  }
'Example:
'?VBAExpression.CompileToAST("IF(""hel""""lo #world"""""",True,#do.diddle,#do.something.cool(""a#sdf"",1))")("Children").Count
Function CompileToAST(formula As String) As Object
  'formula as byte array
  Dim sFormula() As Byte
  sFormula = formula
  
  'Upper bound of array
  Dim iFormulaLen As Long
  iFormulaLen = Len(formula)
  
  'If Formula is blank then exit the function
  If iFormulaLen = 0 Then
    Exit Function
  End If
  
  'Create an Abstract syntax tree to store compilation structure
  Dim AST As Object
  Set AST = CreateCompiledPart("AST")
  
  'Temporary compiled part for use
  Dim TmpCompiledPart As Object
  
  'Test for formula is guaranteed literal
  If InStr(1, formula, "#") = 0 Then
    AST("Children").add CreateLiteralPart(formula, AST)
    Set CompileToAST = AST
    Exit Function
  End If
  
  'Constants used for Compiling fast
  Dim iHashChar As Integer: iHashChar = Asc("#")
  Dim iQuoteChar As Integer: iQuoteChar = Asc("""")
  Dim iBracketCharO As Integer: iBracketCharO = Asc("(")
  Dim iBracketCharC As Integer: iBracketCharC = Asc(")")
  Dim iCommaChar As Integer: iCommaChar = Asc(",")
  
  'Keeping track of parts
  Dim sLiteral As String
  Dim iBracketScope As Integer
  
  'Loop through string
  Dim iChar As Long, char As Integer
  Dim bIsVBAExpression As Boolean
  For i = 1 To iFormulaLen
    'Get char position and bytes
    iChar = (i - 1) * 2
    char = sFormula(iChar)
    
    'Detect quotes
    If char = iQuoteChar Then
      If bString Then
        'If next char is also iQuoteChar, append and skip 2 chars
        If sFormula(iChar + 2) = iQuoteChar Then
          sLiteral = sLiteral & """"""
          i = i + 1
          GoTo NextChar
        Else
          'In this case we need to end the string:
          bString = False
          sLiteral = sLiteral & """"
          AST("Children").add CreateLiteralPart(sLiteral, AST)
          sLiteral = ""
        End If
        
        GoTo NextChar
      Else
        'Append existing sLiteral if exists
        If sLiteral <> "" Then
          AST("Children").add CreateLiteralPart(sLiteral, AST)
        End If
        
        'Begin sLiteral
        bString = True
        sLiteral = """"
        GoTo NextChar
      End If
    End If
    
    'If is string then...
    If bString Then
      sLiteral = sLiteral & Chr(char)
      GoTo NextChar
    End If
    
    Dim j As Integer
    If char = iHashChar Then
      'Append existing sLiteral if exists
      If sLiteral <> "" Then
        AST("Children").add CreateLiteralPart(sLiteral, AST)
        sLiteral = ""
      End If
      
      'Capture 0 depth VBA expressions
      For j = i + 1 To iFormulaLen
        iChar = (j - 1) * 2
        char = sFormula(iChar)
        If Not bString Or char = iQuoteChar Then
          Select Case char
            Case iQuoteChar
              If bString Then
                If (sFormula(iChar + 2) = iQuoteChar) Then
                  j = j + 1
                Else
                  bString = False
                End If
              Else
                bString = True
              End If
            Case iBracketCharO
              iBracketScope = iBracketScope + 1
            Case iBracketCharC
              'If already out of scope, end expression.
              If iBracketScope = 0 Then GoTo VBAExpressionEnd
              
              'Else continue
              iBracketScope = iBracketScope - 1
            Case iCommaChar
              'If already out of scope, end expression.
              If iBracketScope = 0 Then GoTo VBAExpressionEnd
          End Select
        End If
      Next
      
VBAExpressionEnd:
      AST("Children").add CreateMacroPart(Mid(formula, i, j - i))
      i = j
    End If
    
    'If nothing else, append literal
    sLiteral = sLiteral & Chr(char)
NextChar:
  Next
  
  'Final literal part
  If sLiteral <> "" Then
    AST("Children").add CreateLiteralPart(sLiteral, AST)
  End If
  
  'Optimise and return AST
  Set CompileToAST = OptimiseAST(AST)
End Function

Private Sub Class_Initialize()
  Set RegisteredObjects = CreateObject("Scripting.Dictionary")
End Sub
Friend Sub init(Optional formula As String, Optional RegisteredObjects As Object)
  If pInitialised Then
    Err.Raise 1, "VBAExpression::Init", "Cannot initialise initialised object"
  Else
    Me.formula = formula
    If Not RegisteredObjects Is Nothing Then Set Me.RegisteredObjects = RegisteredObjects
    pInitialised = True
  End If
End Sub

Private Function CreateCompiledPart(TypeName As String) As Object
  Set NewCompiledPart = CreateObject("Scripting.Dictionary")
  NewCompiledPart("Type") = TypeName
  Set NewCompiledPart("Children") = New Collection
  Set CreateCompiledPart = NewCompiledPart
End Function

Private Function CreateLiteralPart(value As String, Parent As Object) As Object
  'Create literal part
  Dim TmpCompiledPart As Object
  Set TmpCompiledPart = CreateCompiledPart("Literal")
  Set TmpCompiledPart("Parent") = Parent
  TmpCompiledPart("Value") = value
  
  'Return Literal
  Set CreateLiteralPart = TmpCompiledPart
End Function


Private Function CreateMacroPart(sToCompile As String) As Object
  'Macro part to return
  Dim MacroPart As Object
  Set MacroPart = CreateCompiledPart("Macro")
  
  'Parse VBA:
  '=======================
  'Split based on 0 depth "."
  Dim iLen As Long: iLen = Len(sToCompile)
  
  Dim char As String
  Dim iScope As Long
  Dim bString As Boolean
  Dim i As Integer
  Dim splitLoc As Long: splitLoc = 1
  Dim iPath As Integer: iPath = 0
  
  Dim parts() As String
  For i = 1 To iLen
    char = Mid(sToCompile, i, 1)
    If Not bString Or char = """" Then
      Select Case char
        Case "("
          iScope = iScope + 1
        Case ")"
          iScope = iScope - 1
          If iScope = -1 Then
            'ERROR - Syntax error!
            Err.Raise 1, "VBAExpression::CreateMacroPart", "SYNTAX ERROR - Found Closing bracket "")"" out of scope"
          End If
        Case "."
          If iScope = 0 Then
            'Make parts array...
            ReDim Preserve parts(0 To iPath)
            parts(iPath) = Mid(sToCompile, splitLoc, i - splitLoc)
            
            'Make parts array...
            iPath = iPath + 1
            splitLoc = i + 1
          End If
        Case """"
          If bString Then
            If Mid(sToCompile, i + 1, 1) = """" Then
              i = i + 1
            Else
              bString = False
            End If
          Else
            bString = True
          End If
      End Select
    End If
  Next
  
  If splitLoc < iLen Then
    ReDim Preserve parts(0 To iPath)
    parts(iPath) = Mid(sToCompile, splitLoc, iLen - splitLoc + 1)
  End If
  
  'First part is entry point
  MacroPart("Children").add CreateEntryPointPart(Mid(parts(0), 2), MacroPart)
  
  'All other parts will be getters with the and will occasionally have parameters
  iLen = UBound(parts) - LBound(parts) + 1
  If iLen > 1 Then
    For i = 1 To iLen
      MacroPart("Children").add CreateGetterPart(parts(i), MacroPart)
    Next
  End If
  
  'Return macro part
  Set CreateMacroPart = MacroPart
End Function

Function CreateEntryPointPart(Identifier As String, ByRef Parent As Object) As Object
  'Create literal part
  Dim TmpCompiledPart As Object
  Set TmpCompiledPart = CreateCompiledPart("EntryPoint")
  Set TmpCompiledPart("Parent") = Parent
  TmpCompiledPart("Value") = Identifier
  
  'Return Literal
  Set CreateEntryPointPart = TmpCompiledPart
End Function

'Create getter:
'For #___.something
'  {Type:"Getter", Value:"something", Params:[]}
'
'For #___.cool("a#sdf",1)
'  {Type:"Getter", Value:"cool", Params:[
'    {Type:"Literal", Value:"\"a#sdf\",1"}
'  ]}
Function CreateGetterPart(Identifier As String, ByRef Parent As Object) As Object
  'Create literal part
  Dim TmpCompiledPart As Object
  Set TmpCompiledPart = CreateCompiledPart("Getter")
  Set TmpCompiledPart("Parent") = Parent
  
  Dim iBracketLocation As Long: iBracketLocation = InStr(1, Identifier, "(")
  
  If iBracketLocation = 0 Then iBracketLocation = Len(Identifier) + 1
  TmpCompiledPart("Value") = Mid(Identifier, 2, iBracketLocation - 2)
  
  'Get parameters if they exist
  Set TmpCompiledPart("Params") = New Collection
  
  Dim sParams As String
  sParams = Mid(Identifier, iBracketLocation)
  
  Dim oAST As Object, oHandle As Object
  Set oAST = CompileToAST(sParams)
  
  For Each oHandle In oAST("Children")
    TmpCompiledPart("Params").add oHandle
  Next
  
  'Return Literal
  Set CreateGetterPart = TmpCompiledPart
End Function

Private Function OptimiseAST(AST As Object) As Object
  'TODO: CompressAST by Joining neighbouring literals
  Set OptimiseAST = AST
End Function

