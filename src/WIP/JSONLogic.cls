'This is a test library for JSON Logic.
'JSON Logic is being used as a placeholder for Lambdas at the moment, until a proper lamda syntax is released.
'Lambda syntax will likely be converted to JSONLogic objects where they'll be parsed by this algorithm.



'Example:
'  {
'    "==":[1,1]
'  }
'evaluates 1==1 and returns True
'
'In VBA construction of such a dictionary will likely require stdJSON.cls along the lines of:
'  jsonLogic.call(stdJSON.parse("{""=="":[1,1]}"))
'
'
'jsonLogic will implement stdICallable so it can be used as callbacks.
'
'

implements stdICallable


'Constructor
Public Function Create(optional ByVal Obj as object = nothing) as stdJSONLogic
  Dim oObj as object
  set oObj = new stdJSONLogic
  Call oObj.Init(obj)
  set Create = oObj
End Function
Friend Sub Init(ByVal Obj as object) as stdJSONLogic
  set pState = Obj
End Sub

'Usual public call function
Public Function Run(ParamArray params() as variant) as variant
  CopyVariant Run, ExecuteJSON(params)
End Function

'stdICallable Wrappers
Private Function stdICallable_Run(ParamArray params() as variant) as variant
  CopyVariant stdICallable_Run, ExecuteJSON(params)
End Function

Private Function stdICallable_GetPointer() as Long
  stdICallable_GetPointer = stdPointer.GetLastPrivateMethod(me)
End Function

Private Function CopyVariant(ByRef a as variant, ByVal b as variant)
  if isObject(a) then
    set a = b
  else
    a = b
  end if
End Function

Private Function ExecuteJSON(params() as variant) as variant

  

End Function


Public Sub SetupFunctions()
  Dim obj as object
  set obj = CreateObject("Scripting.Dictionary")
  
End Sub




'OPERATIONS:
'==
Private Function Eql_(ByRef a as variant,ByRef b as variant) as Boolean
  if isObject(a) and isObject(b) then
    Eql_ = a is b
  elseif not isObject(a) and not isObject(b) then
    Eql_ = a = b
  end if
End Function

'===
Private Function Eql2_(ByRef a as variant,ByRef b as variant) as Boolean
  if typename(a)=typename(b) then
    Eql2_ = Eql_(a,b)
  end if
End Function

'!=
Private Function NEql_(ByRef a as variant,ByRef b as variant) as Boolean
  NEql_ = not Eql_(a,b)
End Function

'!==
Private Function NEql2_(ByRef a as variant,ByRef b as variant) as Boolean
  NEql2_ = not Eql2_(a,b)
End Function

'>
Private Function Gt_(ByRef a as variant,ByRef b as variant) as Boolean
  Gt_ = a>b
End Function

'>=
Private Function Gte_(ByRef a as variant,ByRef b as variant) as Boolean
  Gte_ = a>=b
End Function

Private Function GetMissing(Optional missing as variant) as variant
  GetMissing = missing
End Function

'<
Private Function Lt_(ByRef a as variant,ByRef b as variant, optional ByRef c as variant) as Boolean
  if c = GetMissing() then
    Lt_ = a<b
  else
    Lt_ = a<b and b<c
  end if
End Function

'<=
Private Function Lte_(ByRef a as variant,ByRef b as variant) as Boolean
  if c = GetMissing() then
    Lte_ = a<=b
  else
    Lte_ = a<=b and b<=c
  end if
End Function

'!!
Private Function Truthy_(ByRef a as variant)
  Truthy_ = CBool(a)
End Function

'!
Private Function Not_(ByRef a as variant)
  Not_ = not CBool(a)
End Function

'%
Private Function Mod_(ByVal a as Double, ByVal b as Double) as Double
  Mod_ = a Mod b
end Function

'log
Private function Log_(a) as variant
  Debug.Print a
  Log_ = a
End Function

'in          'is 
Private Function In_(ByRef a as variant,ByRef b as variant)
  'Not really sure how to implement in
End Function

'cat
Private Function Cat_(a() as variant) as string
  Dim s as string
  Dim i as long 
  For i = lbound(a) to ubound(a)
    s = s & a(i)
  Next i
End Function

'substr
Private Function Substr_(ByVal source as variant, ByVal start as long, ByVal end as long) as string
  if end < 0 then
    temp = mid(source,start,end-start)
  end if
End Function

'+ (sum)
Private Function Sum_(arr() as double) as double
  Dim x as double, i as long
  For i = lbound(arr) to ubound(arr)
    x = x + arr(i)
  next
  Sum_ = x
End Function

'* (product)
Private Function Product_(arr() as double) as double
  Dim x as double, i as long
  For i = lbound(arr) to ubound(arr)
    x = x * arr(i)
  next
  Product_ = x
End Function

'-
Private Function Subtract_(ByVal a as double, Optional ByVal b as double) as double
  if b = GetMissing() then
    Subtract_ = -1*a
  else
    Subtract_ = a - b
  end if
End Function

'/
Private Function Divide_(ByVal a as double, ByVal b as double) as double
  Divide_ = a/b
End Function

'min
Private Function Min_(arr() as variant) as double
  Dim x as double, Dim i as long
  x = arr(lbound(arr))
  if ubound(arr)-lbound(arr)+1>0 then
    for i = lbound(arr)+1 to ubound(arr)
      if x > arr(i) then
        x = arr(i)
      end if
    next
  end if

  Min_ = x
End Function

'max
Private Function Max_(arr() as variant) as double
  Dim x as double, Dim i as long
  x = arr(lbound(arr))
  if ubound(arr)-lbound(arr)+1>0 then
    for i = lbound(arr)+1 to ubound(arr)
      if x < arr(i) then
        x = arr(i)
      end if
    next
  end if

  Max_ = x
End Function

'merge
'--> flatten array    --> [1,2,[3,4,[5,6]]] ==> [1,2,3,4,[5,6]]
Private Function Merge_(arr() as variant) as variant
  'Calculate length
  Dim i as long, newLength as long
  newLength = ubound(arr)-lbound(arr)+1
  
  'Calculate new length
  For i = lbound(arr) to ubound(arr)
    if isArray(arr(i)) then
      newLength = newLength + ubound(arr(i))-lbound(arr(i))+1
    end if
  next

  Dim newArr(1 to newLength) as variant
  '...
End Function


'var
Private Function Var_(ByVal varPath as string, ByVal default as variant)

End Function

'missing


'missing_some


'method