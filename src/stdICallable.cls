VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdICallable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = false
Attribute VB_Exposed = False

'Call will call the passed function with param array
Public Function Run(ParamArray params() as variant) as variant: End Function

'Call function with supplied array of params
Public Function RunEx(ByVal params as variant) as variant: End Function

'Bind a parameter to the function
Public Function Bind(ParamArray params() as variant) as stdICallable: End Function

'Making late-bound calls to stdICallable members
'@protected
'@param {ByVal String} - Message to send
'@param {ByRef Boolean} - Whether the call was successful
'@param {ByVal Variant} - Any variant, typically parameters as an array. Passed along with the message.
'@returns {Variant} - Any return value.
Public Function SendMessage(ByVal sMessage as string, ByRef success as boolean, ByVal params as variant) as Variant: End Function

'Ideally we would want to get a pointer to the function... However, getting a pointer to an object method is
'going to be defficult, partly due to the first parameter sent to the function is `Me`! We'll likely have to
'use machine code to wrap a call with a `Me` pointer just so we can access the full pointer and use this in
'real life applications.
'Finally it might be better to do something more like: `stdPointer.fromICallable()` anyway
  ''Returns a callback function
  ''Typically this will be achieved with `stdPointer.GetLastPrivateMethod(me)`
  ''If this cannot be implemented return 0
  'Public Function ToPointer() as long

''Bind arguments to functions to appear as first arguments in call.
''e.g. stdLambda.Create("$1.EnableEvents = false: $1.ScreenUpdating = false").bind(Application).Run()
'Public Function Bind(ByVal v as variant) as stdICallable: End Function


