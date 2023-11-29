VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdICallable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = false
Attribute VB_Exposed = False

'Run the function with supplied parameters
'@param params - The parameters to run the function with
'@returns      - The result of the function
Public Function Run(ParamArray params() as variant) as variant: End Function

'Run the function with a array of parameters
'@param params as Variant<Array<Variant>> - The parameters to run the function with
'@returns                                 - The result of the function
Public Function RunEx(ByVal params as variant) as variant: End Function

'Bind a set of parameters to the function call
'@param params - The parameters to bind to the function
'@returns      - A new function with the parameters bound
Public Function Bind(ParamArray params() as variant) as stdICallable: End Function

'Making late-bound calls to `stdICallable` members. Each object which implements `stdICallable`
'will support a different set of latebound calls.
'@protected
'@param sMessage - Message to send. Standard messages include "obj" returning the object, "className" returning the class name. Other messages are implementation specific.
'@param success  - Whether the call was successful
'@param params   - Any variant, typically parameters as an array. Passed along with the message.
'@returns - Any return value.
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


