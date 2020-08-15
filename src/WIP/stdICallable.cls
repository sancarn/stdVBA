'Call will call the passed function with param array
Public Function Run(ParamArray params() as variant) as variant

'Call function with supplied array of params
Public Function RunEx(params() as variant) as variant

'Returns a callback function
'Typically this will be achieved with `stdPointer.GetLastPrivateMethod(me)`
'If this cannot be implemented return 0
Public Function ToPointer() as long






