'Usage:
'
'Dim v as variant
'For each v in Fibonacci.GetEnum()
'  Debug.Print v
'Next

Private pEnumProvider as stdEnumProvider
Private vOld as long
Private v as long
Private v2 as long

Public Function GetEnum() as stdEnumProvider
  if pEnumProvider is nothing then
    set pEnumProvider = stdEnumProvider.Create()
    v  = 0
    v2 = 1
    pEnumProvider.value = v+v2 'initial value
  end if

  set GetEnum = pEnumProvider
End Function


Public Function pEnumProvider_Reset(ByRef value as variant)
  v  = 0
  v2 = 1
  value = v+v2 
End Function
Public Function pEnumProvider_Next(ByVal counter as long, ByRef value as variant, ByRef hasNext as boolean)
  vOld = v
  v = v2
  v2 = v2 + vOld
  value = v2
  hasNext = true
End Function
Public Function pEnumProvider_Skip(ByVal counter as long, ByRef value as variant, ByRef hasNext as boolean)
  pEnumProvider_Next(counter,value,hasNext)
End Function

