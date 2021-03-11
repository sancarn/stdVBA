#General overlook

## Real VB6 Threading

...

## Simulated threading (might go under Runtimes/WScript instead):

```vb
'REQUIRES REGISTRY:
  '  [HKLM\SOFTWARE\Microsoft\Windows Script Host\Settings]
  '  Remote  REG_SZ  1
  '
  Dim ws As Object
  Set ws = CreateObject("WSHController")
  Set RemoteScript = ws.Run("C:\Users\jwa\Desktop\TBD\test.vbs")
```
