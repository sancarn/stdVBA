# stdWebSocket

## Spec

```vb
'  CONSTRUCTORS
'    [ ] Create
'
'  INSTANCE METHODS
'  Many methods were inspired by those in Ruby's Enumerable: https://ruby-doc.org/core-2.7.2/Enumerable.html
'    [ ] send(data)
'    [ ] close(code,reason)
'    [ ] disconnect()
'    [ ] Get url()
'  PROTECTED INSTANCE METHODS
'    [ ] handleEvent()
'  EVENTS
'    [ ] OnOpen(data)
'    [ ] OnClose(data)
'    [ ] OnError(data)
'    [ ] OnMessage(data)
'    [ ] EventRaised(name,data)
'TODO: IDEALLY WE'D DO EVERYTHING ASYNCHRONOUSLY WITH EVENTS, HOWEVER THIS ISN'T DOABLE UNTIL WE GET A OBJECT CALLER THUNK. EVENTS WE'D SUPPORT:
'  OnOpen(ByVal eventData As Variant)
'  OnClose(ByVal eventData As Variant)
'  OnError(ByVal eventData As Variant)
'  OnMessage(ByVal eventData As Variant)
```