http://www.vbforums.com/showthread.php?860333-vb6-Enhancing-VB-s-StdPicture-Object-to-Support-GDI
http://www.vbforums.com/showthread.php?882857-VB-SQLite-Library-(COM-Wrapper)


Pointers (sorta implemented into stdPointer v2)
http://www.vbforums.com/showthread.php?886203-vb6-Getting-AddressOf-for-VB-Class-Object-Modules


Thunking Pointer.GetAddressOf() for C Calls (?):
https://www.codeguru.com/cpp/misc/misc/assemblylanguage/article.php/c12667/Thunking-in-Win32.htm

Thunk experiment:
http://www.vbforums.com/showthread.php?818583-Binary-Code-Thunk-Experiment

More Thunk learning (LaVolpe suggested these):
http://www.vbforums.com/showthread.php?875327-vb6-IDE-Safety-Thunks-A-new-breed

Sockets:
https://github.com/wqweto/VbAsyncSocket

Zip
https://github.com/wqweto/ZipArchive

Subclassing windows (more thunks)
http://www.vbforums.com/showthread.php?834333-La-Volpe-s-cSelfSubHookCallback-cls

Thunks LaVolpe Variants
http://www.vbforums.com/showthread.php?549679-RESOLVED-VB6-ASM-and-Variants

More subclassing thunks
http://www.vbforums.com/showthread.php?724529-Anyone-know-what-this-code-is-what-it-does-and-how-it-works

M2000 - Custom Language Interpreter written in VB6
https://github.com/M2000Interpreter/Version9

```
Module Alfa {
      Print @alfa(100) ' 200
      Function alfa(x)
            =x*2
      End Function
}
Module Beta {
      static M=10
      Print @alfa(M), ' 20 18 16 14 12 10 8 6 4
      M--
      If M>1 then Call Beta Else Print
      Clear ' clear static
}
Alfa
Beta
Beta
```

M2000 Fast collection
https://github.com/M2000Interpreter/Version9/blob/master/FastCollection.cls

M2000 Math.cls
https://github.com/M2000Interpreter/Version9/blob/master/Math.cls

