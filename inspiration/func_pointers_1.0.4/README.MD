# BrickMod

Made by [FireHacker](http://bbs.vbstreets.ru/memberlist.php?mode=viewprofile&u=8077&sid=696b6b7d3f26b413017e819dcb296cf1) on [Russian forums](http://bbs.vbstreets.ru/viewtopic.php?f=28&t=42929).

This code requires `vba6.dll` which is not supplied with Office. It comes with the VB6 IDE and I'm uncertain whether this would even work in VBA. If it does, then the DLL has some fairly amazing capabilities! This particular project enumerates the executing VBProject and harvests all callable functions by enumerating the VBProject's type library.

# Other coolness with vba6.dll

## [`EbExecuteLine`](https://www.tek-tips.com/viewthread.cfm?qid=97332)

Allows you to execute code as if it were written in the immediate window!