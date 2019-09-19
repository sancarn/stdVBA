

--> MAKE SURE YOU REGISTER THE TLB FILE <--

If you down know how..google for "tlb reggie register" reggie is a free
tlb registration program you can find on the net..adds right click option 
for it.


This is a VB only implementation of IActiveScript which lets you integrate 
scripting support in your apps without the need for the MSScript control.

I started working on this because I wanted to work torwards adding full
IActiveScript support including debugging because the MS Script control
sometimes isnt enough.

Anyway, this is the first step. It may not be 100% perfectly implemented 
but it runs and works with objects you pass to it. 

You are free to use it in any commercial or non commercial applications 
as you so desire. Only a brief line "This product contains software written by
David Zimmer" is required.

Enjoy


ps - if you want to see a C activex control that you can use from VB with
debugging support check out ken fousts citrus debugger 

http://sandsprite.com/CodeStuff/CitrusDebugger.7z

For the brave of heart, I also started to try implementing the iActiveScript
debug interfaces directly in VB. You can get a copy here:

http://sandsprite.com/CodeStuff/vbActiveScript_wDbg_incomplete.zip

Currently I have ditched the MS Script engine completely and am now using
DukTape javascript engine 1mb self contained complete with debug support.
search my site or web for duk4vb.
