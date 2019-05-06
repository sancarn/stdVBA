# The trick Advanced tools

Code developed by The Trick [vbforums](http://www.vbforums.com/member.php?246405-The-trick) [ruCyberForum](http://www.cyberforum.ru/blogs/354370/)

Releasing post: http://www.vbforums.com/showthread.php?835951-VB6-Trick-Advanced-Tools

I present to you a small project
Add-in that allows to some extent alleviate debugging some programs as well expand compilation possibilities.

This Add-in has the following features:
* Fixes the bug with Not Not Array statement that causes error "Expression too complex" if you'll work with float numbers;
* Allows the ability to utilize the automatic conditional constants depending on run mode (IDE/EXE) look like in C++ (NDEBUG);
* Allows the ability to disable integer overflow checking in IDE;
* Allows the ability to disable floating point result checking in IDE;
* Allows the ability to disable array bounds checking in IDE;
* Provides the compilation/linking events (both in IDE and EXE), i.e. you can run the commands before/after this events. By using this events you can do many useful things (encryption, replace OBJ files, static linking, etc.)
