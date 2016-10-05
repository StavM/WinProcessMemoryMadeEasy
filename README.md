# WinProcessMemoryMadeEasy
Written in: Visual Basic 6.0

This is a class library which allows users to read and modify external processes virtual memory space, without needing to go to deep into
windows internals. (it's basically an extended wrapper to the ReadProcessMemory and WriteProcessMemory APIs).

The idea is to have a DLL, which allows even beginning programmers, who write in almost any windows-compatible language, to read and write straight off an external application's memory address space,
without having to go through the steep learning curve of windows internals, by that handling most of the complex stuff and hiding them from the user.
Do notice this will only work with 32bit (4bytes) pointers but can easily be worked on to support (8byte) 64bit pointers.

Usage and examples can be found on the \Examples folder.

More examples will be added later on.
