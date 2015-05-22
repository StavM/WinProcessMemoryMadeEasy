# WinProcessMemoryMadeEasy
Written in: Visual Basic 6.0

This is a class library which allows users to read and modify external processes virtual memory space, without needing to go to deep into
windows internals. (it's basically an extended wrapper to the ReadProcessMemory and WriteProcessMemory APIs).

The idea is to have a DLL, which allows even beginning programmers, who write in almost any windows-compatible language, to read and write straight off an external application's memory address space,
without having to go through the steep learning curve of windows internals, by that handling most of the complex stuff and hiding them from the user.

Usage and examples (22.05.2015 - just finished the reading parts)

Const LNG_MEMORY_ADDRESS_IN_HEX As Long = 0 ' The Memory address you would like to read - EDIT THIS!!!
Dim MemoryReading as NewMemClass
Set MemoryReading = New NewMemClass

MemoryReading.AttachToProcess_ByWindowName("Untitled - Notepad")
Form1.Caption = MemoryReading.Read_String(LNG_MEMORY_ADDRESS_IN_HEX)
MemoryReading.ReleaseProcess

More examples will be added this week.
