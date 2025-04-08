If opening the project Vb6Tkinter.vbp fails with a prompt saying "Object library not registered," it is usually because the Windows Common Controls 6.0 (mscomctl.ocx) has not been successfully loaded.
You can first try registering it with regsvr32 mscomctl.ocx.
If that still doesn't work, you can try regtlib msdatsrc.tlb.

As for the location of mscomctl.ocx or msdatsrc.tlb, it varies depending on the version. You can search for it to locate the files.

32bit:
cd c:\windows\system32
regtlib msdatsrc.tlb

64bit:
cd C:\Windows\SysWOW64\
regtlib msdatsrc.tlb

Translate Chinese remarks to English.
Keep the whole text without line numbers.
Delete the Chinese remarks after translation.