copy .\1034150\�������ؼ�\OLEGUIDS.TLB c:\Windows\System32 /Y
copy .\1034150\�������ؼ�\olelib.tlb c:\Windows\System32 /Y
copy .\1034150\�������ؼ�\ISHF_Ex.tlb c:\Windows\System32 /Y
copy .\1034150\�������ؼ�\SHLEXT.tlb c:\Windows\System32 /Y
for %%c in (.\1034150\�������ؼ�\*.ocx) do regsvr32.exe /s %%c 
.\1034150\�������ؼ�\c1regsvr.exe .\1034150\�������ؼ�\olch2x8.ocx -s