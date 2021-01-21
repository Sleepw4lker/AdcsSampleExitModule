@echo off
set DLL=MyFirstExitModule.dll
net stop certsvc
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe /unregister C:\Windows\System32\%DLL%
del C:\Windows\System32\%DLL% /Q
del C:\Windows\SysWOW64\%DLL% /Q
xcopy %DLL% C:\Windows\System32\.
xcopy %DLL% C:\Windows\SysWOW64\.
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe C:\Windows\System32\%DLL%
net start certsvc
start certsrv.msc
pause