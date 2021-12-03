@echo off
set DLL=MyFirstExitModule.dll
net stop certsvc
%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe /unregister %SystemRoot%\System32\%DLL%
del %SystemRoot%\System32\%DLL% /Q
del %SystemRoot%\SysWOW64\%DLL% /Q
net start certsvc