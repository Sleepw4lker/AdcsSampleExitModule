@echo off
set DLL=MyFirstExitModule.dll
net stop certsvc
%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe /unregister %SystemRoot%\System32\%DLL%
del %SystemRoot%\System32\%DLL% /Q
del %SystemRoot%\SysWOW64\%DLL% /Q
xcopy %DLL% %SystemRoot%\System32\.
xcopy %DLL% %SystemRoot%\SysWOW64\.
%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe %SystemRoot%\System32\%DLL%
net start certsvc