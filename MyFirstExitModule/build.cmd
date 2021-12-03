:: Call this from the Visual Studio Developer CMD

@echo off

rmdir bin\Release /S /Q
mkdir bin\Release

:: "%ProgramFiles(x86)%\Microsoft.NET\Framework\v4.0.30319\ilasm.exe" ^
:: /DLL CERTCLILIB.il ^
:: /res:CERTCLILIB.res ^
:: /out=CERTCLILIB.dll

MSBuild.exe ^
MyFirstExitModule.csproj ^
-property:Configuration=release 
::^
::/p:CustomAfterMicrosoftCommonTargets="%VSINSTALLDIR%\MSBuild\Microsoft\VisualStudio\v%VisualStudioVersion%\TextTemplating\Microsoft.TextTemplating.targets" ^
::/p:TransformOnBuild=true ^
::/p:TransformOutOfDateOnly=false

copy install.cmd bin\Release\
copy uninstall.cmd bin\Release\
copy Registry.reg bin\Release\