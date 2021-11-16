:: Call this from the Visual Studio Developer CMD

@echo off

rmdir bin\Release /S /Q
mkdir bin\Release

MSBuild.exe ^
MyFirstExitModule.csproj ^
-property:Configuration=release 
::^
::/p:CustomAfterMicrosoftCommonTargets="%VSINSTALLDIR%\MSBuild\Microsoft\VisualStudio\v%VisualStudioVersion%\TextTemplating\Microsoft.TextTemplating.targets" ^
::/p:TransformOnBuild=true ^
::/p:TransformOutOfDateOnly=false

copy deploy.cmd bin\Release\
copy Registry.reg bin\Release\