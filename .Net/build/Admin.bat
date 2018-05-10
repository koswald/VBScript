:: compile Admin.cs as a 32-bit .dll
@echo off
call ..\config\exeLocations.bat

if %1.==/unregister. (
    set verb=Unregistering
) else (
    set verb=Registering
    echo. & echo Compiling Admin.cs
    %net32%\csc.exe @..\rsp\_common.rsp @..\rsp\Admin.rsp ..\Admin.cs
)

echo %verb% Admin.dll for 32-bit apps
%net32%\regasm.exe /codebase %1 ..\lib\Admin.dll

if exist %net64% (
    echo. & echo %verb% Admin.dll for 64-bit apps
    %net64%\regasm.exe /codebase %1 ..\lib\Admin.dll
)
