:: compile NotifyIcon.cs as a 32-bit library
:: and register NotifyIcon.dll for 64-bit and 32-bit apps
@echo off
call ..\config\exeLocations.bat

if %1.==/unregister. (
    set verb=Unregistering
) else (
    set verb=Registering
    echo. & echo Compiling NotifyIcon.cs
    %net32%\csc.exe @..\rsp\_common.rsp @..\rsp\NotifyIcon.rsp ..\NotifyIcon.cs
)
echo %verb% NotifyIcon.dll for 32-bit apps
%net32%\regasm.exe /codebase %1 ..\lib\NotifyIcon.dll

if exist %net64% (
    echo. & echo %verb% NotifyIcon.dll for 64-bit apps
    %net64%\regasm.exe /codebase %1 ..\lib\NotifyIcon.dll
)
