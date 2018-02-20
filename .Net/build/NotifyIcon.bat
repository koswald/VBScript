
:: compile NotifyIcon.cs as a 32-bit library
:: and register NotifyIcon.dll for 64-bit and 32-bit apps

@echo off

@echo Getting .Net executables locations
call ..\config\exeLocations.bat

if %1.==/unregister. (
    set verb=Unregistering
) else (
    set verb=Registering
    @echo Compiling NotifyIcon.cs
    %net32%\csc.exe @..\rsp\_common.rsp @..\rsp\NotifyIcon.rsp ..\NotifyIcon.cs
)
@echo %verb% .dll for 32-bit apps
%net32%\regasm.exe /codebase %1 ..\lib\NotifyIcon.dll

if exist %net64% (
    @echo. & @echo %verb% .dll for 64-bit apps
    %net64%\regasm.exe /codebase %1 ..\lib\NotifyIcon.dll
)
