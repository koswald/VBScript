
:: compile Timer.cs as a 32-bit library
:: and register Timer.dll for 64-bit and 32-bit apps

@echo off

@echo Getting .Net executables locations
call ..\config\exeLocations.bat

if %1.==/unregister. (
    set verb=Unregistering
) else (
    set verb=Registering
    @echo Compiling Timer.cs
    %net32%\csc.exe @..\rsp\_common.rsp @..\rsp\Timer.rsp ..\Timer.cs
)
@echo %verb% .dll for 32-bit apps
%net32%\regasm.exe /codebase %1 ..\lib\Timer.dll

if exist %net64% (
    @echo. & @echo %verb% .dll for 64-bit apps
    %net64%\regasm.exe /codebase %1 ..\lib\Timer.dll
)
