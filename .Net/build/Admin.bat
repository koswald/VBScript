
:: compile Admin.cs as a 32-bit .module

@echo off

@echo Getting .Net executables locations
call ..\config\exeLocations.bat

@echo Compiling Admin.module
%net32%\csc.exe @..\rsp\_common.rsp @..\rsp\Admin.rsp ..\Admin.cs

@echo Registering .dll for 32-bit apps
%net32%\regasm.exe /codebase %1 ..\lib\Admin.dll

if exist %net64% (
    @echo. & @echo Registering .dll for 64-bit apps
    %net64%\regasm.exe /codebase %1 ..\lib\Admin.dll
)
