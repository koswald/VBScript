:: compile IconExtractor.cs as a 32-bit library
:: and register IconExtractor.dll for 64-bit and 32-bit apps
@echo off
call ..\config\exeLocations.bat

if %1.==/unregister. (
    set verb=Unregistering
) else (
    echo. & echo Compiling IconExtractor.dll
    %net32%\csc.exe @..\rsp\_common.rsp @..\rsp\IconExtractor.rsp ..\IconExtractor.cs
    set verb=Registering
)
echo %verb% IconExtractor.dll for 32-bit apps
%net32%\regasm.exe /codebase %1 ..\lib\IconExtractor.dll

if exist %net64% (
    echo. & echo %verb% IconExtractor.dll for 64-bit apps
    %net64%\regasm.exe /codebase %1 ..\lib\IconExtractor.dll
)
