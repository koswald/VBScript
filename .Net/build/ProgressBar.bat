:: compile ProgressBar.cs as a 32-bit library
:: and register ProgressBar.dll for 64-bit and 32-bit apps
@echo off
call ..\config\exeLocations.bat

if %1.==/unregister. (
    set verb=Unregistering
) else (
    set verb=Registering
    echo. & echo Compiling ProgressBar.cs
    %net32%\csc.exe @..\rsp\_common.rsp @..\rsp\ProgressBar.rsp ..\ProgressBar.cs
)
echo %verb% ProgressBar.dll for 32-bit apps
%net32%\regasm.exe /codebase %1 ..\lib\ProgressBar.dll

if exist %net64% (
    echo. & echo %verb% ProgressBar.dll for 64-bit apps
    %net64%\regasm.exe /codebase %1 ..\lib\ProgressBar.dll
)
