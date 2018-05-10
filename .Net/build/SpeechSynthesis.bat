:: compile SpeechSynthesis.cs as a 32-bit library
:: and register SpeechSynthesis.dll for 64-bit and 32-bit apps
@echo off
call ..\config\exeLocations.bat

if %1.==/unregister. (
    set verb=Unregistering
) else (
    set verb=Registering
    echo. & echo Compiling SpeechSynthesis.cs
    %net32%\csc.exe @..\rsp\_common.rsp @..\rsp\SpeechSynthesis.rsp ..\SpeechSynthesis.cs
)
echo %verb% SpeechSynthesis.dll for 32-bit apps
%net32%\regasm.exe /codebase %1 ..\lib\SpeechSynthesis.dll

if exist %net64% (
    echo. & echo %verb% SpeechSynthesis.dll for 64-bit apps
    %net64%\regasm.exe /codebase %1 ..\lib\SpeechSynthesis.dll
)
