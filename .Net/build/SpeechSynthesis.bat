
:: compile SpeechSynthesis.cs as a 32-bit library
:: and register SpeechSynthesis.dll for 64-bit and 32-bit apps

@echo off

@echo Getting .Net executables locations
call ..\config\exeLocations.bat

@echo Compiling SpeechSynthesis.cs
%net32%\csc.exe @..\rsp\_common.rsp @..\rsp\SpeechSynthesis.rsp ..\SpeechSynthesis.cs

@echo Registering .dll for 32-bit apps
%net32%\regasm.exe /codebase %1 ..\lib\SpeechSynthesis.dll

@echo. & @echo Registering .dll for 64-bit apps
%net64%\regasm.exe /codebase %1 ..\lib\SpeechSynthesis.dll
