:: compile ComEvent.cs as a 32-bit .module
@echo off
call ..\config\exeLocations.bat

if Not %1.==/unregister. (
    echo. & echo Compiling ComEvent.cs
    %net32%\csc.exe @..\rsp\_common.rsp @..\rsp\ComEvent.rsp ..\ComEvent.cs
)