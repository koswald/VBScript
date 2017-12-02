
:: compile ComEvent.cs as a 32-bit .module

@echo off

@echo Getting .Net executables locations
call ..\config\exeLocations.bat

@echo Compiling ComEvent.module
%net32%\csc.exe @..\rsp\_common.rsp @..\rsp\ComEvent.rsp ..\ComEvent.cs
