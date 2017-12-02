
:: compile IconExtractor.cs as a 32-bit .module

@echo off

@echo Getting .Net executables locations
call ..\config\exeLocations.bat

@echo Compiling IconExtractor.module
%net32%\csc.exe @..\rsp\_common.rsp @..\rsp\IconExtractor.rsp ..\IconExtractor.cs
