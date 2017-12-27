
:: compile FolderChooser2.cs as a 32-bit library
:: and register FolderChooser2.dll as 64-bit and 32-bit

@echo off

@echo Getting .Net executables locations
call ..\config\exeLocations.bat

@echo Compiling FolderChooser2.cs
%net32%\csc.exe @..\rsp\_common.rsp @..\rsp\FolderChooser2.rsp ..\FolderChooser2.cs

@echo Registering .dll for 32-bit apps
%net32%\regasm.exe /codebase %1 ..\lib\FolderChooser2.dll

if exist %net64% (
    @echo. & @echo Registering .dll for 64-bit apps
    %net64%\regasm.exe /codebase %1 ..\lib\FolderChooser2.dll
)
