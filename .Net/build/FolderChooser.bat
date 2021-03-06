:: compile FolderChooser.cs as a 32-bit library
:: and register FolderChooser.dll as 64-bit and 32-bit

@echo off
call ..\config\exeLocations.bat

if %1.==/unregister. (
    set verb=Unregistering
) else (
    set verb=Registering
    echo. & echo Compiling FolderChooser.cs
    %net32%\csc.exe @..\rsp\_common.rsp @..\rsp\FolderChooser.rsp ..\FolderChooser.cs
)
echo %verb% FolderChooser.dll for 32-bit apps
%net32%\regasm.exe /codebase %1 ..\lib\FolderChooser.dll

if exist %net64% (
    echo. & echo %verb% FolderChooser.dll for 64-bit apps
    %net64%\regasm.exe /codebase %1 ..\lib\FolderChooser.dll
)
