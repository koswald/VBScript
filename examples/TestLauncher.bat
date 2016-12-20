
:: prepare to manually launch a series of tests from a console window

@echo off

:: define and display environment variables

set launch0=cscript //nologo TestLauncher.vbs
set launch1=%launch0% VBSClipboard.spec.vbs

echo Environment variables & echo.

echo %%launch0%%=%launch0%
echo %%launch1%%=%launch1%

:: define and display doskey macros

set macro0=all=%launch0%
set macro1=wip1=%launch1%

doskey %macro0%
doskey %macro1%

echo. & echo DosKey macros & echo.

echo %macro0%
echo %macro1%
echo.

:: start the first test

cmd /k %launch1% 3
