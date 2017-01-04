
:: prepare to manually launch a series of tests from a console window

@echo off

:: define and display environment variables

echo Environment variables & echo.
set test0=%SystemRoot%\SysWoW64\cscript.exe //nologo TestLauncher.vbs
set test0=%SystemRoot%\System32\cscript.exe //nologo TestLauncher.vbs
set test1=%test0% VBSTimer.spec.vbs
echo %%test0%%=%test0%
echo %%test1%%=%test1%

:: define and display doskey macros

echo. & echo DosKey macros & echo.
set macro0=test0=%test0%
set macro1=test1=%test1%
doskey %macro0%
doskey %macro1%
echo %macro0%
echo %macro1%
echo.

:: run all of the tests

%test0%

:: run a single test

:: %test1%

:: run it again, twice

:: %test1% 2

:: leave the command window open

cmd /k
