:: set locations of csc.exe and regasm.exe,
:: for compiling and registering

@echo Setting .Net executables locations
set net64=%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319
set net32=%SystemRoot%\Microsoft.NET\Framework\v4.0.30319

:: validate
if not exist %net64%\csc.exe    echo Couldn't find 64-bit csc.exe
if not exist %net64%\regasm.exe echo Couldn't find 64-bit regasm.exe
if not exist %net32%\csc.exe    echo Couldn't find 32-bit csc.exe
if not exist %net32%\regasm.exe echo Couldn't find 32-bit regasm.exe
