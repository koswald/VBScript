:: register.vbs a single .wsc file passed in on the command line
@echo off
If %1.==. (
    echo Arg required. Press a key to exit
	pause > nul
	exit /b
)
For %%A in (%1) do (
    echo Registering %%~nxA
)
..\..\examples\Register.vbs %1