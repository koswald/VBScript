###### The `build` folder

# ReadMe.md

### Overview

The `build` folder contains scripts for 
compiling the C# files and registering 
the output files, using the .Net executables `csc.exe` 
and `regasm.exe` present on most Windows&reg; computers.  

Both of these `.exe` files have a 32-bit version and a 64-bit version. 
This project uses the 32-bit `csc.exe` and both `regasm.exe` files.
The location of these `.exe` files is configurable in [DotNetCompiler.config] 
and [exeLocations.bat].

#### Instructions for compiling and registering the class libraries.

1. First generate a strong-name key 
   (recommended) or else opt out.
   See the [ReadMe] in the `key` folder.  

2. Then use the scripts in the this 
   folder to generate the class libraries:  

    Drag and drop one or more of the `.bat` files onto `build.vbs` 
    to compile the associated `C#` file(s) and, if 
    appropriate, to register 
    the compiled file. Or else double-click 
    `build.vbs` to compile and register all.

[ReadMe]: ../key/ReadMe.md
[DotNetCompiler.config]: ../../class/DotNetCompiler.config "../../class/DotNetCompiler.config"
[exeLocations.bat]: ../config/exeLocations.bat "../config/exeLocations.bat"

### Links

[Working with the C# 2.0 Command Line Compiler](https://msdn.microsoft.com/en-us/library/ms379563(v=vs.80).aspx "From msdn.microsoft.com. Dated but still very useful")  
[Compiling a .dll from the command line](https://msdn.microsoft.com/en-us/library/78f4aasd.aspx "msdn.microsoft.com")  
[Registering a .dll with RegAsm](http://stackoverflow.com/questions/13931337/register-comdlg32-dll-gets-regsvr32-dllregisterserver-entry-point-was-not-found "stackoverflow.com")  
[.NET Module vs Assembly](https://stackoverflow.com/questions/9271805/net-module-vs-assembly "stackoverflow.com")  
