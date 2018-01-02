###### The `build` folder

# ReadMe.md

### Overview

The `build` folder contains scripts for 
compiling the C# files and registering 
the output files, using the .Net executables `csc.exe` 
and `regasm.exe` present on most Windows&reg; computers.  

In 64-bit systems, each of these `.exe` files have a 
32-bit version and a 64-bit version. 
This project uses the 32-bit `csc.exe` and both 
`regasm.exe` files, when both are available.
The location of these `.exe` files is configurable in 
[exeLocations.bat].

#### Options for compiling and registering the class libraries.

1. The class libraries are generated automatically 
by running the required `Setup.vbs` in the project 
folder.  

2. To (re)compile and (re)register one or more of the `.cs` files, 
open the `build` folder and drag and drop the `.bat` file(s) 
having the same name as the `.cs` file(s) onto `build.vbs`.  

3. To (re)compile and (re)register 
all, double-click `build.vbs`.  

By default, the compiler doesn't sign the output files, as 
recommended by Microsoft, so (ignorable) warnings will be 
generated when files are compiled and registered.

In order to sign the output files,  
- Download Visual Studio.
- Right click the Developer Command Prompt for VS in the Start Menu.
- Select Run as adminstrator.
- Generate your keypair file by running a command similar to  
  `sn -k %UserProfile%\MyKeyPair.snk`
- Modify [_common.rsp]:
  - Comment out or delete `/delaysign`.
  - Uncomment the `/keyfile` line and specify the location and name of 
    the `.snk` file. It is recommended that you keep your keyfile out 
    of the project folders.
- See the [config folder ReadMe] for git configuration recommendations.

[DotNetCompiler.config]: ../../class/DotNetCompiler.config "../../class/DotNetCompiler.config"
[exeLocations.bat]: ../config/exeLocations.bat "../config/exeLocations.bat"
[_common.rsp]: ../rsp/_common.rsp
[config folder ReadMe]: ../config/ReadMe.md#recommended-git-configuration

### Links

[Working with the C# 2.0 Command Line Compiler](https://msdn.microsoft.com/en-us/library/ms379563(v=vs.80).aspx "From msdn.microsoft.com. Dated but still very useful")  
[Compiling a .dll from the command line](https://msdn.microsoft.com/en-us/library/78f4aasd.aspx "msdn.microsoft.com")  
[Registering a .dll with RegAsm](http://stackoverflow.com/questions/13931337/register-comdlg32-dll-gets-regsvr32-dllregisterserver-entry-point-was-not-found "stackoverflow.com")  
[.NET Module vs Assembly](https://stackoverflow.com/questions/9271805/net-module-vs-assembly "stackoverflow.com")  
