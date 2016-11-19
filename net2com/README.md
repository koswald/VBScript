
### Example of a .NET object accessible to COM

[Overview](#overview)  
[Generate a Strong Name Key Pair](#generate-a-strong-name-key-pair)  
[Adding References](#adding-references)  
[Compile](#compile)  
[Register](#register)  
[Test](#test)  
[Links and References](#links-and-references)  

#### Overview

This document describes how to use the .NET framework compiler in `C:\Windows\Microsoft.NET\Framework\...`, and the scripts contained in this folder, to compile a C# file into a library (`.dll`) and then register that library so that it is available to COM. VBScript for example.

Having Visual Studio installed is not required, although if it is installed, `01-generate-keys.vbs` supports generating a stong name key pair for use by the compiler.

`02-compile.vbs` supports compiling the `.cs` file.

`03-register.vbs` supports registering the `.dll` file.  

Two example `.cs` files are provided, `WSHEventLogger.cs` and `Vox.cs`.
One of them, `Vox.cs`, requires a reference. See [Adding References].

Compiling multiple files is not addressed.

#### Generate a Strong Name Key Pair

Drag a .cs file (e.g. `WSHEventLogger.cs`) onto `01-generate-keys.vbs`. A `.snk` file will be generated. 

If you don't have Visual Studio installed, you will notice an error message to the effect that this step is not absolutely necessary in order to get the `.cs` file to compile.  

#### Adding References

In order to add an assembly reference, hardcode the path in `02-compile.vbs`. You can open the `.vbs` file to uncomment the reference statement that is required to compile `Vox.cs`.

#### Compile

Drag the same .cs file (e.g. `WSHEventLogger.cs`) onto `02-compile.vbs`. You can [ignore compiler warning CS1699](https://msdn.microsoft.com/en-us/library/xc31ft41%28v=vs.110%29.aspx?f=255&MSPPError=-2147217396).    
**If you don't have Visual Studio**, but you have some version of the .NET framework in `C:\Windows\Microsoft.NET\Framework\...`, then you can still compile and register without a strong name: Remove or comment out the line in the .cs file with AssemblyKeyFileAttribute. You will receive a warning message when registering the .dll.

#### Register

Place the generated .dll (e.g. `WSHEventLogger.dll`) in a stable location, then drag it onto `03-register.vbs` or onto a shortcut to the `.vbs` script. The User Account Control dialog will open to verify elevation of privileges.

#### Test

Double click `WSHEventLogger.vbs` to log a test event. You can open the event viewer by typing `EventVwr` at a command prompt. When the event viewer opens, expand Windows Logs and select Application. There should be a recent entry with the data described in `WSHEventLogger.vbs`

#### Links and References

[Exposing .NET components to COM](http://www.codeproject.com/Articles/3511/Exposing-NET-Components-to-COM)  
[Compiling a .dll from the command line](https://msdn.microsoft.com/en-us/library/78f4aasd.aspx)  
[Regestering a .dll with RegAsm](http://stackoverflow.com/questions/13931337/register-comdlg32-dll-gets-regsvr32-dllregisterserver-entry-point-was-not-found)  
[AssemblyKeyFileAttribute (1)](https://msdn.microsoft.com/en-us/library/system.reflection.assemblykeyfileattribute(v=vs.110).aspx)  
[AssemblyKeyFileAttribute (2)](https://msdn.microsoft.com/en-us/library/xc31ft41%28v=vs.110%29.aspx?f=255&MSPPError=-2147217396)  
[Event logging](https://msdn.microsoft.com/en-us/library/w3t54f67\(v=vs.90\).aspx)  


<br />
<br />
<br />
<br />
<br />
<br />
<br />
<br />
<br />
<br />
<br />
<br />
<br />
<br />
