
### Example of a .NET object accessible to COM

[Overview](#overview)  
[Generate a Strong Name Key Pair](#generate-a-strong-name-key-pair)  
[Compile](#compile)  
[Register](#register)  
[Test](#test)  
[Adding References](#adding-references)  
[Links and References](#links-and-references)  

#### Overview

This demonstrates one way to use the .NET framework compiler in `C:\Windows\Microsoft.NET\Framework\...` to compile a C# file into a library (`.dll`) and then register that library so that it is available to COM. VBScript for example.

Having Visual Studio installed is not required, although if it is installed, `01-generate-keys.vbs` supports generating a stong name key pair for use by the compiler.

`02-compile.vbs` supports compiling the `.cs` file.

`03-register.vbs` supports registering the `.dll` file.

#### Generate a Strong Name Key Pair

Drag `WSHEventLogger.cs` onto `01-generate-keys.vbs`. A `.snk` file will be generated. 

If you don't have Visual Studio installed, you will notice an error message to the effect that this step is not absolutely necessary in order to get the `.cs` file to compile. Note also that in the absence of the `.snk` key file, the corresponding line must be removed from the `.cs` file or else commented out.

#### Compile

Drag `WSHEventLogger.cs` onto `02-compile.vbs`. You can [ignore compiler warning CS1699](https://msdn.microsoft.com/en-us/library/xc31ft41%28v=vs.110%29.aspx?f=255&MSPPError=-2147217396).    

#### Register

Place `WSHEventLogger.dll` in a stable location, then drag it onto `03-register.vbs` or onto a shortcut to the `.vbs` script. The User Account Control dialog will open to verify elevation of privileges.

#### Test

Double click `WSHEventLogger.vbs` to log a test event. Open the event viewer by typing `EventVwr` at a command prompt. When the event viewer opens, expand Windows Logs and select Application. There should be a recent entry with the data described in `WSHEventLogger.vbs`

#### Adding References

In order to add an assembly reference, hardcode the path in `02-compile.vbs`. Open the `.vbs` file to uncomment the reference statement that is required to compile `Vox.cs`.

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
