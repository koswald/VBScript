# Description of the `net2com` folder

This folder contains examples for creating and testing a `.dll` file that can be instantiated from a script. 
It uses the .Net compiler and .Net registrar, csc.exe and regasm.exe respectively, which are present on most machines.
The default location for these `.exe` files is hardcoded in the class file `DotNetCompiler.vbs` as 
`C:\Windows\Microsoft.NET\Framework64\v4.0.30319` for the 64-bit versions and
`C:\Windows\Microsoft.NET\Framework\v4.0.30319` for the 32-bit versions.

If Visual Studio installed, then a strong-name key pair is generated without requiring Visual Studio to be opened.

#### Links and References

[Exposing .NET components to COM](http://www.codeproject.com/Articles/3511/Exposing-NET-Components-to-COM)  
[Compiling a .dll from the command line](https://msdn.microsoft.com/en-us/library/78f4aasd.aspx)  
[Registering a .dll with RegAsm](http://stackoverflow.com/questions/13931337/register-comdlg32-dll-gets-regsvr32-dllregisterserver-entry-point-was-not-found)  
[AssemblyKeyFileAttribute (1)](https://msdn.microsoft.com/en-us/library/system.reflection.assemblykeyfileattribute(v=vs.110).aspx)  
[AssemblyKeyFileAttribute (2)](https://msdn.microsoft.com/en-us/library/xc31ft41%28v=vs.110%29.aspx?f=255&MSPPError=-2147217396)  
[Event logging](https://msdn.microsoft.com/en-us/library/w3t54f67\(v=vs.90\).aspx)  