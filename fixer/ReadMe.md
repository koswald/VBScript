
# Having problems running .vbs, .wsf, and .hta scripts on your 64-bit machine?

Fixer.hta is most appropriately used to **troubleshoot** certain Windows&reg; configuration problems that may prevent `.wsf`, `.vbs`, or `.hta` scripts from running.  

> Specifically, scripts may be 
> 1.) prevented from running at all when Windows is configured to open the file type with the executable of a particular bitness, or  
> 2.) prevented from instantiating a particular COM object when the .dll was compiled or registered incorrectly.  

For example, if a `.vbs` script attempts to instantiate a COM object that has been incorrectly registered or compiled for 32-bit use only, and the registry is configured to open that file type with the 64-bit `wscript.exe`, then the script may show an error similar to `ActiveX component can't create object...`.  

This type of problem might be identified by using Fixer to change the registry entries that control the bitness of the executables that open .vbs, .hta, and/or .wsf files.  
 
1) Double click `Fixer.bat`, which launches `Fixer.hta` with elevated privileges.  

2) In the `Fixer` window, for the `.vbs` file type, select the other bitness.  

For an explanation of how to compile and register correctly for both bitnesses, see the documentation for the DotNetCompiler class. Or if you haven't cloned the git project, or haven't run the [doc generator](https://github.com/koswald/VBScript/blob/master/examples/Generate-the-docs.vbs), you can read the code comments [here](https://github.com/koswald/VBScript/blob/master/class/DotNetCompiler.vbs) (buried in a very long line #7 or so).
