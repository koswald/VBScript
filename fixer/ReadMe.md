
# Having problems running .vbs, .wsf, and .hta scripts on your 64-bit machine?

Fixer.hta might help with troubleshooting certain Windows&reg; configuration problems that may prevent `.wsf`, `.vbs`, or `.hta` scripts from running.  

> Fixer.hta may help to identify the source of a problem in the following situations:  
> 1.) A script won't run at all, or  
> 2.) A script can't instantiate a particular COM object.  

### Example .vbs script problem

If a `.vbs` script attempts to instantiate a COM object that has been incorrectly registered or compiled for 32-bit use only, and the registry is configured to open `.vbs` files with the 64-bit `wscript.exe`, then the script may show an error similar to `ActiveX component can't create object...`.  

This type of problem might be identified by using Fixer to change the registry entries that control the bitness of the executables that open `.vbs` files:  
 
1) Double click `Fixer.bat`, which which call `Fixer.vbs` and help to ensure that `Fixer.hta` is opened with the 64-bit executable. `Fixer.vbs` will call `Fixer.hta` with elevated privileges.  

2) In the `Fixer` window, in the `..\VBSFile\..` section, select `32-bit`, and then rerun the script.  

---

For an explanation of how to compile and register correctly for both bitnesses, see the documentation for the DotNetCompiler class. If you haven't cloned the git project, or haven't run the [doc generator](https://github.com/koswald/VBScript/blob/master/examples/documentation%20generator/Generate-the-docs.vbs), then you don't have access to the documentation, but you can still read the pertinent code comments [here](https://github.com/koswald/VBScript/blob/master/class/DotNetCompiler.vbs) (buried in a very long line #7 or so).
