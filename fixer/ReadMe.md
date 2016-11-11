
# Having problems running .vbs, .wsf, and .hta scripts on your 64-bit machine?

Fixer.hta is intended to easily troubleshoot and fix certain Windows&reg; configuration problems that may prevent `.wsf`, `.vbs`, or `.hta` scripts from running.  

Specifically, scripts may be prevented from running in 64 bit systems when the registry is configured to use the wrong executable.  

For example, if a `.vbs` script attempts to instantiate a COM object, but the registry is configured to open that file type with the 64-bit `wscript.exe`, then the script may show an error similar to `ActiveX component can't create object...`.  

This type of problem might be fixed by changing the registry entries to use the 32-bit executable to open .vbs files. Keeping in mind that other .vbs scripts that rely on 64-bit executables can break when you do this,  
 
1) Double click `Fixer.bat`, which launches `Fixer.hta` with elevated privileges.  

2) In the `Fixer` window, for the `.vbs` file type, select 32-bit.