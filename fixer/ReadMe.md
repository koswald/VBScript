
# Fixing problems running .vbs, .wsf, and .hta scripts on 64-bit machines

Fixer.hta is intended to easily troubleshoot and fix certain Windows&reg; configuration problems that may prevent `.wsf`, `.vbs`, or `.hta` scripts from running.  

Specifically, scripts may be prevented from running in 64 bit systems when the registry is configured to use the wrong executable.  

For example, if a `.vbs` script attempts to instantiate a COM object, but the registry is configured to open that file type with the 64-bit `wscript.exe`, then the script may show an error similar to `ActiveX component can't create object...`.  

This type of problem may be easily fixed as follows:  
 
1) Double click `Fixer.bat`, which launches `Fixer.hta`. 

2) In the `Fixer` window, for the `.vbs` file type, select 32-bit.