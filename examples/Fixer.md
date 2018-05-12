# Fixer.hta

[Overview](#overview)  
[Example](#example)  
[Anomaly](#anomaly)

## Overview

Are you having problems running .vbs and .wsf scripts 
on your 64-bit machine?

Fixer.hta might help with troubleshooting certain Windows&reg; 
configuration problems that prevent `.wsf` or `.vbs` 
scripts from running.  

> Fixer.hta may help to identify the source of a problem 
> in the following situations:  
> 1.) A script won't run at all, or  
> 2.) A script can't instantiate a particular COM object.  

## Example

If a `.vbs` script attempts to instantiate a COM object that 
has been incorrectly registered or compiled for 32-bit use 
only, and the registry is configured to open `.vbs` files 
with the 64-bit `wscript.exe`, then the script may show an error 
similar to `ActiveX component can't create object...`.  

This type of problem might be identified and possibly fixed by 
using Fixer to change the registry entries that control the 
bitness of the executables that open `.vbs` files:  
 
1) Double click `Fixer.bat`in order to launch `Fixer.hta`.  
2) The User Account Control dialog will open to verify elevation 
   privileges for Microsoft (R) HTML Application host.  
3) In the `Fixer` window, in the `..\VBSFile\..` section,  
   select `32-bit`, and then rerun the script.  

## Anomaly

Support for .hta files was removed after an apparent conflict
with anti-malware software, probably Avast or Windows Defender.
Fixer.hta would freeze, preventing it from changing the registry entry
for HKLM\Software\Classes\htafile\Open\Command (default value). 
Fixer.hta worked fine after this capability was removed. 
The registry entry was then successfully changed manually with regedit.exe 
to use the 64-bit mshta.exe, before which some tests didn't run 
as expected.

---

For intructions on one way to compile and register, 
see the `build` folder [ReadMe].

[ReadMe]: ../.Net/build/ReadMe.md