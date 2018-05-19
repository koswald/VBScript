# Fixer.hta

[Overview](#overview)  
[Background](#background)  
[How it works](#how-it-works)  
[Fixing the fixer](#fixing-the-fixer)

## Overview

Fixer.hta can help to solve certain problems 
in 64-bit systems when the problems are 
related to bitness, allowing you to quickly 
toggle script bitness without resorting to 
regedit.exe.  

## Background

For a 64-bit system configured in the typical 
manner, the default values of the following 
registry keys determine which executable 
is used to open .vbs, .wsf, and .hta files:  
```
HKLM\Software\Classes\VBSFile\Shell\Open\Command
HKLM\Software\Classes\WSFFile\Shell\Open\Command
HKLM\Software\Classes\htafile\Shell\Open\Command
```
For a system configured to open .vbs 
scripts with the 64-bit wscript, the 
command might look something like,  

`%SystemRoot%\System32\wscript.exe "%1" %*`

## How it works

**With Fixer.hta you can easily toggle the 
bitness without resorting to regedit.exe.**

When you select the *32 bit* radio button, Fixer.hta 
changes the `System32` to `SysWow64`.  

## Fixing the fixer

One situation where you might have to 
jump start this process is when .hta files 
are configured to open with the 32-bit 
mshta.exe. In this case, the radio buttons 
have no effect: the command stays at 
`...\SysWow64\...`. One way to force Fixer.hta 
to open with the 64-bit mshta.exe, is

- Press Win + X, R (to open the Run dialog)
- Type `mshta "<project folder>\examples\Fixer.hta"`

<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>


  



