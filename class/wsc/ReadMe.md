# The `wsc` folder

- [Overview](#overview)  
- [Registration](#registration)  
- [Unregistration](#unregistration)  
- [Instantiation without registration](#instantiation-without-registration)  

## Overview

The `wsc` folder contains Windows Script Component files that make use of like-named class files from the `class` folder and [wrappers](src/ReadMe.md). After the `.wsc` files are [registered](#registration), the associated objects are instantiated by syntax similar to

```vbs
Set formatter = CreateObject( "VBScripting.StringFormatter" )
```

which is more concise than [using VBScripting\.Includer](../../docs/VBScriptClasses.md#includer).

## Registration

Usually, `.wsc` files are registered prior to [instantiation](#overview). Ways to register a `.wsc` file:

- [Setup.vbs](../../Setup.vbs) automatically registers project `.wsc` files.

- Use [RegisterWsc.wsf](../../examples/RegisterWsc.wsf) from the command line or as a drop target.

- In a command console, use syntax similar to

    ``` cmd
    %SystemRoot%\System32\regsvr32.exe /s /i:"<absolute-path-to>\WscFile.wsc" scrobj.dll
    %SystemRoot%\SysWow64\regsvr32.exe /s /i:"<absolute-path-to>\WscFile.wsc" scrobj.dll
    ```

    The first line registers the `.wsc` for 64-bit processes, and the second line registers the `.wsc` for 32-bit processes.

## Unregistration

Unregister a `.wsc` file with syntax similar to

``` cmd
%SystemRoot%\System32\regsvr32.exe /u /n /s /i:"<absolute-path-to>\WscFile.wsc" scrobj.dll
%SystemRoot%\SysWow64\regsvr32.exe /u /n /s /i:"<absolute-path-to>\WscFile.wsc" scrobj.dll
```

Project `.wsc` files are automatically unregistered by running [Uninstall.vbs](../../Uninstall.vbs).

## Instantiation without registration

An unregistered `.wsc` file can be instantiated with syntax similar to

```vbs
Set obj = GetObject("script:" & AbsolutePathToWscFile)
```
