# VBScripting utilities

- [Features](#features)  
- [Requirements](#requirements)  
- [Install](#install)  
- [Uninstall](#uninstall)  
- [Issues](#issues)  
- [References](#references)  

## Features

- [VBScript utility classes] and [documentation](./docs/VBScriptClasses.md).  
- [C# classes] for extending VBScript and [documentation](./docs/CSharpClasses.md).  
- [Integration tests](./spec/ReadMe.md) use an ultralight TestingFramework class ( [code](./class/TestingFramework.vbs) | [doc](https://github.com/koswald/VBScript/blob/master/docs/VBScriptClasses.md#user-content-testingframework) ) written in VBScript.
- A VBScript statement [interpreter]/console.  
- A dependency manager ( [code](./class/Includer.vbs) | [doc](https://github.com/koswald/VBScript/blob/master/docs/VBScriptClasses.md#user-content-includer) ).  
- A Configurer class ( [code](./class/Configurer.vbs) | [doc](https://github.com/koswald/VBScript/blob/master/docs/VBScriptClasses.md#user-content-configurer) )
- A logger class ( [code](./class/VBSLogger.vbs) | [doc](https://github.com/koswald/VBScript/blob/master/docs/VBScriptClasses.md#user-content-vbslogger) ).
- A [registry classes] manager UI.  
- An [icon extractor] UI.  
- A [startup items] editor UI.  
- A [speech synthesis] UI.  
- A [system tray icon] proof of concept.
- A [progress bar] proof of concept.
- A script for keeping the computer awake while
  giving a [presentation], with a system tray icon.  
- A doc generator for the C# classes ( [example code](./examples/Generate-the-CSharp-docs.vbs) | [class code](./class/DocGeneratorCS.vbs) | [doc](https://github.com/koswald/VBScript/blob/master/docs/VBScriptClasses.md#user-content-docgeneratorcs) ) and a doc generator for the VBScript classes ( [example code](./examples/Generate-the-VBScript-docs.vbs) | [class code](./class/DocGenerator.vbs) | [doc](https://github.com/koswald/VBScript/blob/master/docs/VBScriptClasses.md#user-content-docgenerator) ), both based on code comments.  
- More [examples] of .vbs and .hta scripts.
- [Windows Script Component files].

## Requirements

Windows 11, 10, 8.1, 8, 7, Vista, ... 98.

## Install

- Clone or download the repo. [CopyToProgramFiles.vbs](./CopyToProgramFiles.vbs) can be used, if desired, to make the project available to all users before running [Setup.vbs].

- Double-click [Setup.vbs] or type the following command in a console window. If the console does not have elevated privileges, then the User Account Control dialog will open, in order to request permission to elevate.  

``` cmd
Setup.vbs
```

or for a non-interactive install, type the following command in an elevated console window:  

``` cmd
Setup.vbs /s
```

This will register the [Windows Script Component files], compile and register the [VBScript extensions], and create the VBScripting event log source.  

## Uninstall

Uninstalling unregisters the project .dll files and .wsc files and removes the VBScripting event log source, without removing files.  

From a console window, type

``` cmd
Uninstall.vbs
```

or

``` cmd
Setup.vbs /u
```

Or type

``` cmd
start ms-settings:appsfeatures
```

or

``` cmd
control /name Microsoft.ProgramsAndFeatures
```

and then select VBScripting Utility Classes and Extensions and click Uninstall.  

Or for a silent uninstall  type the following command from an elevated console window:

``` cmd
Uninstall.vbs /s
```

or

``` cmd
Setup.vbs /u /s
```

> Note: Uninstalling does not remove files.

## Issues

After a major Windows 10 version update, rerunning [Setup.vbs] may be required in order to reregister the project classes. A restart may be required before rerunning [Setup.vbs] after updating to Windows 10 version 20H2.  

## References  

Links to the official online documentation:

- [VBScript Fundamentals](https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/0ad0dkea(v=vs.84))  
- [VBScript Language Reference](https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/d1wf56tt(v=vs.84))  
- [FileSystemObject](https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/6kxy1a51(v=vs.84))  
- [WshShell object](https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/aew9yb99(v=vs.84))  
- [WshScriptExec object](https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/2f38xsxe(v=vs.84))  
- [Dictionary object](https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/x4k5wbx4(v=vs.84))  
- [Regular expressions](https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/6wzad2b2(v=vs.84))  
- [WScript object](https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/at5ydy31(v=vs.84))  
- [WshArguments object](https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/ss1ysb2a(v=vs.84))  
- [WshNamed object](https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/d6y04sbb(v=vs.84))  
- [WshEnvironment object](https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/6s7w15a0(v=vs.84))  
- [WshNetwork object](https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/s6wt333f(v=vs.84))  
- [Script Components](https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/asxw6z3c(v=vs.84))  
- [Windows Script Host](https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/9bbdkx3k(v=vs.84))  
- [StdRegProv object](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/regprov/stdregprov)  
- [WMI Tasks for Scripts and Applications](https://learn.microsoft.com/en-us/windows/win32/wmisdk/wmi-tasks-for-scripts-and-applications)  
- [WQL (SQL for WMI)](https://learn.microsoft.com/en-us/windows/win32/wmisdk/wql-sql-for-wmi)  
- [Shell object (Shell.Appliction)](https://learn.microsoft.com/en-us/windows/win32/shell/shell)  
  - [ShellExecute](https://learn.microsoft.com/en-us/windows/win32/shell/shell-shellexecute),,, "[runas](https://learn.microsoft.com/en-us/windows/win32/shell/launch#object-verbs)"  
  - [MoveHere](https://learn.microsoft.com/en-us/windows/win32/shell/folder-movehere)  
  - [CopyHere](https://learn.microsoft.com/en-us/windows/win32/shell/folder-copyhere)  
  - [MinimizeAll](https://learn.microsoft.com/en-us/windows/win32/shell/shell-minimizeall)  
  - [UndoMinimizeAll](https://learn.microsoft.com/en-us/windows/win32/shell/shell-undominimizeall)  
  - [ShellSpecialFolderConstants enumeration (shldisp.h)](https://learn.microsoft.com/en-us/windows/win32/api/shldisp/ne-shldisp-shellspecialfolderconstants)  



[VBScript utility classes]: ./class
[C# classes]: ./.Net
[examples]: ./examples
[Setup.vbs]: ./Setup.vbs
[Windows Script Component files]: https://github.com/koswald/VBScript/blob/master/class/wsc/ReadMe.md#user-content-overview
[VBScript extensions]: ./.Net
[registry classes]: ./examples/RegistryClasses.hta
[icon extractor]: ./examples/IconExtractor.hta
[startup items]: ./examples/StartItems.hta
[speech synthesis]: ./examples/SpeechSynthesis.hta
[speech synthesis]: ./examples/SpeechSynthesis.hta
[presentation]: ./examples/Presentation.vbs
[interpreter]: ./examples/VBSInterpreter.hta
[system tray icon]: ./.Net/test/NotifyIcon-test.vbs
[progress bar]: ./.Net/test/ProgressBar-test.vbs
