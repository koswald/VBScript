# VBScripting utilities

- [Features](#features)  
- [Requirements](#requirements)  
- [Installation](#installation)  
- [Uninstall](#uninstall)  
- [Issues](#issues)  
- [References](#references)  

## Features

- [VBScript utility classes] and [documentation](docs/VBScriptClasses.md).  
- [C# classes] for extending VBScript and [documentation](docs/CSharpClasses.md).  
- [Integration tests](spec/ReadMe.md) use an ultralight [testing framework] written in VBScript.
- A VBScript statement [interpreter]/console.  
- A [dependency manager].  
- A [logger].  
- A [registry classes] manager UI.  
- An [icon extractor] UI.  
- A [startup items] editor UI.  
- A [speech synthesis] UI.  
- A [system tray icon] proof of concept.
- A [progress bar] proof of concept.
- A script for keeping the computer awake while
  giving a [presentation], with a system tray icon.  
- A [doc generator for the C# classes] and a [doc generator for the VBScript classes], both based on code comments.  
- More [examples] of .vbs and .hta scripts.
- [Windows Script Component files].

## Requirements

Windows 10, 8, 7, Vista, ... 98.

## Installation

- Clone or download the repo. If desired, put it in a location visible to all users, or run [CopyToProgramFiles.vbs](CopyToProgramFiles.vbs).  

- Double-click [Setup.vbs] or from a console window type

``` cmd
Setup.vbs
```

or for a non-interactive install type

``` cmd
Setup.vbs /s
```

This will register the [Windows Script Component files], and compile and register the [VBScript extensions].  

## Uninstall

To unregister the project .dll files and .wsc files, and remove the VBScripting event log source,

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

and then select VBScripting Utility Classes and Extensions and click Uninstall.  

Or for a silent uninstall from a console window, type

``` cmd
Uninstall.vbs /s
```

or

``` cmd
Setup.vbs /u /s
```

> Note: Files are not removed.

## Issues

After a Windows 10 version update, rerunning [Setup.vbs] is usually required in order to reregister project classes. A restart may be required before rerunning [Setup.vbs] after updating to Windows 10 version 20H2.  

## References  

[Scripting documentation online]  
[Scripting links]

[VBScript utility classes]: class
[C# classes]: .Net
[doc generator for the C# classes]: examples/Generate-the-CSharp-docs.vbs
[doc generator for the VBScript classes]: examples/Generate-the-VBScript-docs.vbs
[testing framework]: docs/VBScriptClasses.md#testingframework
[dependency manager]: docs/VBScriptClasses.md#includer
[logger]: docs/VBScriptClasses.md#vbslogger
[examples]: examples
[Setup.vbs]: Setup.vbs
[Windows Script Component files]: class/wsc/ReadMe.md#the-wsc-folder
[VBScript extensions]: .Net
[registry classes]: examples/RegistryClasses.hta
[icon extractor]: examples/icon-extractor.hta
[startup items]: examples/StartItems.hta
[speech synthesis]: examples/SpeechSynthesis.hta
[speech synthesis]: examples/SpeechSynthesis.hta
[presentation]: examples/Presentation.vbs
[interpreter]: examples/VBSInterpreter.hta
[system tray icon]: .Net/test/NotifyIcon-test.vbs
[progress bar]: .Net/test/ProgressBar-test.vbs
[Scripting documentation online]: https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/d1et7k7c(v%3dvs.84) "docs.microsoft.com"
[Scripting links]: https://technet.microsoft.com/en-us/library/cc498722.aspx "technet.microsoft.com"
