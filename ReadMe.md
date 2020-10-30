# VBScripting utilities

This project features

- [VBScript utility classes] and [documentation](docs/VBScriptClasses.md).  
- [C# classes] for extending VBScript and [documentation](docs/CSharpClasses.md).  
- [Integration Tests](spec).  
- A VBScript statement [interpreter]/console.  
- An ultralight [testing framework].  
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

Clone or download the repo and run [Setup.vbs].
This will register the [Windows Script Component files], and compile and register the [VBScript extensions].

## References  

[Scripting documentation online]  
[Scripting links]

`` ``

Read or edit the [wiki](../../wiki)

[VBScript utility classes]: class
[C# classes]: .Net
[doc generator for the C# classes]: examples/Generate-the-CSharp-docs.vbs
[doc generator for the VBScript classes]: examples/Generate-the-VBScript-docs.vbs
[testing framework]: class/TestingFramework.vbs
[logger]: class/VBSLogger.vbs
[examples]: examples
[Setup.vbs]: Setup.vbs
[Windows Script Component files]: class/wsc
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
