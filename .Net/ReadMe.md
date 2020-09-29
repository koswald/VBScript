# The `.NET` folder

[Overview]  
[Compiling and registering]  
[Features]  
[Compatibility]  
[Links and references]

## Overview

The `.NET` folder contains C# examples for creating `.dll` libraries that can be called from VBScript, extending the power of VBScript using .NET Framework.

## Compiling and registering

Among the options for building (compiling and registering) the class libraries are

1. Run [Setup.vbs] in the project folder.  
2. Run `build.vbs` in the [build] folder.  
3. Run one or more of the `.bat` scripts in the [build] folder from an elevated command prompt.  
4. Drag one or more of the `.bat` scripts in the [build] folder onto `build.vbs`.
5. Use Visual Studio (not tested).  

## Features

Features include the following:  

1) [Documentation].
2) [Manual tests] that demonstrate functionality and how-to-use.
3) [Automated integration tests].
4) For an example of a COM event, or callback, see [NotifyIcon.cs]
   and [NotifyIcon-test.vbs].
5) For an example of a progress bar (for illustration only). See [ProgressBar.cs] and [ProgressBar-test.vbs].
6) For a simple example of making a C# method available to VBScript, see [EventLogger.cs].
7) For an example of a class requiring an assembly reference, and  an illustration of how to do it, see [SpeechSynthesis.cs], and [SpeechSynthesis.rsp].
8) For a user-friendly file chooser dialog, see [FileChooser.cs].
9) For a user-friendly folder chooser dialog, see [FolderChooser.cs]  and [FolderChooser2.cs]. These two files are adapted from  stackoverflow.com posts. The exposed features of the two are identical. `FolderChooser2` is a backup in case `FolderChooser.cs` breaks due to future changes in the private members that are invoked  using [Reflection].

## Compatibility

Most thoroughly tested on Windows 10, the libraries are all expected to work on Windows versions as old as Vista, with the exception of `SpeechSynthesis`, which requires a reference probably not available on most Vista machines.

## Links and References

[Interoperability (C# Programming Guide)](https://docs.microsoft.com/en-us/dotnet/csharp/programming-guide/interop/ "docs.microsoft.com")  
[Example COM Class](https://docs.microsoft.com/en-us/dotnet/csharp/programming-guide/interop/example-com-class "docs.microsoft.com")  
[Exposing .NET components to COM](http://www.codeproject.com/Articles/3511/Exposing-NET-Components-to-COM "www.codeproject.com")  
[Calling back to VBScript from C#](https://stackoverflow.com/questions/1044872/calling-back-to-vbscript-from-c-sharp#45927249 "stackoverflow.com")  
[Extracting an icon from a .dll file](https://stackoverflow.com/questions/6872957/how-can-i-use-the-images-within-shell32-dll-in-my-c-sharp-project#6873026 "stackoverflow.com")  
[Invoking the NotifyIcon context menu](https://stackoverflow.com/questions/2208690/invoke-notifyicons-context-menu#2208910 "stackoverflow.com")  
[Browse for a directory in C#](https://stackoverflow.com/questions/11767/browse-for-a-directory-in-c-sharp#33817043 "stackoverflow.com")  
[Show detailed browser from a property grid](https://stackoverflow.com/questions/15368771/show-detailed-folder-browser-from-a-propertygrid#15386992 "stackoverflow.com")  
[Component Object Model (COM)](https://docs.microsoft.com/en-us/windows/win32/com/component-object-model--com--portal?redirectedfrom=MSDN "docs.microsoft.com")

### Compiler-supported code comments

[Documenting your code with XML comments](https://docs.microsoft.com/en-us/dotnet/csharp/codedoc "docs.microsoft.com")  
[XML Documentation Comments (C# Programming Guide)](https://github.com/dotnet/docs/blob/master/docs/csharp/programming-guide/xmldoc/xml-documentation-comments.md "github.com/dotnet/docs")  

`` `` `` ``


[Overview]: #overview
[Compiling and registering]: #compiling-and-registering
[Features]: #features
[Compatibility]: #compatibility
[Links and references]: #links-and-references

[Documentation]: ../docs/CSharpClasses.md
[build]: build
[EventLogger.cs]: EventLogger.cs
[SpeechSynthesis.cs]: SpeechSynthesis.cs
[SpeechSynthesis.rsp]: rsp/SpeechSynthesis.rsp
[NotifyIcon.cs]: NotifyIcon.cs
[NotifyIcon-test.vbs]: test/NotifyIcon-test.vbs
[ProgressBar.cs]: ProgressBar.cs
[ProgressBar-test.vbs]: test/ProgressBar-test.vbs
[FileChooser.cs]: FileChooser.cs
[FolderChooser.cs]: FolderChooser.cs
[FolderChooser2.cs]: FolderChooser2.cs
[Reflection]: https://docs.microsoft.com/en-us/dotnet/api/system.reflection?view=netframework-4.7.1 "docs.microsoft.com"
[Setup.vbs]: ../Setup.vbs
[Manual tests]: test
[Automated integration tests]: ../tests/dll
