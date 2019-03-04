###### The `.Net` folder

# ReadMe.md

The `.Net` folder contains C# examples for creating 
`.dll` libraries that can be called from VBScript. 

## Compiling and registering

Among the options for building (compiling and registering) the class libraries are
 
1. Use Visual Studio.  
2. Run [Setup.vbs] in the project folder.  
3. Run `build.vbs` in the [build] folder.  
4. Run one or more of the `.bat` scripts in the [build] folder from an elevated 
   command prompt.  
5. Drag one or more of the `.bat` scripts in the [build] folder 
   onto `build.vbs`.

## Features

Features include the following:  

1) [Documentation].
2) [Manual tests] that demonstrate functionality.
3) For an example of a COM event, or callback, see [NotifyIcon.cs]
   and [NotifyIcon-test.vbs].
4) For an example of a progress bar (for illustration only), 
   see [ProgressBar.cs] and [ProgressBar-test.vbs]. 
5) For a simple example of making a C# method available to 
   VBScript, see [EventLogger.cs].
6) For an example of a class requiring an assembly reference, and 
   an illustration of how to do it, see [SpeechSynthesis.cs],
   [SpeechSynthesis.rsp].
7) For a user-friendly file chooser dialog, see [FileChooser.cs].
8) For a user-friendly folder chooser dialog, see [FolderChooser.cs] 
   and [FolderChooser2.cs]. These two files are adapted from 
   stackoverflow.com posts. The exposed features of the two are 
   identical. `FolderChooser2` is a backup in case `FolderChooser.cs`
   breaks due to future changes in the private members that are invoked 
   using [Reflection].

## Compatibility

Most thoroughly tested on Windows 10, the libraries 
are all expected to work on Windows versions as old as Vista, with 
the exception of `SpeechSynthesis`, which requires a reference 
probably not available on most Vista machines.

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

## Links and References

[Interoperability (C# Programming Guide)](https://docs.microsoft.com/en-us/dotnet/csharp/programming-guide/interop/ "docs.microsoft.com")  
[Example COM Class](https://docs.microsoft.com/en-us/dotnet/csharp/programming-guide/interop/example-com-class "docs.microsoft.com")  
[Exposing .NET components to COM](http://www.codeproject.com/Articles/3511/Exposing-NET-Components-to-COM "www.codeproject.com")  
[Calling back to VBScript from C#](https://stackoverflow.com/questions/1044872/calling-back-to-vbscript-from-c-sharp#45927249 "stackoverflow.com")  
[Extracting an icon from a .dll file](https://stackoverflow.com/questions/6872957/how-can-i-use-the-images-within-shell32-dll-in-my-c-sharp-project#6873026 "stackoverflow.com")  
[Invoking the NotifyIcon context menu](https://stackoverflow.com/questions/2208690/invoke-notifyicons-context-menu#2208910 "stackoverflow.com")  
[Browse for a directory in C#](https://stackoverflow.com/questions/11767/browse-for-a-directory-in-c-sharp#33817043 "stackoverflow.com")  
[Show detailed browser from a property grid](https://stackoverflow.com/questions/15368771/show-detailed-folder-browser-from-a-propertygrid#15386992 "stackoverflow.com")  

##### Compiler-supported code comments

[Documenting your code with XML comments](https://docs.microsoft.com/en-us/dotnet/csharp/codedoc "docs.microsoft.com")  
[XML Documentation Comments (C# Programming Guide)](https://github.com/dotnet/docs/blob/master/docs/csharp/programming-guide/xmldoc/xml-documentation-comments.md "github.com/dotnet/docs")  
