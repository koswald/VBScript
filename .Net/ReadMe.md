###### The `.Net` folder

# ReadMe.md

The `.Net` folder contains C# examples for creating 
`.dll` libraries that can be called from VBScript. 

#### Compiling and registering

Two options for creating the class libraries are
 
1. Use Visual Studio, or 
2. Use the .Net executables in a way that is more command-line oriented. 
For detailed help with this approach refer 
to the [ReadMe] in the `build` folder.

#### Features

1) For the simplest illustration of making a C# method available to 
   VBScript, see [EventLogger.cs].
2) For an example of a class requiring an assembly reference, and 
   an illustration of how to do it, see [SpeechSynthesis.cs],
   [SpeechSynthesis.rsp].
3) For an example of a COM event, or callback, see [NotifyIcon.cs]
   and [NotifyIcon-test.vbs].
4) For an example of a progress bar (for illustration only), 
   see [ProgressBar.cs] and [ProgressBar-test.vbs].

[ReadMe]: build/ReadMe.md
[EventLogger.cs]: EventLogger.cs
[SpeechSynthesis.cs]: SpeechSynthesis.cs
[SpeechSynthesis.rsp]: rsp/SpeechSynthesis.rsp
[NotifyIcon.cs]: NotifyIcon.cs
[NotifyIcon-test.vbs]: test/NotifyIcon-test.vbs
[ProgressBar.cs]: ProgressBar.cs
[ProgressBar-test.vbs]: test/ProgressBar-test.vbs

#### Links and References

[Interoperability (C# Programming Guide)](https://docs.microsoft.com/en-us/dotnet/csharp/programming-guide/interop/ "docs.microsoft.com")  
[Example COM Class (MSDN)](https://docs.microsoft.com/en-us/dotnet/csharp/programming-guide/interop/example-com-class "docs.microsoft.com")  
[Exposing .NET components to COM](http://www.codeproject.com/Articles/3511/Exposing-NET-Components-to-COM "www.codeproject.com")  
[Calling back to VBScript from C#](https://stackoverflow.com/questions/1044872/calling-back-to-vbscript-from-c-sharp "stackoverflow.com")  
[Extracting an icon from a .dll file](https://stackoverflow.com/questions/6872957/how-can-i-use-the-images-within-shell32-dll-in-my-c-sharp-project "stackoverflow.com")  
[Invoking the NotifyIcon context menu](https://stackoverflow.com/questions/2208690/invoke-notifyicons-context-menu "stackoverflow.com")  
[Browse for a directory in C#](https://stackoverflow.com/questions/11767/browse-for-a-directory-in-c-sharp#33817043 "stackoverflow.com")  

##### Compiler-supported code comments

The compiler-supported code comments are used by Visual Studio for 
intellisense.  

[Documenting your code with XML comments](https://docs.microsoft.com/en-us/dotnet/csharp/codedoc "docs.microsoft.com")  
[XML Documentation Comments (C# Programming Guide)](https://github.com/dotnet/docs/blob/master/docs/csharp/programming-guide/xmldoc/xml-documentation-comments.md "github.com/dotnet/docs")  
