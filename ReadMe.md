# VBScript Utilities

This project has a number of facets:

- [VBScript utility classes] and [documentation](docs/VBScriptClasses.md).
- [C# classes] for extending VBScript and [documentation](docs/CSharpClasses.md).
- A [doc generator for the C# classes].
- A [doc generator for the VBScript classes].
- A lightweight [testing framework].
- A [logger].
- [Examples].  

## Requirements

Windows 10, 8, 7, Vista, Server 20xx, ... 98.

## Installation

Clone the repo and double-click [Setup.vbs]. 
This will register the required Windows Script Component 
file (.wsc) used to manage dependencies, and build the 
[VBScript extensions]. You will be prompted to elevate privileges.

---

Read or edit the [wiki](../../wiki)

[VBScript utility classes]: class
[C# classes]: .Net
[doc generator for the C# classes]: examples/Generate-the-CSharp-docs.vbs 
[doc generator for the VBScript classes]: examples/Generate-the-VBScript-docs.vbs
[testing framework]: class/TestingFramework.vbs
[logger]: class/VBSLogger.vbs
[Examples]: examples
[C# examples]: .Net
[Setup.vbs]: Setup.vbs
[VBScript extensions]: .Net
