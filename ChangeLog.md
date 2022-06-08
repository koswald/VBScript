- [Version 1.4.1 (the latest version)](#version-141)  
- [Version 1.4.0](#version-140)  

# Version 1.4.1

- [WMIUtility class](#wmiutility-class)  
- [ReadMe](#readme)  
- [RegisterWsc.wsf](#registerwscwsf)  

## WMIUtility class

Improved the [code comments](class/WMIUtility.vbs) and (therefore) the [docs](docs/VBScriptClasses.md#wmiutility) for the WMIUtiltiy class: added links to specific [Computer System Hardware Classes](https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/computer-system-hardware-classes).  

## ReadMe

Added reference links to [ReadMe.md](./ReadMe.md#references). Corrected [link](./ReadMe.md#installation) to [CopyToProgramFiles.vbs](./CopyToProgramFiles.vbs) (formerly CopyToProgramFiles.wsf).

## RegisterWsc.wsf

Corrected comment syntax in [RegisterWsc.wsf](examples/RegisterWsc.wsf), which had no runtime effect.

# Version 1.4.0

- [Configurer class](#configurer-class) - New  
- [CommandParser class](#commandparser-class)  
- [FolderSender class](#foldersender-class) - New  
- [StartupItems class](#startupitems-class)  
- [ArrayOfObjects class](#arrayofobjects-class)  
- [Error numbers](#error-numbers)  
- [LoadObject method](#loadobject-method)  
- [Refactoring](#refactoring)  
- [Change log](#change-log)  
- [Issues](#issues)  

## Configurer class

An improved configuration scheme was implemented with the Configurer class ( [code](class/Configurer.vbs) | [doc](docs/VBScriptClasses.md#configurer) ), using files with the `configure` filename extension. The old scheme uses the `config` filename extension and is still functional.  

The Configurer class uses comma-delimited key/value pairs in the `.configure` files, which are created manually.  Configuration files can be associated with class files, script files, the entire project, or a particular user. Configuration files associated with a class or script or hta take the same base name of the file with which they are associated, and reside in the same folder.

## CommandParser class

The CommandParser class ( [code](class/CommandParser.vbs) | [doc](docs/VBScriptClasses.md#commandparser) ) has been completely reworked for simplicity and testability. This is a breaking change. None of the previous pubic members still work, but the overall functionality is similar.

## FolderSender class

A FolderSender class ( [code](class/FolderSender.vbs) | [doc](docs/VBScriptClasses.md#foldersender) ) was added, formerly FolderCopier, leveraging a rich and familiar Windows-native graphical interface, thanks to the <code> Shell.Application</code> object's CopyHere and MoveHere methods. See [example](CopyToProgramFiles.vbs).

## StartupItems class

The StartupItems class ( [code](class/StartupItems.vbs) | [doc](docs/VBScriptClasses.md#startupitems) ) was moved out of StartItems.hta ( [code](examples/StartItems.hta) ) and into the class folder.

## ArrayOfObjects class

The ArrayOfObjects class  ( [code](class/ArrayOfObjects.vbs) | [doc](docs/VBScriptClasses.md#arrayofobjects) ) was moved out of StartItems.hta ( [code](examples/StartItems.hta) ) and into the class folder.

## LoadObject method

An experimental method LoadObject has been added to the [VBScripting.Includer](docs/VBScriptClasses.md#includer) object. LoadObject doesn't work with all project classes and it doesn't work within a class block. LoadObject is the default member of its class.  

## Error numbers

`Err.Raise` statement error numbers have been updated to conform to published error codes.

For example, this statement

```vb
Err.Raise 1,, "Command-line argument required: a filespec for the file to open."
```

becomes

```vb
Err.Raise 449,, "Command-line argument required: a filespec for the file to open."
```

The second statement above typically will cause a modal message with the code `800A01C1`. Hexadecimal 1C1 can be converted to decimal 449, and then the generic run-time error description can be looked up [online](https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/5ta518cw(v=vs.84)) or in the following table.

| Dec | Hex | Err.Description                           |  
| --: | --: | ------------------------------------- |  
|   5 |   5 | Invalid prodedure call or argument    |  
|  13 |   D | Type mismatch                         |  
|  17 |  11 | Can't perform the requested operation |  
|  51 |  33 | Internal error                        |  
| 449 | 1C1 | Argument not optional                 |  
| 450 | 1C2 | Wrong number of arguments or invalid property assignment |  
| 500 | 1F4 | Variable undefined                    |  
| 505 | 1F9 | Invalid or unqualified reference      |  
| 507 | 1FB | An exception occurred                 |  

## Refactoring

Selected files were refactored:

- Variable declarations were moved closer to the top of the file.  

- Sub Class_Initialize blocks were moved closer to the top of the file, in part to more quickly assess dependencies.  

- The CommandParser class was refactored. See [above](#commandparser-class).  

- Many lines-ending spaces were removed in multiple files. Double-spaces at the end of lines in lists in `.md` files--this line for example--were intentionally retained.  

## Change log

The change log (this file) was added.

## Issues

See [general project issues](ReadMe.md#issues).  
