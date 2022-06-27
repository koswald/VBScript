# Version information

- [Version 1.4.2 (the latest version)](#version-142)  
- [Version 1.4.1](#version-141)  
- [Version 1.4.0](#version-140)  

# Version 1.4.2

- [PushPrep.hta updated](#pushprephta-updated)  
- [ShellSpecialfolders class added](#shellspecialfolders-class-added)  
- [Issues with v1.4.2](#issues-with-v142)  
- [Issues fixed with v1.4.2](#issues-fixed-with-v142)  

## PushPrep.hta updated

Moved the main script to a separate file so that the VBScript code would appear correctly in github.com.

## ShellSpecialFolders class added

Added a new class ShellSpecialFolders ( [code](class/ShellSpecialFolders.vbs) | [docs](docs/VBScriptClasses.md#shellspecialfolders) | [spec](spec/ShellSpecialFolders.spec.vbs) ).  

## Issues with v1.4.2

- [Incomplete Setup exit with PushPrep.hta](#incomplete-setup-exit-with-pushprephta)  

- See [general project issues](ReadMe.md#issues).  

### Incomplete Setup exit with PushPrep.hta

**Bug description:** If Setup.vbs is started from PushPrep.hta and then cancelled from the User Account Control dialog, when elevation of privileges is requested, then Setup.bat must be manually deleted before PushPrep.hta continues. This is a pre-existing issue.  

**Mitigating factors:** Edge case. Similar behavior is not seen and not applicable when Setup.vbs is started another way.

## Issues fixed with v1.4.2

- [No file extension error](#no-file-extension-error)  
- [Error running FolderSender.spec.wsf from %ProgramFiles%](#error-running-foldersenderspecwsf-from-programfiles)  

### No file extension error

**Bug description:** In recent versions, individual integration tests and test suites may fail with the error message, "Input Error: There is no file extension in...&#60;path&#62;", where &#60;path&#62; is the left part of the test or test suite filespec, up to the first space. All of the following conditions must be present in order to duplicate the error:  

- The project version is v1.4.0 or v1.4.1.
- The project path includes a space, as when the project is installed in C:\Program Files\VBScripting.  
- A script/hta restart method is being called, either the RestartWith method of the VBSHoster class or the RestartUsing method of the VBSApp class, as when the TestingFramework class is instantiated but cscript.exe is not the host.  
- The restart method is configured to use PowerShell or Windows PowerShell as the shell.  
- The restart method is called/configured so that a script/hta restart will be attempted. If the desired state is already present, then the restart method will not attempt to restart the script/hta. Desired and current states are compared for 1) privileges, whether they are elevated, and 2) the scripting host, whether cscript.exe or wscript.exe or mshta.exe is hosting the script.  

**Resolution:** This issue was fixed by surrounding the script filespec with single quotes in the powershell command--this is in addition to the double quotes that were already in use. Regression tests were incorporated into VBSHoster.spec.wsf and VBSApp.spec.vbs. 

### Error running FolderSender.spec.wsf from %ProgramFiles%

**Bug description:** With the project located in %ProgramFiles%\VBScripting, the integration test FolderSender.spec.wsf may request that privileges be elevated, which is typically not necessary or appropriate for this test. Only certain tests are intended to be run with elevated privileges, and FolderSender.spec.wsf is not one of them.

**Resolution:** FolderSender.spec.wsf now makes use of the %AppData% folder, where elevated privileges are not required for creating files and folders.

# Version 1.4.1

- [WMIUtility class links added](#wmiutility-class-links-added)  
- [ReadMe links added, corrected)](#readme-links-added-corrected)  
- [RegisterWsc.wsf updated](#registerwscwsf-updated)  

## WMIUtility class links added

Improved the [code comments](class/WMIUtility.vbs) and the [docs](docs/VBScriptClasses.md#wmiutility) for the WMIUtiltiy class: added links to specific [Computer System Hardware Classes](https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/computer-system-hardware-classes).  

## ReadMe links added, corrected

- Added [reference links](./ReadMe.md#references) to ReadMe.md.  

- Corrected [link](./ReadMe.md#installation) to [CopyToProgramFiles.vbs](./CopyToProgramFiles.vbs) (formerly CopyToProgramFiles.wsf).  

## RegisterWsc.wsf updated

Corrected comment syntax in [RegisterWsc.wsf](examples/RegisterWsc.wsf), which had no runtime effect.

# Version 1.4.0

- [Configurer class added](#configurer-class-added) - New  
- [CommandParser class breaking change](#commandparser-class-breaking-change)  
- [FolderSender class added](#foldersender-class-added) - New  
- [CopyToProgramFiles.vbs updated](#copytoprogramfilesvbs-updated)
- [StartupItems class moved](#startupitems-class-moved)  
- [ArrayOfObjects class added](#arrayofobjects-class-added)  
- [Error numbers updated](#error-numbers-updated)  
- [LoadObject method added](#loadobject-method-added)  
- [Numerous files refactored](#numerous-files-refactored)  
- [Change log added](#change-log-added)  

## Configurer class added

An improved configuration scheme was implemented with the Configurer class ( [code](class/Configurer.vbs) | [doc](docs/VBScriptClasses.md#configurer) ), using files with the `configure` filename extension. The old scheme uses the `config` filename extension and is still functional.  

The Configurer class uses comma-delimited key/value pairs in the `.configure` files, which are created manually.  Configuration files can be associated with class files, script files, the entire project, or a particular user. Configuration files associated with a class or script or hta take the same base name of the file with which they are associated, and reside in the same folder.

## CommandParser class breaking change

The CommandParser class ( [code](class/CommandParser.vbs) | [doc](docs/VBScriptClasses.md#commandparser) ) has been completely reworked for simplicity and testability. This is a breaking change. None of the previous pubic members still work, but the overall functionality is similar.

## FolderSender class added

A FolderSender class ( [code](class/FolderSender.vbs) | [doc](docs/VBScriptClasses.md#foldersender) ) was added, leveraging a rich and familiar Windows-native graphical interface, thanks to the <code> Shell.Application</code> object's CopyHere and MoveHere methods. The class was formerly located in CopyToProgramFiles.vbs and was named FolderCopier.

## CopyToProgramFiles.vbs updated

The target folder in [CopyToProgramFiles.vbs](CopyToProgramFiles.vbs) was changed from %ProgramFiles%&#92;KOswald to %ProgramFiles%&#92;VBScripting and the FolderCopier class was renamed to FolderSender and removed to the `class` folder. See [FolderSender class](#foldersender-class).

## StartupItems class moved

The StartupItems class ( [code](class/StartupItems.vbs) | [doc](docs/VBScriptClasses.md#startupitems) ) was moved out of StartItems.hta ( [code](examples/StartItems.hta) ) and into the class folder.

## ArrayOfObjects class added

The ArrayOfObjects class  ( [code](class/ArrayOfObjects.vbs) | [doc](docs/VBScriptClasses.md#arrayofobjects) ) was moved out of StartItems.hta ( [code](examples/StartItems.hta) ) and into the class folder.

## LoadObject method added

An experimental method LoadObject has been added to the [VBScripting.Includer](docs/VBScriptClasses.md#includer) object. LoadObject doesn't work with all project classes and it doesn't work within a class block. LoadObject is the default member of its class.  

## Error numbers updated

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

## Numerous files refactored

Many files were refactored:

- Variable declarations were moved closer to the top of the file.  

- In class blocks, Sub Class_Initialize blocks were moved closer to the top of the file, in part to more quickly assess dependencies.  

- The CommandParser class was refactored. See [above](#commandparser-class).  

- Many lines-ending spaces were removed in multiple files. Double-spaces at the end of lines in lists in `.md` files--this line for example--were intentionally retained.  

## Change log added

The change log (this file) was added.
