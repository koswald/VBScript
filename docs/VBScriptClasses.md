# VBScript Utility Classes Documentation

## Contents

[ArrayOfObjects](#arrayofobjects)  
[Chooser](#chooser)  
[CommandParser](#commandparser)  
[Configurer](#configurer)  
[DocGenerator](#docgenerator)  
[DocGeneratorCS](#docgeneratorcs)  
[EncodingAnalyzer](#encodinganalyzer)  
[EscapeMd](#escapemd)  
[FolderSender](#foldersender)  
[GUIDGenerator](#guidgenerator)  
[HTAApp](#htaapp)  
[Includer](#includer)  
[KeyDeleter](#keydeleter)  
[MathConstants](#mathconstants)  
[MathFunctions](#mathfunctions)  
[NameValue](#namevalue)  
[PrivilegeChecker](#privilegechecker)  
[RegExFunctions](#regexfunctions)  
[RegistryUtility](#registryutility)  
[SetupHelper](#setuphelper)  
[ShellConstants](#shellconstants)  
[ShellSpecialFolders](#shellspecialfolders)  
[SpecialFolders](#specialfolders)  
[StartupItems](#startupitems)  
[StringFormatter](#stringformatter)  
[TestingFramework](#testingframework)  
[TextStreamer](#textstreamer)  
[TimeFunctions](#timefunctions)  
[ValidFileName](#validfilename)  
[VBSApp](#vbsapp)  
[VBSArguments](#vbsarguments)  
[VBSArrays](#vbsarrays)  
[VBSClipboard](#vbsclipboard)  
[VBSEnvironment](#vbsenvironment)  
[VBSEventLogger](#vbseventlogger)  
[VBSExtracter](#vbsextracter)  
[VBSFileSystem](#vbsfilesystem)  
[VBSHoster](#vbshoster)  
[VBSLogger](#vbslogger)  
[VBSMessages](#vbsmessages)  
[VBSPower](#vbspower)  
[VBSStopwatch](#vbsstopwatch)  
[VBSTestRunner](#vbstestrunner)  
[VBSTroubleshooter](#vbstroubleshooter)  
[VBSValidator](#vbsvalidator)  
[WindowsUpdatesPauser](#windowsupdatespauser)  
[WMIUtility](#wmiutility)  
[WoWChecker](#wowchecker)  

## ArrayOfObjects

The default property of the ArrayOfObjects class, Items, acts like a rudimentary C# ArrayList.  
  
Example  
```vb
 Option Explicit
 Dim aoo 'ArrayOfObjects object
 Dim incl 'VBScripting.Includer object
 Initialize
 Add "tree", "pear"
 Add "tree", "walnut"
 ShowAll
 ShowAll2
 Sub Initialize
     Set incl = CreateObject( "VBScripting.Includer" )
     Execute incl.Read( "ArrayOfObjects" )
     Set aoo = New ArrayOfObjects
 End Sub
 Sub Add( noun, example )
     Execute incl.Read( "NameValue" )
     aoo.Add New NameValue.Init( noun, example )
 End Sub
 Sub ShowAll
     Dim obj, s
     For Each obj In aoo() 'or aoo.Items or aoo.Items()
         s = s & obj.Name & vbTab & obj.Value & vbLf
     Next
     MsgBox s,, "ShowAll"
 End Sub
 Sub ShowAll2
     Dim i, s
     For i = 0 To UBound(aoo) 'or aoo() or aoo.Items or aoo.Items()
         s = s & aoo()(i).Name & vbTab & aoo()(i).Value & vbLf
     Next
     MsgBox s,, "ShowAll2"
 End Sub
```
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Property | Items | None | an array of objects | Returns an array of the objects that were added using the Add method. This is default property, so the name (Items) may not need to be specified. However, it may be necessary to add empty parens to the object name: See the example. |
| Method | Add | an object | N/A | Expands the Items array and adds the specified object to it. |
| Property | Count | None | an integer | Returns the number of items in the Items array. |

## Chooser

Get a folder or file chosen by the user  
  
<strong> Deprecated</strong> in favor of the <a target="_blank" href="https://github.com/koswald/VBScript/blob/master/.Net/ReadMe.md#user-content-overview"> .NET extensions</a> VBScripting.FolderChooser ( <a target="_blank" href="https://github.com/koswald/VBScript/blob/master/.Net/FolderChooser.cs"> code</a> &#124; <a target="_blank" href="https://github.com/koswald/VBScript/blob/master/docs/CSharpClasses.md#user-content-folderchooser"> doc</a> ) and VBScripting.FileChooser ( <a target="_blank" href="https://github.com/koswald/VBScript/blob/master/.Net/FileChooser.cs"> code</a> &#124; <a target="_blank" href="https://github.com/koswald/VBScript/blob/master/docs/CSharpClasses.md#user-content-filechooser"> doc</a> ), which are more versatile and user friendly.  
  
Usage example  
  
```vb
 With CreateObject( "VBScripting.Includer" ) 
     Execute .Read( "Chooser" )
 End With 

 Dim choose : Set choose = New Chooser 
 MsgBox choose.folder 
 MsgBox choose.file 
```
  
Browse for file <a target="_blank" href="http://stackoverflow.com/questions/21559775/vbscript-to-open-a-dialog-to-select-a-filepath"> reference</a>.  
Browse for folder <a target="_blank" href="http://ss64.com/vb/browseforfolder.html"> reference</a>.  
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Property | File | None | a file path | Opens a Choose File dialog and returns the path of a file chosen by the user. Returns an empty string if no folder was selected. Note: The title bar text will say Choose File to Upload. |
| Property | Folder | None | a folder path | Opens a Browse For Folder dialog and returns the path of a folder chosen by the user. Returns an empty string if no folder was selected. |
| Property | FolderTitle | None | the folder title | Opens a Browse For Folder dialog and returns the title of a folder chosen by the user. The title for a normal folder is just the folder name. For a special folder like %UserProfile%, it may be something entirely different. Returns an empty string if no folder was selected. |
| Property | FolderObject | None | an object | Opens a Browse For Folder dialog and returns a Shell.Application BrowseForFolder object for a folder chosen by the user. This object has methods Title and Self.Path, corresponding to this class's FolderTitle and FolderPath, respectively. This method is recommended for when you need both the FolderTitle and FolderPath but only want the user to have to choose once. If no folder was selected, then TypeName(folderObj) = "Nothing" is True. |
| Method | SetWindowTitle | a string | N/A | Sets the title of the Browse For Folder window: i.e. the text below the titlebar. |

## CommandParser

The CommandParser class' Result method runs a command and searches its output for a phrase.  
  
Example:  
```vb
 Dim includer : Set includer = CreateObject( "VBScripting.Includer" ) 
 Execute includer.Read( "CommandParser" ) 
 Dim cp : Set cp = New CommandParser 
 Dim cmd : cmd = "cmd /c If defined ProgramFiles^(X86^) (echo 64-bit) else (echo 32-bit)" 
 Dim phrase : phrase = "64-bit" 
 MsgBox cp.Result( cmd, phrase ) 'True expected for 64-bit systems
```
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Property | Result | cmd, phrase | a boolean | Runs the specified command and returns a boolean: True if the specified phrase is found in the output of the specified command. Not case sensitive by default. |
| Property | CaseSensitive | a boolean | a boolean | Gets or sets whether the search is case sensitive. Default is False.  |

## Configurer

Allows for keeping configuration data for a class or script separate from the code.  
  
Requirements  
1. The configuration files are manually created with comma-delimited key/value pairs that are read/loaded into a dictionary and accessed through the Item property.  
2. The configuration files must have the <code>configure</code> filename extension. See LoadUserConfig for the one exception.  
3. The configuration files must have the same base name as the associated class file or calling script. Two exceptions: the UserConfigFile and GlobalConfigFile do not have base names.  
4. The configuration file for a script must be located in the same folder as the script.  
5. The configuration file for a class should be in the project's <code>class</code> folder, or else in another folder that is specified by the LibraryPath property. If using another folder, then the LibraryPath property must be set before calling the LoadClassConfig method or getting the ClassConfigFile property.  
6. The configuration files can have in-line or whole-line # comments.  
7. Leading and trailing whitespace is ignored in both the key and the value.  
  
 Note: Three config files GlobalConfigFile, UserConfigFile, and ScriptConfigFile, are loaded in that order on instantiation of the Configure class. The most recently loaded file takes precedence if there is a conflict, so if a different precedence is desired, then the files can be reloaded in a different order. A class configuration file is loaded by the <code> LoadClassConfig</code> method or the <code> LoadFile</code> method.  
  
Example:  
  
```vb
 'Test1.vbs (located anywhere)
 With CreateObject( "VBScripting.Includer" )
     Execute .Read( "Configurer" )
 End With
 With New Configurer
     If .Exists( "command1" ) Then
         MsgBox "command1: " & .Item( "command1" )
     Else MsgBox "command1 key not found."
     End If
 End With
```
  
<code> # Test1.configure (located in the same folder as Test1.vbs)</code>  
<code> command1, wt powershell # requires Windows Terminal</code>  
  
```vb
 'Test2.vbs (located in the "class" folder)
 Class Test2
     Sub Class_Initialize
         With CreateObject( "VBScripting.Includer" )
             Execute .Read( "Configurer" )
         End With
         With New Configurer
             .LoadClassConfig me
             If .Exists( "command2" ) Then
                 MsgBox .Item( "command2" )
             End If
         End With
     End Sub
 End Class
```
  
<code> # Test2.configure (also located in the "class" folder)</code>  
<code> command2, pwsh # requires PowerShell 6 or higher</code>  
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Property | Item | a key (string) | a value (string) | Returns the value of the key/value pair for the specified key. Returns Empty if the key is not found. |
| Property | Count | None | an integer | Gets the number of key/value pairs in the Configurer dictionary. |
| Property | Exists | a string (key) | a boolean | Gets whether a given key/value pair exists in the Configurer dictionary. Parameter is the key. |
| Property | Dictionary | None | an object reference | Returns a reference to the Configurer object's dictionary object. Properties: CompareMode, Item, Key. Methods: Add, Exists, Items, Keys, Remove, RemoveAll. See the <a target="_blank" href="https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/x4k5wbx4(v=vs.84)"> online docs</a> for the Dictionary object. |
| Method | LoadFile | a filespec | N/A | Loads the specified configuration file's key/value pairs into the object's dictionary. See Item property. See also the LoadClassConfig and LoadScriptConfig methods. |
| Method | LoadScriptConfig | None | N/A | Loads the configuration file associated with the calling script. The configuration file's key/value pairs are added to the Configurer object's dictionary object, or if the key exists already, the value is updated. |
| Property | ScriptConfigFile | None | a filespec | Returns the filespec of the configuration file associated with the script that is using the Configurer object, the calling script or .hta. The file doesn't have to exist. |
| Method | LoadClassConfig | a string or an object reference | N/A | Loads the configuration file associated with a class file. The configuration file's key/value pairs are added to the Configurer object's dictionary object, or if the key exists already, the value is updated. The parameter may be 1) the class name, or 2) an object reference to an instance of the class, or 3) the keyword me, if called from within the class. |
| Property | ClassConfigFile | a string or an object reference. | a filespec | Returns the filespec of the configuration file associated with a class (.vbs) file. The file doesn't have to exist. The parameter may be 1) the class name, or 2) an object reference to an instance of the class, or 3) the keyword me, if called from within the class. |
| Method | LoadUserConfig | None | N/A | Loads the user configuration file at <code>%UserProfile%&#92;.VBScripting</code>. See Note for UserConfigFile. |
| Property | UserConfigFile | None | a filespec | Returns the filespec of a user-specific configuration file, related to the project but outside of the project folders, at <code>%UserProfile%&#92;.VBScripting</code>. The file doesn't have to exist. Note: Care should be taken when privileges are elevated and the user is not a member of the Administrators group, because as privileges are elevated, %UserProfile% changes. |
| Method | LoadGlobalConfig | None | N/A | Loads the configuration file in the project folder. See comments for the GlobalConfigFile property. Equivalent to calling <code>LoadFile GlobalConfigFile</code>. |
| Property | GlobalConfigFile | None | a filespec | Returns the filespec of the global configuration file. The word global refers to the project only. Depending on the location of the project, the configuration file may or may not be accessible to all users. The file does not have to exist. Expected value: <code>&lt;project folder&gt;&#92;.configure</code>. |
| Property | LibraryPath | None | a path | Gets or sets the location, i.e. the parent folder, of the class file and/or its associated configuration file. See the LoadClassConfig and LoadFile methods. Obscure. For an example, see the integration test Configurer.spec.wsf. |
| Property | ToArray | a string | an array | Converts a string to an array. Uses the delimiter set by the Delimiter property, a vertical bar ( &#124; ) by default. Excess spaces on the left and right of each element are trimmed off. |
| Property | PowerShell | None | a string | Returns a string useful for starting a PowerShell process. If PowerShell 6 or 7 is installed, then the return value is the expanded filespec of the first "pwsh candidates" executable found that is listed in the file <code>.configure</code> in the project's root folder. If the cross-platform PowerShell is not found, returns the string <code>powershell</code>, which may be used to start a Windows PowerShell process. Since the return value may contain spaces, the string may need to be surrounded by quotes, depending on how it is used. For example, if the return value is used as the first argument of the Shell.Appliction object's ShellExecute method, then quotes are not recommended. But if the return value is used in the first argument of the WScript.Shell object's Run method, then quotes are recommended. |
| Property | WT | None | a string | Returns the filespec of a Windows Terminal executable, if installed and listed in <code>.configure</code> in the project folder. Returns <code>Empty</code> if Windows Terminal is not installed or not found. |
| Property | Delimiter | None | a character | Gets or sets the delimiter used in converting strings to arrays. Default is a vertical bar ( &#124; ). |
| Property | PsFallback | None | a string | Returns a ten-character string suitable for starting a Windows PowerShell process: <code>powershell</code>. This becomes the default PowerShell when the newer cross-platform PowerShell is not installed or not found. |
| Property | Init | an object | an object self-reference | Initializes the Configurer object so that it can find the name of the calling script. The parameter is the WScript object, for .vbs or .wsf files, or the 'Document' object for .hta files. Required if the Configurer object was instantiated with the <a href="#includer"> VBScripting.Includer</a> object's experimental LoadObject method. Example: <pre> With CreateObject( "VBScripting.Includer" )<br />     Set c = .LoadObject( "Configurer" ).Init( WScript )<br /> End With</pre> |

## DocGenerator

Generate html and markdown documentation for VBScript code based on well-formed code comments.  
Usage Example  
```vb
 With CreateObject( "VBScripting.Includer" )
     Execute .Read( "DocGenerator" )
 End With
 With New DocGenerator
     .SetTitle "VBScript Utility Classes Documentation"
     .SetDocName "VBScriptClasses"
     .SetFilesToDocument "*.vbs | *.wsf | *.wsc"
     .SetScriptFolder "..\class"
     .SetDocFolder "..\docs"
     .Generate
     .ViewMarkdown
 End With
```
  
Example of well-formed comments before a Sub statement  
 Note: A remark is required for Methods (Subs).  
  
```vb
'Method: SubName
'Parameters: param1Name, param2Name
'Remark: Details about the method and parameters.
```
Example of well-formed comments before a Property or Function statement.  
Note: A Returns (or Return or Returns: or Return:) is required with a Property or Function.  
  
```vb
'Property: PropertyName
'Returns: a string
'Remark: A remark is not required for a Property or Function, but usually is a good idea.
```
Notes for the comment syntax at the beginning of a script  
Use a single quote ( ' ) for general comments <br />  
- use a single quote by itself for an empty line <br />  
- Wrap VBScript code with <code>pre</code> tags, separating multiple lines with &lt;br /&gt;. <br />  
- Wrap other code with <code> code</code> tags, with each line surrounded with <code> code</code> tags.  
  
Use three single quotes for remarks that should not appear in the documentation <br />  
  
Use four single quotes ( '''' ), if the script doesn't contain a class statement, to separate the general comments at the beginning of the file from the rest of the file.  
  
For some characters to render correctly, they may need to be replaced by escape codes, even when used within &#60;code&#62; or &#60;pre&#62; tags:  
 - For &#124; use &#38;#124; (vertical bar)  
 - For &#60; use &#38;#60; (less than)  
 - For &#62; use &#38;#62; (greater than)  
 - For &#92; use &#38;#92; (backslash)  
 - For &#38; use &#38;#38; (ampersand)  
 - For other characters,  <code>examples\HTML_EscapeCodes.hta</code> can be used to generate an escape code that works with both of the generated files: Markdown and HTML. The numerical portion of the escape code is returned by the VBScript function Asc.  
  
Visual Studio and VS Code extensions may render Markdown files differently than Git-Flavored Markdown.  
  
Issues:  
- Introductory comments at the beginning of a class file should be followed by a line containing a single quote character, or else the markdown table may not render correctly.  
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Method | SetScriptFolder | a folder | N/A | Required. Must be set before calling the Generate method. Sets the folder containing the scripts to include in the generated documentation. Environment variables OK. Relative paths OK. |
| Method | SetDocFolder | a folder | N/A | Required. Must be set before calling the Generate method. Sets the folder of the documentation file. Environment variables OK. Relative paths OK. |
| Method | SetDocName | a filename | N/A | Required. Must be set before calling the Generate method. Specifies the name of the documentation file. Do not include the extension name. |
| Method | SetTitle | a string | N/A | Required. Must be set before calling the Generate method. Sets the title for the documentation. |
| Method | SetFilesToDocument | wildcard(s) | N/A | Specifies which files to document. Optional. Default is <strong> *.vbs </strong>. Separate multiple wildcards with &#124; |
| Method | Generate | None | N/A | Generate comment-based documentation for the scripts in the specified folder. |
| Method | View | None | N/A | Open the html document in the default viewer. Same as ViewHtml. |
| Method | ViewHtml | None | N/A | Open the html document in the default viewer. Same as View method. |
| Method | ViewMarkdown | None | N/A | Open the markdown document in the default viewer. |
| Property | Colorize | boolean | boolean | Gets or sets whether &lt;pre&gt; code blocks (assumed to be VBScript) in the markdown document are colorized. If False (experimental, with Git Flavored Markdown), the code lines should not wrap. Default is True. |

## DocGeneratorCS

 DocGeneratorCS class  
  
 Generates html and markdown documentation for C# code from compiler-generated xml files based on three-slash ( /// ) code comments.  
  
 Four base tags are supported: summary, parameters, returns, and remarks. Within these tags, html tags are allowed, although Markdown typically does not render all html tags.  
  
 Note: When changes are made to source-code comments, the code must be compiled again in order for new .xml files to be generated, before running the doc-generator script.  
  
 Note: Html tags may result in malformed markdown table rows when there is whitespace between adjacent tags.  
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Property | XmlFolder | folder | folder | Required. Sets (or gets) the folder containing the .xml files autogenerated by the C# compiler. Relative paths and environment variables are supported. |
| Property | OutputFile | filespec | filespec | Required. Sets (or gets) the path and base name of the output files. Do not include the .html or .md extension name: they will be added automatically. Older versions, if any, will be overwritten. Relative paths and environment variables are supported. |
| Method | Generate | None | N/A | Generates html and markdown code documentation. Requires .xml files to have been generated by the C# compiler. |
| Method | ViewHtml | None | N/A | Opens the html document with the default viewer. |
| Method | ViewMarkdown | None | N/A | Opens the markdown document with the default viewer. |

## EncodingAnalyzer

Provides various properties to analyze a file's encoding.  
  
FOR ILLUSTRATION PURPOSES ONLY. The algorithm used assumes that there is a Byte Order Mark, which in many cases is a wrong assumption.  
  
Usage example  
```vb
With CreateObject( "VBScripting.Includer" )
    Execute .Read( "EncodingAnalyzer" )
End With
 
With New EncodingAnalyzer.SetFile(WScript.Arguments(0))
    MsgBox "isUTF16LE: " & .isUTF16LE
End With
```
  
Stackoverflow references: <a target="_blank" href="http://stackoverflow.com/questions/3825390/effective-way-to-find-any-files-encoding"> 1</a>, <a target="_blank" href="http://stackoverflow.com/questions/1410334/filesystemobject-reading-unicode-files"> 2</a>.  
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Property | SetFile | a filespec | an object self reference | Required. Specifies the file whose encoding is to be determined. Relative paths are permitted, relative to the current directory. |
| Property | isUTF16LE | None | a boolean | Returns a boolean indicating whether the file specified by SetFile is Unicode Little Endian, <strong> aka Unicode</strong>. |
| Property | isUTF16BE | None | a boolean | Returns a boolean indicating whether the file specified by SetFile is Unicode Big Endian. |
| Property | isUTF7 | None | a boolean | Returns a boolean indicating whether the file specified by SetFile is UTF7. |
| Property | isUTF8 | None | a boolean | Returns a boolean indicating whether the file specified by SetFile is UTF8. |
| Property | isUTF32 | None | a boolean | Returns a boolean indicating whether the file specified by SetFile is UTF32. |
| Property | isAscii | None | a boolean | Returns a boolean indicating whether the file specified by SetFile is Ascii. |
| Property | GetType | None | a string | Returns one of the following strings according the format of the file set by SetFile: Ascii, UTF16LE, UTF16BE, UTF7, UTF8, UTF32. |
| Property | GetCurrentDirectory | None | a folder | Returns the current directory |
| Method | SetCurrentDirectory | a folder | N/A | Sets the current directory. |
| Property | GetByte | BOM byte number | an integer | Returns the Ascii value, 0 to 255, of the byte specified. The parameter must be an integer: one of 0, 1, 2, or 3. These represent the first four bytes in the file, the Byte Order Mark (BOM). |

## EscapeMd

EscapeMd and EscapeMd2 Functions  
Escape markdown special characters.  
Usage example  
```vb
    Dim includer : Set includer = CreateObject( "VBScripting.Includer" )
    ExecuteGlobal includer.Read( "EscapeMD" )
    MsgBox EscapeMd("```") ' \`\`\`
```
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Property | EscapeMd | unescaped string | escaped string | Returns a string with Markdown special characters escaped. |
| Property | EscapeMd2 | unescaped string | escaped string | Returns a string with a minimal amount of Markdown special characters escaped. <a target="_blank" href="http://www.theukwebdesigncompany.com/articles/entity-escape-characters.php"> Escape codes</a>. |

## FolderSender

The FolderSender class supplies methods that copy or move (send) the specified SourceFolder to the specified TargetFolder. Operator action may be  required.   
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Method | Copy | None | N/A | Copies a folder. The SourceFolder and TargetFolder properties must be specified in advance or else an error will occur. A familiar Windows-native graphical interface appears for sizeable operations or when it is necessary to overwrite existing files or to elevate privileges: Operator action may be required.  |
| Method | Move | None | N/A | Moves a folder. The SourceFolder and TargetFolder properties must be specified in advance or else an error will occur. A familiar Windows-native graphical interface appears for sizeable operations or when it is necessary to overwrite existing files or to elevate privileges: Operator action may be required.  |
| Property | SourceFolder | a string (folder) | a string (folder) | Required. Sets or gets the source folder for the Copy and Move methods. Relative paths are allowed. Environment variables are allowed. The source folder must exist or an error will occur. |
| Property | TargetFolder | a string (folder) | a string (folder) | Required. Sets or gets the target folder for the Copy and Move methods. Relative paths are allowed (see the CurrentDirectory property). Environment variables are allowed. The target folder will be created if it does not exist. The User Account Control dialog may appear to request permission to create a folder if it is in a location that has restricted write permissions such as %ProgramFiles%. |
| Property | CurrentDirectory | a string (folder) | a string (folder) | Gets or sets the current directory or working directory. Relative paths are allowed. Environment variables are allowed. |

## GUIDGenerator

Generate a unique GUID  
Usage example  
```vb
 With CreateObject( "VBScripting.Includer" )
     Execute .Read( "GUIDGenerator" )
 End With
 InputBox "",, New GUIDGenerator
```
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Property | Generate | None | a GUID | Returns a unique GUID. Generate is the default property for the class, so the property name is optional. A sample GUID: {928507A9-7958-4E6E-A0B1-C33A5D4D602A} |
| Method | SetUppercase | None | N/A | Configure the Generate property to return uppercase, the default. |
| Method | SetLowercase | None | N/A | Configure the Generate property to return lowercase |

## HTAApp

HTAApp class  
Supports the VBSApp class, providing .hta functionality. *Intended for use only within the VBSApp class*.  
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Method | Sleep | an integer | N/A | Pauses execution of the script or .hta for the specified number of milliseconds. |
| Method | PrepareToSleep | None | N/A | Required before calling the Sleep method when AlwaysPrepareToSleep is False in HTAApp.config. |
| Property | GetFilespec | None | a string | Returns the filespec of the calling .hta file. |
| Property | GetArgs | None | an array | Returns the mshta.exe command line args as an array, including the .hta filespec, which has index 0. |

## Includer

  
The Includer object helps with dependency management, and can be used in a .wsf, .vbs, or .hta script.  
  
How it works: The Read property returns the contents of a .vbs class file--or any other text file.  
  
Usage example  
```vb
 With CreateObject( "VBScripting.Includer" )
     Execute .Read( "WMIUtility.vbs" ) '.vbs may be omitted
     Execute .Read( "TextStreamer" )
 End With
 Dim wmi : Set wmi = New WMIUtility
 Dim streamer : Set streamer = New TextStreamer 
```
  
Relative paths may be used and are relative to the location of the class folder.  
  
Registration  
  
Although Windows Script Component (.wsc) files must be registered--unless used with GetObject("script:" & AbsolutePathToWscFile)--right clicking <code> Includer.wsc</code> and selecting Register probably <strong> will not work</strong>. Instead,  
1) Run the Setup.vbs in the project folder. Or,  
2) Run the following commands in a command window with elevated privileges. The first command applies to 64-bit systems and 32-bit systems. The second command applies only to 64-bit systems.  
  
<code>     %SystemRoot%\System32\regsvr32.exe &lt;absolute-path-to&gt;\Includer.wsc </code> <br /> <code>     %SystemRoot%\SysWow64\regsvr32.exe &lt;absolute-path-to&gt;\Includer.wsc </code>  
  
<a target="_blank" href="http://github.com/koswald/VBScript/blob/master/class/wsc/ReadMe.md#user-content-registration">Alternate registration method</a>.  
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Property | GetObj | class name | An object | Returns an object based on the VBScript class with the specified name. Requires a .wsc Windows Script Component file in \class\wsc. The object does not need to be registered, although the VBScripting.Includer (this) object must be registered. See StringFormatter.wsc for an example. |
| Property | LoadObject | class name | an object | Experimental. Returns an object based on a class (.vbs) file located in the project's <code> class</code> folder. The parameter is the class name, which is also the base name of the class .vbs file. Classes having an Init method may need to have the WScript object or the Document object passed in, using the Init method, before calling certain procedures. See the Configurer and VBSApp classes for examples of using an Init method in this way. Experimental. Does not work well when used within a Class block. |
| Property | Read | a file | the file contents | Returns the contents of the specified file, which may be expressed either as an abolute path, or as a relative path relative to the <code> class</code> folder. The file name extension may be omitted for .vbs files. |
| Property | ReadFrom | file, path | file contents | Returns the contents of the specified file, which may be expressed either as an abolute path, or as a relative path relative to the path specified. The file name extension may be omitted for .vbs files. |
| Property | LibraryPath | None | a folder path | Returns the resolved, absolute path of the <code> class</code> folder, which is the reference for relative paths passed to the Read method. |

## KeyDeleter

The KeyDeleter class provides a method for deleting a registry key and all of its subkeys.  
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Method | DeleteKey | root, key | N/A | Deletes the specified registry key and all of its subkeys. Use one of the root constants for the first parameter. |
| Property | HKCR | None | &H80000000 | Provides a value suitable for the first parameter of the DeleteKey method. |
| Property | HKCU | None | &H80000001 | Provides a value suitable for the first parameter of the DeleteKey method. |
| Property | HKLM | None | &H80000002 | Provides a value suitable for the first parameter of the DeleteKey method. |
| Property | HKU | None | &H80000003 | Provides a value suitable for the first parameter of the DeleteKey method. |
| Property | HKCC | None | &H80000005 | Provides a value suitable for the first parameter of the DeleteKey method. |
| Property | Result | None | an integer | Returns a code indicating the result of the most recent DeleteKey call. Codes can be looked up in <a target="_blank" href="https://docs.microsoft.com/en-us/windows/desktop/api/wbemdisp/ne-wbemdisp-wbemerrorenum">WbemErrEnum</a> or <a target="_blank" href="https://docs.microsoft.com/en-us/windows/win32/wmisdk/wmi-error-constants">WMI Error Constants</a>. |
| Property | Delete | a boolean | a boolean | Gets or sets the boolean that controls whether the key is actually deleted. Default is True. Used for testing. |

## MathConstants

| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Property | pi | None | 3.14159265358979 | pi can be generated by the expression <code> 4 * Atn(1)</code>. |
| Property | DegRad | None | pi/180 | To convert degrees to radians, multiply degrees by DegRad. |
| Property | RaDeg | None | 180/pi | To convert radians to degrees, multiply radians by RaDeg. Same as RadDeg. Included for backwards compatibility. |
| Property | RadDeg | None | 180/pi | To convert radians to degrees, multiply radians by RadDeg. |
| Property | e | None | 2.71828182845905 | <em> e</em> can be generated by the expression <code> Exp( 1 )</code>. |

## MathFunctions

The MathFunctions class provides math functions not native to VBScript.  
These functions are derived from functions that are native to VBScript: Sin, Cos, Tan, Atn, and Log.  
  
Log is base <em> e</em>. Angles are in radians. Convert from degrees to radians by multiplying by pi/180.  
Adapted from <a href=https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/keywords/derived-math-functions> Derived Math Functions (Visual Basic)</a>. See also the <a target =_blank href=#mathconstants> MathConstants</a> class.  
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Property | pi | None | 3.14159265358979 | pi can be generated by the expression <code> 4 * Atn(1)</code>. |
| Property | DegRad | None | pi/180 | To convert degrees to radians, multiply degrees by DegRad. |
| Property | RaDeg | None | 180/pi | To convert radians to degrees, multiply radians by RaDeg. Same as RadDeg. |
| Property | RadDeg | None | 180/pi | To convert radians to degrees, multiply radians by RadDeg. |
| Property | e | None | 2.71828182845905 | <em> e</em> can be generated by the expression <code> Exp( 1 )</code>. |
| Property | Sec | Angle in radians | Secant | Sec = 1 / Cos(X) |
| Property | Cosec | Angle in radians | Cosecant | Cosec = 1 / Sin(X) |
| Property | Cotan | Angle in radians | Cotangent | Cotan = 1 / Tan(X) |
| Property | Arcsin | A ratio | Arcsine | Arcsin = Atn(X / Sqr(-X * X + 1)) |
| Property | Arccos | A ratio | Inverse Cosine | Arccos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1) |
| Property | Arcsec | A ratio | Inverse Secant | Arcsec = Atn(X / Sqr(X * X - 1)) + Sgn((X) -1) * (2 * Atn(1)) |
| Property | Arccosec | A ratio | Inverse Cosecant | Arccosec = Atn(X / Sqr(X * X - 1)) + (Sgn(X) - 1) * (2 * Atn(1)) |
| Property | Arccotan | A ratio | Inverse Cotangent | Arccotan = Atn(X) + 2 * Atn(1) |
| Property | HSin | Hyperbolic angle | Hyperbolic Sine | HSin = (Exp(X) - Exp(-X)) / 2 |
| Property | HCos | Hyperbolic angle | Hyperbolic Cosine | HCos = (Exp(X) + Exp(-X)) / 2 |
| Property | HTan | Hyperbolic angle | Hyperbolic Tangent | HTan = (Exp(X) - Exp(-X)) / (Exp(X) + Exp(-X)) |
| Property | HSec | Hyperbolic angle | Hyperbolic Secant | HSec = 2 / (Exp(X) + Exp(-X)) |
| Property | HCosec | Hyperbolic angle | Hyperbolic Cosecant | HCosec = 2 / (Exp(X) - Exp(-X)) |
| Property | HCotan | Hyperbolic angle | Hyperbolic Cotangent | HCotan = (Exp(X) + Exp(-X)) / (Exp(X) - Exp(-X)) |
| Property | HArcsin | X | Inverse Hyperbolic Sine of X | HArcsin = Log(X + Sqr(X * X + 1)) |
| Property | HArccos | X | Inverse Hyperbolic Cosine of X | HArccos = Log(X + Sqr(X * X - 1)) |
| Property | HArctan | X | Inverse Hyperbolic Tangent of X | HArctan = Log((1 + X) / (1 - X)) / 2 |
| Property | HArcsec | X | Inverse Hyperbolic Secant of X | HArcsec = Log((Sqr(-X * X + 1) + 1) / X) |
| Property | HArccosec | X | Inverse Hyperbolic Cosecant of X | HArccosec = Log((Sgn(X) * Sqr(X * X + 1) +1) / X) |
| Property | HArccotan | X | Inverse Hyperbolic Cotangent of X | HArccotan = Log((X + 1) / (X - 1)) / 2 |
| Property | LogN | X, N | Logarithm of X to base N | LogN = Log(X) / Log(N) |

## NameValue

The NameValue class has two properties, Name and Value, which can be used, for example, to describe a startup item in the registry Run key. See the <a href="#startupitems"> StartupItems class</a>.  
  
```vb
 With CreateObject( "VBScripting.Includer" )
     Execute .Read( "NameValue" )
 End With
 Set obj = New NameValue.Init( "age", 70 )
```
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Property | Name | a variant | a variant | None |
| Property | Value | a variant | a variant | None |
| Property | Init | name, value | an object self reference | Initializes the object. The Init property returns an object self reference, so an object may be instantiated and initialized in the same statement. See the example. See the <a target="_blank" href="https://github.com/koswald/VBScript/blob/master/class/NameValue.vbs"> code</a>. |

## PrivilegeChecker

The default property of the PrivilegeChecker class, Privileged, returns True if the calling script has elevated privileges.  
Usage example  
```vb
 With CreateObject( "VBScripting.Includer" ) 
     Execute .Read( "PrivilegeChecker" ) 
 End With 
 Dim pc : Set pc = New PrivilegeChecker 
 If pc Then 
     WScript.Echo "Privileges are elevated" 
 Else 
     WScript.Echo "Privileges are not elevated" 
 End If 
```
  
Reference: <a target="_blank" href="http://stackoverflow.com/questions/4051883/batch-script-how-to-check-for-admin-rights/21295806"> stackoverflow.com</a>  
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Property | Privileged | None | a boolean | Returns True if the calling script is running with elevated privileges, False if not. Privileged is the default property. |

## RegExFunctions

Regular Expression functions - a work in progress  
  
Usage example  
```vb
  With CreateObject( "VBScripting.Includer" )
      Execute .Read( "RegExFunctions" )
  End With
  
  Dim reg : Set reg = New RegExFunctions
  reg.SetTestString "'Method SetSomething"
  reg.SetPattern "(M).*(od).*(tS)"
  
  Dim s, submatch, subs : s = ""
  Set subs = reg.GetSubMatches
  
  For Each submatch In subs
      s = s & " " & submatch
  Next
  MsgBox s 'M od tS 
```
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Property | Pattern | wildcard | a regex expression | Returns a regular expression equivalent to the specified wildcard expression(s). Delimit multiple wildcards with a vertical bar ( &#124; ). See <a href=https://github.com/koswald/VBScript/blob/master/docs/algorithm/ReadMe.md target=_blank> algorithm/ReadMe.md</a> for more comments. |
| Property | re | None | an object reference | Returns a reference to the RegExp object instance. |
| Method | SetPattern | a regex pattern | N/A | Required before calling FirstMatch or GetSubMatches. Sets the pattern of the RegExp object instance. |
| Method | SetTestString | a string | N/A | Required before calling FirstMatch or GetSubMatches. Specifies the string against which the regex pattern will be tested. |
| Method | SetIgnoreCase | a boolean | N/A | Optional. Specifies whether the regex object will ignore case. Default is False. |
| Method | SetGlobal | a boolean | N/A | Optional. Specifies whether the pattern should match all occurrences in the search string or just the first one. Default is False. |
| Property | GetSubMatches | None | an object | Returns the RegExp SubMatches object for the specified pattern and test string. The matches can be accessed with a For Each loop. See general usage comments. Work in progress. You must handle errors in case there are no matches. |
| Property | FirstMatch | None | a string | Regarding the string specified by SetTestString, returns the first substring in the string that matches the regex pattern specified by SetPattern. |

## RegistryUtility

Provides functions relating to the Windows&reg; registry  
  
Usage example  
```vb
  With CreateObject( "VBScripting.Includer" ) 
      Execute .Read( "RegistryUtility" ) 
  End With 
  Dim reg : Set reg = New RegistryUtility 
  Dim key : key = "SOFTWARE\Microsoft\Windows NT\CurrentVersion" 
  MsgBox reg.GetStringValue( reg.HKLM, key, "ProductName" ) 
```
  
Set valueName to vbEmpty or "" (two double quotes) to specify a key's default value.  
  
StdRegProv docs <a target="_blank" href="https://docs.microsoft.com/en-us/previous-versions/windows/desktop/regprov/stdregprov"> online</a>.  
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Method | SetPC | a computer name | N/A | Optional. A dot (.) can be used for the local computer (default), in place of the computer name. |
| Property | Reg | None | an object | Returns a reference to the StdRegProv object. |
| Property | GetStringValue | rootKey, subKey, valueName | a string | Returns the value of the specified registry location. The specified registry entry must be of type string (REG_SZ). |
| Method | SetStringValue | rootKey, subKey, valueName, value | N/A | Writes the specified REG_SZ value to the specified registry location. Writing to HKLM or HKCR requires elevated privileges. |
| Property | GetExpandedStringValue | rootKey, subKey, valueName | a string | Returns the value of the specified registry location. The specified registry entry must be of type REG_EXPAND_SZ. |
| Method | SetExpandedStringValue | rootKey, subKey, valueName, value | N/A | Writes the specified REG_EXPAND_SZ value to the specified registry location. Writing to HKLM or HKCR requires elevated privileges. |
| Property | GetDWordValue | rootKey, subKey, valueName | an integer | Returns the value of the specified registry location. The specified registry entry must be of type REG_DWORD. |
| Method | SetDWordValue | rootKey, subKey, valueName, value | N/A | Writes the specified REG_DWORD value to the specified registry location. Writing to HKLM or HKCR requires elevated privileges. |
| Property | HKLM | None | &H80000002 | Represents HKEY_LOCAL_MACHINE. For use with the rootKey parameter. |
| Property | HKCU | None | &H80000001 | Represents HKEY_CURRENT_USER. For use with the rootKey parameter. |
| Property | HKCR | None | &H80000000 | Represents HKEY_CLASSES_ROOT. For use with the rootKey parameter. |
| Property | GetPC | None | a string | Returns the name of the current computer. <strong> .</strong> (dot) indicates the local computer. |
| Property | GetRegValueType | rootKey, subKey, valueName | an integer | Returns a registry key value type integer. |
| Method | CreateKey | rootKey, subKey | N/A | Creates the specified subKey and all of it's parent keys, if necessary. |
| Method | EnumValues | rootKey, subKey, aNames, aTypes | N/A | Enumerates the value names and their types for the specified key. The aNames and aTypes parameters are populated with arrays of key value name strings and type integers, respectively. Wraps the StdRegProv EnumValues method, effectively fixing its <a target="_blank" href="https://groups.google.com/forum/#!topic/microsoft.public.win32.programmer.wmi/10wMqGWIfms"> lonely Default Value bug</a>, except that with HKCR and HKLM, elevated privileges are required or else aNames and aValues may be null if the default value is the only value. |
| Property | REG_SZ | None | 1 | Returns a registry value type constant. |
| Property | REG_EXPAND_SZ | None | 2 | Returns a registry value type constant. |
| Property | REG_BINARY | None | 3 | Returns a registry value type constant. |
| Property | REG_DWORD | None | 4 | Returns a registry value type constant. |
| Property | REG_MULTI_SZ | None | 7 | Returns a registry value type constant. |
| Property | REG_QWORD | None | 11 | Returns a registry value type constant. |
| Property | GetRegValueTypeString | rootKey, subKey, valueName | a string | Returns a registry key value type string suitable for use with WScript.Shell RegWrite method argument #3. That is, one of "REG_SZ", "REG_EXPAND_SZ", "REG_BINARY", or "REG_DWORD". |

## SetupHelper

 Class SetupHelper  
 Supported alternative, experimental, setup scenarios:  
 1. The original purpose was to provide custom registration of project Windows Script Component (.wsc) files and VBScript extension .dll files using HKey_Current_User instead of HKey_Local_Machine. For a brief explanation of why this approach was abandoned, see <a href=https://github.com/koswald/VBScript/blob/master/SetupPerUser.md> SetupPerUser.md</a>.  
 2. Another alternate use was for experimental registration of .wsc (Windows Script Component) files when the registration failed after the Windows 10 feature edition 20H2 update on Windows 10 Home edition. The same behavior was not observed on Windows 10 Pro edition, or after the second Windows restart.  
 If the calling script is not in the project root folder (recommended), then the ComponentFolder and ConfigFile properties must be set before calling the Setup method, specifying the paths or relative paths to the items. It is suggested that the working directory be set first, so that the other properties can be set with reference to that, without ambiguity. This can be done with the class CurrentDirectory property or by using the WScript.Shell CurrentDirectory property, or by other means.  
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Method | Init | None | N/A | Initialize certain properties, if they have not been already. |
| Method | EnsureValidRegData | arr, indexStart, indexStep, indexOffset, pattern | N/A | Ensure that the registration data to be entered into the registry is valid by raising an error when invalid data is found, which will stop the calling script, provided that the error is not supressed with an 'On Error Resume Next' statement. indexOffset: the integer to add to the current index, i, to get the array index of the partial class progid or partial interface progid. |
| Method | Char2IsUpperCase | None | N/A | If the second char of the partial progid is upper case, then the type is an interface, in which case the validation may be ignored. In this project the interface is compiled into the same .dll as the associated class. |
| Property | HKCU | None | &H80000001 | Returns a value suitable for use with the root parameter of the KeyExists property. |
| Property | HKLM | None | &H80000002 | Returns a value suitable for use with the root parameter of the KeyExists property. |

## ShellConstants

Constants for use with WScript.Shell.Run  
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Property | RunHidden | None | 0 | Window opens hidden. <br /> For use with Run method parameter #2 |
| Property | RunNormal | None | 1 | Window opens normal. <br /> For use with Run method parameter #2 |
| Property | RunMinimized | None | 2 | Window opens minimized. <br /> For use with Run method parameter #2 |
| Property | RunMaximized | None | 3 | Window opens maximized. <br /> For use with Run method parameter #2 |
| Property | Synchronous | None | True | Script execution halts and waits for the called process to exit. <br /> For use with Run method parameter #3 |
| Property | Asynchronous | None | False | Script execution proceeds without waiting for the called process to exit. <br /> For use with Run method parameter #3 |

## ShellSpecialFolders

 ShellSpecialFolders class  
  
 Adapted from <a href="https://docs.microsoft.com/en-us/windows/win32/api/shldisp/ne-shldisp-shellspecialfolderconstants"> ShellSpecialFolderConstants enumeration (shldisp.h)</a>: Specifies unique, system-independent values that identify special folders. These folders are frequently used by applications but which may not have the same name or location on any given system. For example, the system folder can be "C:\Windows" on one system and "C:\Winnt" on another.  
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Property | Path | an integer | a path | Returns the path to a special folder. The parameter is one of the ssf constants. This path is suitable for navigating in Windows Explorer. For ssfCONTROLS, ssfPRINTERS, ssfBITBUCKET, ssfDRIVES, and ssfNETWORK, the return value looks different than a typical path: for ssfDrives it is ::{20D04FE0-3AEA-1069-A2D8-08002B30309D}. |
| Property | AllConstants | None | an array | Returns an array with all of the ssf constants (integers). |
| Property | AllPaths | None | an array | Returns an array with all of the ssf paths. |
| Property | ssfDESKTOP | None | 0 | None |
| Property | ssfPROGRAMS | None | &h2 | None |
| Property | ssfCONTROLS | None | &h3 | Virtual folder that contains icons for the Control Panel applications. |
| Property | ssfPRINTERS | None | &h4 | Virtual folder that contains installed printers. |
| Property | ssfPERSONAL | None | &h5 | File system directory that serves as a common repository for a user's documents. A typical path is C:\Users&#92;<em>username</em>\Documents. |
| Property | ssfFAVORITES | None | &h6 | None |
| Property | ssfSTARTUP | None | &h7 | None |
| Property | ssfRECENT | None | &h8 | None |
| Property | ssfSENDTO | None | &h9 | None |
| Property | ssfBITBUCKET | None | &ha | According to the <a href="https://docs.microsoft.com/en-us/windows/win32/api/shldisp/ne-shldisp-shellspecialfolderconstants"> docs</a>: "Virtual folder that contains the objects in the user's Recycle Bin." |
| Property | ssfSTARTMENU | None | &hb | None |
| Property | ssfDESKTOPDIRECTORY | None | &h10 | According to the <a href="https://docs.microsoft.com/en-us/windows/win32/api/shldisp/ne-shldisp-shellspecialfolderconstants"> docs</a>: "File system directory used to physically store the file objects that are displayed on the desktop. It is not to be confused with the desktop folder itself, which is a virtual folder." A typical path is C:\Users&#92;<em>username</em>\Desktop. |
| Property | ssfDRIVES | None | &h11 | My Computer—the virtual folder that contains everything on the local computer: storage devices, printers, and Control Panel. This folder can also contain mapped network drives. |
| Property | ssfNETWORK | None | &h12 | Network Neighborhood—the virtual folder that represents the root of the network namespace hierarchy. |
| Property | ssfNETHOOD | None | &h13 | A file system folder that contains any link objects in the My Network Places virtual folder. It is not the same as ssfNETWORK, which represents the network namespace root. A typical path is C:\Users&#92;<em>username</em>\AppData\Roaming\Microsoft\Windows\Network Shortcuts. |
| Property | ssfFONTS | None | &h14 | None |
| Property | ssfTEMPLATES | None | &h15 | None |
| Property | ssfCOMMONSTARTMENU | None | &h16 | None |
| Property | ssfCOMMONPROGRAMS | None | &h17 | None |
| Property | ssfCOMMONSTARTUP | None | &h18 | None |
| Property | ssfCOMMONDESKTOPDIR | None | &h19 | None |
| Property | ssfAPPDATA | None | &h1a | None |
| Property | ssfPRINTHOOD | None | &h1b | None |
| Property | ssfLOCALAPPDATA | None | &h1c | None |
| Property | ssfALTSTARTUP | None | &h1d | None |
| Property | ssfCOMMONALTSTARTUP | None | &h1e | None |
| Property | ssfCOMMONFAVORITES | None | &h1f | None |
| Property | ssfINTERNETCACHE | None | &h20 | None |
| Property | ssfCOOKIES | None | &h21 | None |
| Property | ssfHISTORY | None | &h22 | File system directory that serves as a common repository for Internet history items. |
| Property | ssfCOMMONAPPDATA | None | &h23 | None |
| Property | ssfWINDOWS | None | &h24 | None |
| Property | ssfSYSTEM | None | &h25 | None |
| Property | ssfPROGRAMFILES | None | &h26 | None |
| Property | ssfMYPICTURES | None | &h27 | None |
| Property | ssfPROFILE | None | &h28 | None |
| Property | ssfSYSTEMx86 | None | &h29 | None |
| Property | ssfPROGRAMFILESx86 | None | &h30 | None |

## SpecialFolders

An enum and wrapper for WScript.Shell.SpecialFolders  
Usage example  
```vb
     With CreateObject( "VBScripting.Includer" ) 
         Execute .Read( "SpecialFolders" ) 
     End With 
   
     Dim sf : Set sf = New SpecialFolders 
     MsgBox sf.GetPath(sf.AllUsersDesktop) 'C:\Users\Public\Desktop 
```
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Property | GetPath | a special folder alias | a folder path | Returns the absolute path of the specified special folder. This is the default property, so the property name is optional. |
| Property | GetAliasList | None | a string | Returns a comma + space delimited list of the aliases of all the special folders. |
| Property | GetAliasArray | None | an array of strings | Returns an array of the aliases of all the special folders. |
| Property | AllUsersDesktop | None | a string | Returns a special folder alias having the exact same characters as the property name |
| Property | AllUsersStartMenu | None | a string | Returns a special folder alias having the exact same characters as the property name |
| Property | AllUsersPrograms | None | a string | Returns a special folder alias having the exact same characters as the property name |
| Property | AllUsersStartup | None | a string | Returns a special folder alias having the exact same characters as the property name |
| Property | Desktop | None | a string | Returns a special folder alias having the exact same characters as the property name |
| Property | Favorites | None | a string | Returns a special folder alias having the exact same characters as the property name |
| Property | Fonts | None | a string | Returns a special folder alias having the exact same characters as the property name |
| Property | MyDocuments | None | a string | Returns a special folder alias having the exact same characters as the property name |
| Property | NetHood | None | a string | Returns a special folder alias having the exact same characters as the property name |
| Property | PrintHood | None | a string | Returns a special folder alias having the exact same characters as the property name |
| Property | Programs | None | a string | Returns a special folder alias having the exact same characters as the property name |
| Property | Recent | None | a string | Returns a special folder alias having the exact same characters as the property name |
| Property | SendTo | None | a string | Returns a special folder alias having the exact same characters as the property name |
| Property | StartMenu | None | a string | Returns a special folder alias having the exact same characters as the property name |
| Property | Startup | None | a string | Returns a special folder alias having the exact same characters as the property name |
| Property | Templates | None | a string | Returns a special folder alias having the exact same characters as the property name |

## StartupItems

The StartupItems class provides a way to manage the programs that run automatically when Windows is started.  
  
Creating, updating, and deleting operations that affect all users must be performed with elevated privileges or else an error will occur. See comments for the Root property.  
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Property | Items | None | a collection | Returns a collection of startup item objects, each object having a Name and a Value property: The Value property is the Windows command that starts the program that is identified by the Name property. For 64-bit systems, one of four possible collecttions may be returned, depending on the values of the Root and Key properties: two of the four collections are for the current user (Root = HKCU, the default) and two are for the local machine or all users (Root = HKLM). There are separate collections for 64-bit programs (Key = StandardBranch, the default) and for 32-bit programs (Key = WowBranch). |
| Property | Item | name | an object | Returns a startup item object corresponding to the specified name. Return value depends on the values of the Root and Key properties. See comments for those properties and for the Items property. |
| Method | CreateItem | name, command | N/A | Creates a new startup item in the registry with the specified name and command. For Root = HKLM, an error will occur if privileges are not elevated. The Root and Key properties both affect where in the registry the item will be created. For 32-bit apps on a 64-bit system, use Key = WowBranch. See comments for the Items property.  |
| Method | UpdateItem | name, command | N/A | Same as the CreateItem method. |
| Method | RemoveItem | name | N/A | Same as the DeleteItem method. |
| Method | DeleteItem | name | N/A | Deletes the startup item with the specified name. For Root = HKLM, an error will occur if privileges are not elevated. The Root and Key properties both affect where in the registry the item will be deleted from. For 32-bit apps on a 64-bit system, use Key = WowBranch. See comments for the Items property.  |
| Property | Root | an integer | an integer | Together with the Key property, gets or sets the location in the registry where items will be read from, deleted from, or written to by the other properties and methods. The Root value can be specified by the property HKCU or HKLM. Root determines whether items apply to all users (HKLM) or to the current user only (HKCU). Creating, updating, and deleting operations that affect all users must be performed with elevated privileges or else an error will occur.  |
| Property | HKLM | None | an integer | Returns <code> &H80000002</code>, an integer suitable for setting the Root property. HKLM corresponds to HKEY_LOCAL_MACHINE, the system-wide all-users registry hive. |
| Property | HKCU | None | an integer | Returns <code> &H80000001</code>, an integer suitable for setting the Root property. HKCU corresponds to HKEY_CURRENT_USER, the registry hive that contains information applicable only to the current user. <strong> Note:</strong> If the current user is not a member of the Administrators group, then the current user changes when privileges are elevated. |
| Property | Key | a string | a string | Together with the Root property, gets or sets the location in the registry where items will be read from, deleted from, or written to by the other properties and methods. The Key value can be specified by the property StandardBranch (the default) or WowBranch. |
| Property | StandardBranch | None | a string | Returns the string "Software\Microsoft\Windows\CurrentVersion\Run", which partially describes a registry location that contains information about which programs start automatically on computer startup. WoWBranch and StandardBranch are the two strings suitable for setting the Key property. |
| Property | WoWBranch | None | a string | Returns the string "Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Run", which partially describes a registry location that contains information about which programs start automatically on computer startup. WoWBranch is used with 64-bit systems to store paths to 32-bit programs. StandardBranch and WoWBranch are the two strings suitable for setting the Key property. |
| Method | OpenTaskMgr | None | N/A | Opens the Task Manager at the Startup page. |

## StringFormatter

Provides string formatting functions  
  
Three instantiation examples:  
```vb
 With CreateObject( "VBScripting.Includer" ) 
      Execute .Read( "StringFormatter" ) 
      Dim fm : Set fm = New StringFormatter 
 End With 
```
or  
```vb
 With CreateObject( "VBScripting.Includer" ) 
      Dim fm : Set fm = .GetObj( "StringFormatter" ) 
 End With 
```
or  
```vb
 Dim fm : Set fm = CreateObject( "VBScripting.StringFormatter" ) 
```
Usage examples:  
```vb
 WScript.Echo fm.format(Array("MsgBox ""%s: "" & %s", "Result", -5.1)) 'MsgBox "Result: " & -5.1 
 
 WScript.Echo fm.pluralize(3, "dog") '3 dogs 
 WScript.Echo fm.pluralize(1, "dog") '1 dog 
 WScript.Echo fm.pluralize(0, "dog") '0 dogs 
 fm.SetZeroSingular 
 WScript.Echo fm.pluralize(0, "dog") '0 dog 
 WScript.Echo fm.pluralize(1, Split("person people")) '1 person 
 WScript.Echo fm.pluralize(2, Split("person people")) '2 people 
 WScript.Echo fm.pluralize(12, "egg") '12 eggs 
```
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Property | Format | array | a string | Returns a formatted string. The parameter is an array whose first element contains the pattern of the returned string. The first %s in the pattern is replaced by the next element in the array. The second %s in the pattern is replaced by the next element in the array, and so on. Variant subtypes tested OK with %s include string, integer, and single. Format is the default property for the class, so the property name is optional. If there are too many or too few %s instances, then an error will be raised. |
| Method | SetSurrogate | a string | N/A | Optional. Sets the string that the Format method will replace with the specified array element(s), %s by default. |
| Property | Pluralize | count, noun | a string | Returns a string that may or may not be pluralized, depending on the specified count. If the noun has irregular pluralization, pass in a two-element array: <code> Split("person people")</code>. Otherwise, you may pass in either a singular noun as a string, <code> red herring</code>, or else a two-element array, <code> Split("red herring &#124; red herrings", "&#124;")</code>. |
| Method | SetZeroSingular | None | N/A | Optional. Changes the default behavior of considering a count of zero to be plural. |
| Method | SetZeroPlural | None | N/A | Optional. Restores the default behavior of considering a count of zero to be plural. |

## TestingFramework

A lightweight testing framework  
Usage example  
 ```vb
     With CreateObject( "VBScripting.Includer" ) 
         Execute .Read( "VBSValidator" ) 
         Execute .Read( "TestingFramework" ) 
     End With 
     With New TestingFramework 
         .Describe "VBSValidator class" 
             Dim val : Set val = New VBSValidator 'class under test 
         .It "should return False when IsBoolean is given a string" 
             .AssertEqual val.IsBoolean( "sdfjke" ), False 
         .It "should raise an error when EnsureBoolean is given a string" 
             Dim nonBool : nonBool = "a string" 
             On Error Resume Next 
                 val.EnsureBoolean(nonBool) 
                 .AssertErrorRaised 
                 Dim errDescr : errDescr = Err.Description
                 Dim errSrc : errSrc = Err.Source 
             On Error Goto 0 
     End With 
```
  
 When a test file such as <code>spec\Configurer.spec.wsf</code> is double-clicked in Windows Explorer, the default Windows behavior is to open the script with wscript.exe, but the test requires cscript.exe, so the file is automatically restarted with cscript.exe. By default, the test opens with PowerShell in Windows Terminal, if installed. This behavior may changed by adding a "shell" key/value pair to <code>class\VBSHoster.configure</code>, overriding the default behavior.  
  
 See also <a href="#vbstestrunner"> VBSTestRunner</a> and <a href="#vbshoster"> VBSHoster</a>.  
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Method | Describe | unit description | N/A | Sets the description for the unit under test. E.g. .describe "DocGenerator class" |
| Method | It | an expectation | N/A | Sets the specification, a.k.a. spec, which is a description of some expectation to be met by the unit under test. E.g. .it "should return an integer" |
| Property | GetSpec | None | a string | Returns the specification string for the current spec. |
| Method | ShowPendingResult | None | N/A | Flushes any pending results. Generally for internal use, but may occasionally be helpful prior to an ad hoc StdOut comment, so that the comment shows up in the output in its proper place. |
| Method | AssertEqual | actual, expected | N/A | Asserts that the specified two variants, of any subtype, are equal. |
| Method | AssertErrorRaised | None | N/A | Asserts that an error should be raised by one or more of the preceeding statements. The statement(s), together with the AssertErrorRaised statement, should be wrapped with an <br /> <pre style='white-space: nowrap;'> On Error Resume Next <br /> On Error Goto 0 </pre> block. |
| Method | DeleteFile | a filespec | N/A | Deletes the specified file. Relative paths and environment variables are allowed. |
| Method | DeleteFiles | an array | N/A | Deletes the specified files. The parameter is an array of filespecs. Relative paths and environment variables are allowed. |
| Method | WriteTempMessage | a string | N/A | Writes a temporary message to the test output that can be, and should be, erased later with the EraseTempMessage method, after some behind the scenes work has been done that does not write to the console. Note: The message will not appear when the test(s) are initiated by the TestRunner class. |
| Method | EraseTempMessage | None | N/A | Erases the message written by the WriteTempMessage method. |
| Property | MessageAppeared | caption, seconds, keys | a boolean | Waits for the specified maximum time (seconds) for a dialog with the specified title-bar text (caption). If the dialog appears, acknowleges it with the specified keystrokes (keys) and returns True. If the time elapses without the dialog appearing, returns False. Note: SendKeys-related features are deprecated. |
| Method | ShowSendKeysWarning | None | N/A | Shows a SendKeys warning: a warning message to not make mouse clicks or key presses. Note: SendKeys-related features are deprecated. |
| Method | CloseSendKeysWarning | None | N/A | Closes the SendKeys warning. Note: SendKeys-related features are deprecated. |

## TextStreamer

Open a file as a text stream for reading, writing, or appending.  
Methods for use with the text stream that is returned by the Open method:  
<em> Reading methods: </em> Read, ReadLine, ReadAll <br /> <em> Writing methods: </em> Write, WriteLine, WriteBlankLines <br /> <em> Reading or Writing methods: </em> Close, Skip, SkipLine <br /> <em> Reading or writing properties: </em> AtEndOfLine, AtEndOfStream, Column, Line  
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Property | Open | None | an object | Returns a text stream object according to the specified settings (methods beginning with Set...) |
| Method | SetFile | a filespec | N/A | Specifies the file to be opened by the text streamer. Can include environment variable names. The default file is a random-named .txt file on the desktop. |
| Method | SetFolder | a folder | N/A | Specifies the folder of the file to be opened by the text streamer. Can include environment variables. Default is %UserProfile%\Desktop |
| Method | SetFileName | a file name | N/A | Specifies the file name, including extension, of the file to be opened by the text streamer. Default is a randomly named .txt file. |
| Method | SetForReading | None | N/A | Prepares the text stream to be opened for reading |
| Method | SetForWriting | None | N/A | Prepares the text stream to be opened for writing |
| Method | SetForAppending | None | N/A | Prepares the text stream to be opened for appending (default) |
| Method | SetCreateNew | None | N/A | Allows a new file to be created (default) |
| Method | SetDontCreateNew | None | N/A | Prevents a new file from being created if the file doesn't already exist |
| Method | SetAscii | None | N/A | Sets the expectation that the file will be Ascii (default) |
| Method | SetUnicode | None | N/A | Sets the expectation that the file will be Unicode |
| Method | SetSystemDefault | None | N/A | Uses Ascii or Unicode according to the system default |
| Method | View | None | N/A | Opens the file for viewing |
| Method | CloseViewer | None | N/A | Close the file viewer. From the docs: Use the Terminate method only as a last resort since some applications do not clean up properly. As a general rule, let the process run its course and end on its own. The Terminate method attempts to end a process using the WM_CLOSE message. If that does not work, it kills the process immediately without going through the normal shutdown procedure. |
| Method | SetViewer | filespec | N/A | Sets the filespec of an alternate file viewer to use with the View method.The default viewer is Notepad. |
| Method | Delete | None | N/A | Deletes the streamer file |
| Method | Run | None | N/A | Open/Run the file, assuming it has an executable file extension. |
| Property | GetFile | None | a filespec | Returns the filespec of the file that is open or set to be opened by the text streamer. Environment variables are not expanded. |
| Property | GetFileName | None | a file name | Returns the file name of the file that is open or set to be opened by the text streamer. Environment variables are not expanded. |
| Property | GetFolder | None | a folder | Returns the folder of the file that is open or set to be opened by the text streamer. Environment variables are not expanded. |
| Property | GetCreateMode | None | a boolean | Gets the current CreateMode setting. Returns one of these stream constants: bDontCreateNew or bCreateNew. |
| Property | GetStreamMode | None | an integer | Gets the current StreamMode setting. Returns one of these stream constants: iForReading, iForWriting, iForAppending |
| Property | GetStreamFormat | None | a tristate boolean | Gets the current StreamFormat setting. Returns one of these stream constants: tbAscii, tbUnicode, tbSystemDefault |

## TimeFunctions

| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Method | SetFirstDOW | an integer | N/A | Specifies the first day of the week. Parameter can be one of the VBScript constants vbSunday, vbMonday, ... |
| Property | LetDOWBeAbbreviated | a boolean | N/A | Specifies whether day-of-the-week strings should be abbreviated: Default is False. |
| Property | TwoDigit | a number | a two-char string | Returns a two-char string that may have a leading 0, given a numeric integer/string/variant of length one or two |
| Property | DOW | a date | a day of the week | Returns a day of the week string, e.g. Monday, given a VBS date |
| Property | GetFormattedDay | a date | a date string | Returns a formatted day string; e.g. 2016-09-15-Sat |
| Property | GetFormattedTime | a date | a date string | Returns a formatted 24-hr time string: e.g. 13:38:45 or 00:45:32 |

## ValidFileName

VBS function GetValidFileName and associated functions provide for modifying a string to remove characters that are not suitable for use in a Windows&reg; file name.  
Usage Example  
```vb
     With CreateObject( "VBScripting.Includer" ) 
         ExecuteGlobal .Read( "ValidFileName" ) 
     End With 
  
     MsgBox GetValidFileName("test\ing") 'test-ing 
```
  
ValidFileName.vbs provides an example of introductory comments in a script that lacks a Class statement: With DocGenerator.vbs, a line beginning with '''' (four single quotes) may be used instead of a Class statement, in order to end the introductory comments section.  
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Property | GetValidFileName | a file name candidate | a valid file name | Returns a string suitable for use as a file name: Removes <strong> \ / : * ? " < > &#124; %20 # </strong> and replaces them with a hyphen/dash (-). Limits length to maxLength value in ValidFileName.config. |
| Property | InvalidWindowsFilenameChars | None | an array | Returns an array of characters that are not allowed in Windows&reg; filenames. |
| Property | InvalidChromeFilenameStrings | None | an array | Returns an array of strings, either one of which if included in the filename of a local .html file, Chrome will not open the file. |

## VBSApp

VBSApp class  
Intended to support identical handling of class procedures by .vbs/.wsf files and .hta files.  
This can be useful when writing a class that might be used in both types of "apps".  
Four ways to instantiate  
For .vbs/.wsf scripts,  
 ```vb
  Dim app : Set app = CreateObject( "VBScripting.VBSApp" ) 
  app.Init WScript 
```
For .hta applications,  
 ```vb
  Dim app : Set app = CreateObject( "VBScripting.VBSApp" ) 
  app.Init document 
```
If the script may be used in .vbs/.wsf scripts or .hta applications  
 ```vb
  With CreateObject( "VBScripting.Includer" ) 
      Execute .Read( "VBSApp" ) 
  End With 
  Dim app : Set app = New VBSApp 
```
Alternate method for both .hta and .vbs/.wsf,  
 ```vb
  Set app = CreateObject( "VBScripting.VBSApp" ) 
  If "HTMLDocument" = TypeName(document) Then 
      app.Init document 
  Else app.Init WScript 
  End If 
```
Examples  
 ```vb
  'test.vbs "arg one" "arg two" 
  With CreateObject( "VBScripting.Includer" ) 
      Execute .Read( "VBSApp" ) 
  End With 
  Dim app : Set app = New VBSApp 
  MsgBox app.GetFileName 'test.vbs 
  MsgBox app.GetArg(1) 'arg two 
  MsgBox app.GetArgsCount '2 
  app.Quit 
```
  
 ```vb
  <!-- test.hta "arg one" "arg two" --> 
  <hta:application icon="msdt.exe"> 
      <script language="VBScript"> 
          With CreateObject( "VBScripting.Includer" ) 
              Execute .Read( "VBSApp" ) 
          End With 
          Dim app : Set app = New VBSApp 
          MsgBox app.GetFileName 'test.hta 
          MsgBox app.GetArg(1) 'arg two 
          MsgBox app.GetArgsCount '2 
          app.Quit 
      </script> 
  </hta:application> 
```
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Property | GetArgs | None | array of strings | Returns an array of command-line arguments. |
| Property | GetArgsString | None | a string | Returns the command-line arguments string. Can be used when restarting a script for example, in order to retain the original arguments. Arguments are wrapped with double quotes, if they contain spaces or if WrapAll is set to True. The return string has a leading space, by design, unless there are no arguments. |
| Property | GetArg | an integer | a string | Returns the command-line argument having the specified zero-based index. |
| Property | GetArgsCount | None | an integer | Returns the number of arguments. |
| Property | GetFullName | None | a string | Returns the filespec of the calling script or hta. |
| Property | GetFileName | None | a string | Returns the name of the calling script or hta, including the filename extension. |
| Property | GetBaseName | None | a string | Returns the name of the calling script or hta, without the filename extension. |
| Property | GetExtensionName | None | a string | Returns the filename extension of the calling script or hta. |
| Property | GetParentFolderName | None | a string | Returns the folder that contains the calling script or hta. |
| Property | GetExe | None | a string | Returns "mshta.exe" to hta files, and "wscript.exe" or "cscript.exe" to scripts, depending on the host. |
| Method | RestartWith | #1: host; #2: switch; #3: elevating | N/A | <strong> Deprecated</strong> in favor of the RestartUsing method. Restarts the script/app with the specified host (typically "wscript.exe", "cscript.exe", or "mshta.exe"), retaining the command-line arguments. Uses cmd.exe for the shell. Parameter #2 is a cmd.exe switch, "/k" or "/c". Parameter #3 is a boolean, True if restarting with elevated privileges. If userInteractive, first warns user that the User Account Control dialog will open. |
| Method | RestartUsing | #1: host; #2: exit?; #3: elevate? | N/A | Restarts the script/hta with the specified host, "wscript.exe", "cscript.exe", "mshta.exe", or a full path to one of these, retaining the command-line arguments. Uses pwsh.exe for the shell, if available, or falls back to powershell.exe. Unusual or custom paths for pwsh.exe can be specified in the file <code>.configure</code> in the project root folder. Parameter #2 is a boolean specifying whether the powershell window should exit after completion. Parameter #3 is a boolean, True if restarting with elevated privileges. If userInteractive, first warns user that the User Account Control dialog will open. If it is desired to elevate privileges, and privileges are already elevated, and the desired host is already hosting, then the script does not restart: The calling script or hta does not have to check whether privileges are elevated or explicitly call the Quit method. |
| Property | DoExit | None | True | Suitable for use with the RestartUsing method, argument #2 |
| Property | DoNotExit | None | False | Suitable for use with the RestartUsing method, argument #2 |
| Property | DoElevate | None | True | Suitable for use with the RestartUsing method, argument #3 |
| Property | DoNotElevate | None | False | Suitable for use with the RestartUsing method, argument #3 |
| Method | SetUserInteractive | boolean | N/A | Sets userInteractive value. Setting to True can be useful for debugging. Default is True. |
| Property | GetUserInteractive | None | boolean | Returns the userInteractive setting. This setting also may affect the visibility of selected console windows. |
| Method | SetVisibility | 0 (hidden) or 1 (normal) | N/A | Sets the visibility of selected command windows. SetUserInteractive also affects this setting. Default is 1. |
| Property | GetVisibility | None | 0 (hidden) or 1 (normal) | Returns the current visibility setting. SetUserInteractive also affects this setting. |
| Method | Quit | None | N/A | Gracefully closes the hta/script. |
| Method | Sleep | an integer | N/A | Pauses execution of the script or .hta for the specified number of milliseconds. |
| Property | WScriptHost | None | "wscript.exe" | Can be used as an argument for the method RestartWith. |
| Property | CScriptHost | None | "cscript.exe" | Can be used as an argument for the method RestartWith. |
| Property | GetHost | None | "wscript.exe" or "cscript.exe" or "mshta.exe" | Returns the current host. Can be used as an argument for the method RestartWith. |

## VBSArguments

Functions related to VBScript command-line arguments.  
  
Not suitable for use with .hta files. For .hta files, use VBSApp instead.  
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Property | GetArgumentsString | None | a string containing all command-line arguments | For use when restarting a script, in order to retain the original arguments. Each argument is wrapped wih quotes, which are stripped off as they are read back in. The return string has a leading space, by design, unless there are no arguments |

## VBSArrays

| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Property | Uniques | an array | an array | Returns an array with no duplicate items, given an array that may have some. |
| Property | RemoveFirstElement | an array of strings | an array of strings | Returns a array without the first element of the specified array. |
| Property | CollectionToArray | a collection of strings | array of strings | Can be used to convert the WScript.Arguments object to an array, for example. |

## VBSClipboard

Clipboard procedures  
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Method | SetClipboardText | a string | N/A | Copies the specified string to the clipboard. Uses clip.exe, which shipped with Windows&reg; Vista / Server 2003 through Windows 10. |
| Property | GetClipboardText | None | a string | Returns text from the clipboard |

## VBSEnvironment

| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Property | Expand | a string | a string | Expands environment variable(s); e.g. convert %UserProfile% to C:\Users\user42 |
| Property | Collapse | a string | a string | Collapses a string that may contain one or more substrings that can be shortened to an environment variable. |
| Method | CreateUserVar | varName, varValue | N/A | Create or set a user environment variable |
| Method | SetUserVar | varName, varValue | N/A | Set or create a user environment variable |
| Property | GetUserVar | a variable name | the variable value | Returns the value of the specified user environment variable |
| Method | RemoveUserVar | varName | N/A | Removes a user environment variable |
| Method | CreateProcessVar | varName, varValue | N/A | Create a process variable |
| Method | SetProcessVar | varName, varValue | N/A | Sets or creates a process environment variable |
| Property | GetProcessVar | varName | the variable value | Returns the value of the specified environment variable |
| Method | RemoveProcessVar | varName | N/A | Removes the specified process environment variable |
| Property | GetDefaults | None | an array | Returns an array of common environment variables pre-installed with some versions of Windows&reg;. Not exhaustive. |

## VBSEventLogger

Logs messages to the Application event log.  
  
Wraps the LogEvent method of the WScript.Shell object.  
  
To see a log entry, type EventVwr at the command prompt to open the Event Viewer, expand Windows Logs, and select Application. The log Source will be WSH. Or you can use the CreateCustomView method to create an entry in the Event Viewer's Custom Views section.  
  
Usage example:  
 ```vb
  With CreateObject( "VBScripting.Includer" ) 
      Execute .Read( "VBSEventLogger" ) 
  End With 
   
  Dim logger : Set logger = New VBSEventLogger 
  logger.log logger.INFORMATION, "message 1" 
  logger logger.INFORMATION, "message 2" 
  logger 4, "message 3" 
  logger 1, "error message" 
   
  logger.CreateCustomView 'create a custom view in the Event Viewer 
  logger.OpenViewer 'open EventVwr.msc 
```
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Method | Log | eventType, message | N/A | Adds an event entry to a log file with the specified message. This is the default method, so the method name is optional. |
| Method | CreateCustomView | None | N/A | Creates a Custom View in the Event Viewer, eventvwr.msc, named WSH Logs. The User Account Control dialog will open, in order to confirm elevation of privileges. Based on VBSEventLoggerCustomView.xml. |
| Method | OpenViewer | None | N/A | Opens the Windows&reg; Event Viewer, eventvwr.msc |
| Property | SUCCESS | None | 0 | Returns a value for use as an "eventType" parameter |
| Property | ERROR | None | 1 | Returns a value for use as an "eventType" parameter |
| Property | WARNING | None | 2 | Returns a value for use as an "eventType" parameter |
| Property | INFORMATION | None | 4 | Returns a value for use as an "eventType" parameter |
| Property | AUDIT_SUCCESS | None | 8 | Returns a value for use as an "eventType" parameter |
| Property | AUDIT_FAILURE | None | 16 | Returns a value for use as an "eventType" parameter |
| Method | OpenConfigFolder | None | N/A | Opens the Event Viewer configuration folder, by default "%ProgramData%\Microsoft\Event Viewer". The Views subfolder contains the .xml files defining the custom views. |
| Method | OpenLogFolder | None | N/A | Opens the folder with the .evtx files that contain the event logs, by default "%SystemRoot%\System32\Winevt\Logs". Application.evtx holds the WSH data. |

## VBSExtracter

For extracting a string from a text file, given a regular expression  
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Method | SetPattern | a regex pattern | N/A | Required. Specifies the text to be extracted. Non-regex expressions containing any of the regex special characters <strong>(  )  .  $  +  [  ?  \  ^  {  &#124;</strong> must preceed the special character with a <strong>&#092;</strong> |
| Method | SetFile | filespec | N/A | Required. Specifies the file to extract text from. |
| Method | SetIgnoreCase | a boolean | N/A | Set whether to ignore case when matching text. Default is False. |
| Property | Extract | None | a string | Returns the first string that matches the specified regex pattern. Returns an empty string if there is no match. Before calling this method, you must specify the file and the pattern: see SetPattern and SetFile. |
| Property | Extract0 | None | a string | Deprecated for not spanning multiple lines. Formerly named Extract. Returns the string that matches the specified regex pattern. Returns an empty string if there is no match. Before calling this method, you must specify the file and the pattern: see SetPattern and SetFile. |

## VBSFileSystem

General utility functions  
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Property | SBaseName | None | a file name, no extension | Returns the name of the calling script, without the file name extension. |
| Property | SName | None | a file name | Returns the name of the calling script, including file name extension |
| Property | SFullName | None | a filespec | Returns the filespec of the calling script |
| Property | SFolderName | None | a folder | Returns the parent folder of the calling script. |
| Property | MakeFolder | a path | a boolean | Create a folder, and if necessary create also its parent, grandparent, etc. Returns False if the folder could not be created. |
| Property | Parent | a folder, file, or registry key | the item's parent | Returns the parent of the folder or file or registry key, or removes a trailing backslash. The parent need not exist. |
| Method | SetReferencePath | a path | N/A | Optional. Specifies the base path from which relative paths should be referenced. By default, the reference path is the parent folder of the calling script. See also Resolve and ResolveTo. |
| Property | Resolve | a relative path | a resolved path | Resolves a relative path (e.g. "../lib/WMI.vbs"), to an absolute path (e.g. "C:\Users\user42\lib\WMI.vbs"). The relative path is by default relative to the parent folder of the calling script, but this behavior can be changed with SetReferencePath. See also property ResolveTo. |
| Property | ResolveTo | relativePath, absolutePath | a resolved path | Resolves the specified relative path, e.g. "../lib/WMI.vbs", relative to the specified absolute path, and returns the resolved absolute path, e.g. "C:\Users\user42\lib\WMI.vbs". Environment variables are allowed. |
| Property | Expand | a string | an expanded string | Given a string which may contain environment variables, returns the string with environment variable(s) expanded. E.g. %WinDir% => C:\Windows |
| Method | Elevate | command, arguments, folder | N/A | Runs the specified command with elevated privileges, with the specified arguments and working folder |
| Property | FoldersAreTheSame | folder1, folder2 | a boolean | Determines whether the two specified folders are the same. If so, returns True. |
| Method | DeleteFile | filespec | N/A | Deletes the specified file. |
| Method | SetForceDelete | boolean | N/A | Controls the behavior of the DeleteFile method: Specify True to force a file deletion even when the file is read-only. Optional. Default is False. |

## VBSHoster

 Manage which script host is hosting the currently running script: cscript.exe or wscript.exe.  
  
 Not suitable for .hta scripts. For .hta scripts, use the VBSApp class.  
  
 If Windows Terminal is installed, a suggested setting in %LocalAppData%\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json: <code>"windowingBehavior": "useAnyExisting"</code> or <code>"windowingBehavior": "useExisting"</code>. The same setting in the Windows Terminal GUI: Settings &#124; Startup &#124; New instance behavior &#124; Attach to the most recently used window (or Attach to the most recently used window on this desktop). This applies to the RestartWith method's default behavior. The RestartWith method is used by the TestingFramework class when a test file is double-clicked in Windows Explorer.  
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Method | EnsureCScriptHost | None | N/A | Restart the script hosted with cscript.exe if it isn't already hosted with cscript.exe. |
| Method | SetSwitch | /k or /c | N/A | Optional. Specifies a switch (command-line argument) for %ComSpec% for use with the EnsureCScriptHost method: controls whether the command window, if newly created, remains open (/k). Useful for troubleshooting, in order to be able to read error messages. Unnecessary if starting the script from a console window, because /c is the default. If pwsh or powershell (or wt pwsh, etc.) is the Shell, then the equivalent string is substituted. |
| Method | SetDefaultHostWScript | None | N/A | Sets wscript.exe to be the default script host. If privileges are not already elevated, then the User Account Control dialog will open for permission to elevate privileges. |
| Method | SetDefaultHostCScript | None | N/A | Sets cscript.exe to be the default script host. If privileges are not already elevated, then the User Account Control dialog will open for permission to elevate privileges. |
| Property | GetDefaultHost | None | a string | Returns "wscript.exe" or "cscript.exe", according to which .exe opens .vbs files by default. |
| Method | RestartWith | a string: the host .exe | N/A | Restarts the .vbs or .wsf script with the specified host, "cscript.exe" or "wscript.exe". By default, Windows Terminal will be used, if available. Also by default, pwsh.exe (PowerShell) will be used if available. A custom or unusual pwsh.exe install path can be specified if necessary in the file <code>.configure</code> in the project root folder. Use <code> class/VBSHoster.configure</code> to specify another shell configuration. <br /> Examples:<br /><code>shell, cmd</code><br /><code>shell, powershell</code><br /><code>shell, pwsh</code><br /><code>shell, wt cmd</code><br /><code>shell, wt pwsh</code><br /><code>shell, wt "%ProgramFilesX86%\PowerShell\7\pwsh.exe"</code><br /><code>shell, %ProgramFilesX86%\PowerShell\7\pwsh.exe</code><br />This setting can be overridden by the Shell property. See also the RestartUsing method of the <a href="#vbsapp"> VBSApp class</a>. |
| Property | Shell | a string | a string | Gets or sets the shell used when restarting a script (see the RestartWith method). Examples: cmd, powershell, pwsh, wt pwsh. Overrides the shell read from <code>VBSHoster.configure</code>. |

## VBSLogger

A lightweight VBScript logger  
Instantiation  
```vb
     With CreateObject( "VBScripting.Includer" ) 
         Execute .Read( "VBSLogger" ) 
     End With 
     Dim log : Set log = New VBSLogger 
```
  
Usage method one. This method has the advantage that the log doesn't remain open, allowing other scripts to write to the log.  
 ```vb
     log "test one" 
```
Usage method two. This method has the advantage that the name of the calling script is not written on each line of the log.  
 ```vb
     log.Open 
     log.Write "test two" 
     log.Close 
```
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Method | Log | a string | N/A | Opens the log file, writes the specified string, then closes the log file. This is the default method for the VBSLogger class. |
| Method | SetLogFolder | a folder path | N/A | Optional. Customize the log folder. The folder will be created if it does not exist. Environment variables are allowed. See GetDefaultLogFolder. |
| Method | Open | None | N/A | Opens the log file for writing. The log file is opened and remains open for writing. While it is open, other processes/scripts will be unable to write to it. |
| Method | Write | a string | N/A | Writes the specified string to the log file. |
| Method | Close | None | N/A | Closes the log file text stream, enabling other process to write to it. |
| Method | View | None | N/A | Opens the log file for viewing. Notepad is the default editor. See SetViewer. |
| Method | SetViewer | a filespec | N/A | Optional. Customize the program that the View method uses to view log files. Default: Notepad. |
| Method | ViewFolder | None | N/A | Open the log folder |
| Property | WordPad | None | a filespec | Can be used as the argument for the SetViewer method in order to open files with WordPad when the View method is called. |
| Property | GetDefaultLogFolder | None | a folder | Retrieves the default log folder, %AppData%\VBScripting\logs |
| Property | GetLogFilePath | None | a filespec | Retreives the filespec for the log file, with environment variables expanded. Default: &lt;GetDefaultLogFolder&gt;\YYYY-MM-DD-DayOfWeek.txt |

## VBSPower

Power functions: shutdown, restart, logoff, sleep, and hibernate.  
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Property | Shutdown | None | a boolean | Shuts down the computer. Returns True if the operation completes with no errors. |
| Property | Restart | None | a boolean | Restarts the computer. Returns True if the operation completes with no errors. |
| Property | Logoff | None | a boolean | Logs off the computer. Returns True if the operation completes with no errors. |
| Method | Sleep | None | N/A | Puts the computer to sleep. Requires <a target="_blank" href="https://docs.microsoft.com/en-us/sysinternals/downloads/psshutdown"> PsTools</a> download and PsShutdown.exe to be located somewhere on your %Path%. Recovery from sleep is faster than from hibernation, but uses more power. |
| Method | Hibernate | None | N/A | Puts the computer into hibernation. Will not work if hibernate is disabled in the Control Panel, in which case the EnableHibernation method may be used to reenable hibernation. Hibernate is more power-efficient than sleep, but recovery is slower. If the computer wakes after pressing a key or moving the mouse, then it was sleeping, not in hibernation. Recovery from hibernation typically requires pressing the power button. |
| Method | EnableHibernation | None | N/A | Enables hibernation. The User Account Control dialog will open to request elevated privileges. |
| Method | DisableHibernation | None | N/A | Disables hibernation. The User Account Control dialog will open to request elevated privileges. |
| Method | SetForce | a boolean | N/A | Optional. Setting this to True forces the Shutdown or Restart, discarding unsaved work. Default is False. Logoff always forces apps to close. Windows 10 may force the specified action regardless of this setting. |
| Method | SetDebug | a boolean | N/A | Used for testing. True prevents the computer from actually shutting down, etc., during testing. Default is False. |

## VBSStopwatch

A timer  
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Property | Split | None | a rounded number (Single) | Returns the seconds elapsed since object instantiation or since calling the Reset method. Split is the default Property. |
| Method | SetPrecision | 0, 1, or 2 | N/A | Sets the number of decimal places to round the Split function return value. Default is 2. |
| Property | GetPrecision | None | 0, 1, or 2 | Returns the current precision. |
| Method | Reset | None | N/A | Sets the timer to zero. |

## VBSTestRunner

Run a test or group of tests  
Usage example  
 ```vb
    'test-launcher.vbs 
    'run this file from a console window; e.g. cscript //nologo test-launcher.vbs 
   
     With CreateObject( "VBScripting.Includer" ) 
         Execute .Read( "VBSTestRunner" ) 
     End With 
   
     With New VBSTestRunner 
         .SetSpecFolder "../spec" 'location of test files relative to test-launcher.vbs 
         .Run 
     End With 
```
  
See also <a href=#testingframework> TestingFramework</a>.  
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Method | Run | None | N/A | Initiate the specified tests |
| Method | SetSpecFolder | a folder | N/A | Optional. Specifies the folder containing the test files. Can be a relative path, relative to the calling script. Default is the parent folder of the calling script. |
| Method | SetSpecPattern | wildcard(s) | N/A | Optional. Specifies which file types to run. Default is *.spec.vbs. Standard wildcard notation with &#124; delimiter. |
| Method | SetSpecFile | a file | N/A | Optional. Specifies a single file to test. Include the filename extension. E.g. SomeClass.spec.vbs. A relative path is OK, relative to the spec folder. If no spec file is specified, all test files matching the specified pattern will be run. See SetSpecPattern. |
| Method | SetSearchSubfolders | a boolean | N/A | Optional. Specifies whether to search subfolders for test files. True or False. Default is False. |
| Method | SetPrecision | 0, 1, or 2 | N/A | Optional. Sets the number of decimal places for reporting the elapsed time. Default is 2. |
| Method | SetRunCount | an integer | N/A | Optional. Sets the number of times to run the test(s). Default is 1. |

## VBSTroubleshooter

| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Method | LogAscii | a string | N/A | Write to the log the Ascii codes for each character in the specified string. |

## VBSValidator

A working example of how validation can be accomplished.  
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Property | GetClassName | None | the class name | Returns                           "VBSValidator". Useful for verifying Err.Source in a test. |
| Property | IsBoolean | a boolean candidate | a boolean | Returns True if the parameter is a boolean subtype; False if not. |
| Property | EnsureBoolean | a boolean candidate | boolean | Raises an error if the parameter is not a boolean. Unless an error is raised, returns the same value passed to it. |
| Property | IsInteger | an integer candidate | a boolean | Returns True if the parameter is an integer subtype; False if not. |
| Property | EnsureInteger | an integer candidate | integer | Raises an error if the parameter is not an integer. Unless an error is raised, returns the same value passed to it. |
| Property | ErrDescrBool | None | a string | " is not a boolean." Useful for verifying Err.Description in a test. |
| Property | ErrDescrInt | None | a string | " is not an integer." Useful for verifying Err.Description in a test. |

## WindowsUpdatesPauser

Pause Windows Updates to get more bandwidth. Don't forget to resume.  
For configuration settings, see the .config file in %AppData%\VBScripting that has the same base name as the calling script/hta.  
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Method | PauseUpdates | None | N/A | Pauses Windows Updates. |
| Method | ResumeUpdates | None | N/A | Resumes Windows Updates. |
| Property | GetStatus | None | a string | Returns Metered or Unmetered. If Metered, then Windows Updates has paused to save money, incidentally not soaking up so much bandwidth. If TypeName(GetStatus) = "Empty", then the status could not be determined, possibly due to a bad network name (internal name: profileName). |
| Property | GetAppName | None | a string | Returns the base name of the calling script |
| Property | GetProfileName | None | a string | Returns the name of the network. The name is set by editing the .config file in %AppData%\VBScripting that has the same base name as the calling script/hta. |
| Property | GetServiceType | None | a string | Returns the service type |
| Method | OpenConfigFile | None | N/A | Opens the .config file for editing. |

## WMIUtility

Examples of the Windows Management Instrumentation object.  
  
 See <a href=https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/computer-system-hardware-classes > Computer System Hardware Classes</a>.  
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Property | TerminateProcessById | process id | a boolean | Terminates any Windows&reg; process with the specified id. Returns True if the process was found, False if not. |
| Property | TerminateProcessByIdAndName | id, name | a boolean | Terminates a process with the specified id and name. Returns True if the process was found, False if not. |
| Method | TerminateProcessByIdAndNameDelayed | id, name, milliseconds | N/A | Terminates a process with the specified id (integer), name (string, e.g. notepad.exe), and delay (integer: milliseconds), asynchronously. |
| Property | GetProcessIDsByName | a process name | a boolean | Returns an array of the process ids of all processes that have the specified name. The process name is what would appear in the Task Manager's Details tab. <br /> E.g. <code> notepad.exe</code>. |
| Property | GetProcessesWithNamesLike | a string like jav% | an array of process names | None |
| Property | IsRunning | a process name | a boolean | Returns a boolean indicating whether at least one instance of the specified process is running. <br /> E.g. <code> wmi.IsRunning( "notepad.exe" ) 'True or False</code>. |
| Property | partitions | None | a collection | Returns a collection of <a href=https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-diskpartition> Win32_DiskPartition</a> objects, each having these properties, among others: Caption, Name, DiskIndex, Index, PrimaryPartition, Bootable, BootPartition, Description, Type, Size, StartingOffset, BlockSize, DeviceID, Access, Availability, ErrorMethodology, HiddenSectors, Purpose, Status. |
| Property | disks | None | a collection | Returns a collection of <a href=https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-logicaldisk> Win32_LogicalDisk</a> objects, each having these properties, among others: FileSystem, DeviceID. |
| Property | cpu | None | an object | Returns a <a href=https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-processor> Win32_Processor</a> object that has these properties, among others: Architecture, Description. |
| Property | CPUs | None | a collection | Returns a collection of <a href=https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-processor> Win32_Processor</a> objects, each of which has these properties, among others: Architecture, Description |
| Property | os | None | an object | Returns a <a href=https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-operatingsystem> Win32_OperatingSystem</a> object having these properties, among others: Name, Version, Manufacturer, WindowsDirectory, Locale, FreePhysicalMemory, TotalVirtualMemorySize, FreeVirtualMemory, SizeStoredInPagingFiles. |
| Property | pc | None | an object | Returns a <a href=https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-computersystem> Win32_ComputerSystem</a> object which has these properties, among others: Name, Manufacturer, Model, CurrentTimeZone, TotalPhysicalMemory. |
| Property | Bios | None | an object | Returns a <a href=https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-bios> Win32_BIOS</a> object which has a Version property, among others. |
| Property | Battery | None | an object | Returns a <a target="_blank" href="https://docs.microsoft.com/en-us/windows/desktop/CIMWin32Prov/win32-battery"> Win32_Battery</a> object, which has these properties, among others: BatteryStatus, EstimatedChargeRemaining. |

## WoWChecker

Provides an object whose default property, isWoW, returns a boolean indicating whether the calling script was itself called by a SysWoW64 (32-bit) .exe file. WoW64 stands for Windows 32-bit on Windows 64-bit.  
  
How it works: .exe files in %SystemRoot%\System32 and %SystemRoot%\SysWoW64 are compared by size or checksum. If the files are the same, then the calling script is assumed to be running in a 32-bit process.  
  
Usage examples  
```vb
 MsgBox New WoWChecker.BySize.isWoW 
 MsgBox New WoWChecker.isWoW 
 With New WoWChecker : .BySize : MsgBox .isWoW : End With 
 With New WoWChecker.BySize : MsgBox .isWoW : End With 
 MsgBox New WoWChecker 
```
  
| Member type | Name | Parameter | Returns | Comment |
| :---------- | :--- | :-------- | :------ | :------ |
| Property | OSIs64Bit | None | a boolean | Returns a boolean that indicates whether the Windows OS is 64-bit. |
| Property | isWoW | None | a boolean | Returns a boolean that indicates whether the calling script was itself called by a SysWoW64 (32-bit) .exe file. This is the class default property. |
| Property | isSysWoW64 | None | a boolean | Wraps isWoW: Same as calling isWoW. |
| Property | isSystem32 | None | a boolean | Returns the opposite of isSysWoW64 |
| Property | BySize | None | an object self reference | Optional. Specifies that the .exe files will be compared by size. BySize will not distinguish between the 32- and 64-bit .exe files if they are the same size, which is unlikely but possible. ByCheckSum is therefore more reliable. |
| Property | ByCheckSum | None | an object self reference | Selected by default. Specifies that the .exe files will be compared by checksum. ByCheckSum uses CertUtil, which ships with Windows&reg; 7 through 10, and can be manually installed on older versions. |
| Property | File | None | a string | Optional. Sets or gets the name of the file used in comparisons. A file by this name must be found in both %SystemRoot%\System32 and %SystemRoot%\SysWoW64. The default is <code> cmd.exe</code>. |
