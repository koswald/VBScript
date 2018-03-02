# VBScript Classes

### Contents

[Chooser](#chooser)  
[CommandParser](#commandparser)  
[DocGenerator](#docgenerator)  
[DocGeneratorCS](#docgeneratorcs)  
[EncodingAnalyzer](#encodinganalyzer)  
[EscapeMd](#escapemd)  
[GUIDGenerator](#guidgenerator)  
[HTAApp](#htaapp)  
[Includer](#includer)  
[KeyDeleter](#keydeleter)  
[MathConstants](#mathconstants)  
[PrivilegeChecker](#privilegechecker)  
[RegExFunctions](#regexfunctions)  
[RegistryUtility](#registryutility)  
[ShellConstants](#shellconstants)  
[SpecialFolders](#specialfolders)  
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
[VBSTestRunner](#vbstestrunner)  
[VBSTimer](#vbstimer)  
[VBSTroubleshooter](#vbstroubleshooter)  
[VBSValidator](#vbsvalidator)  
[WindowsUpdatesPauser](#windowsupdatespauser)  
[WMIUtility](#wmiutility)  
[WoWChecker](#wowchecker)  


## Chooser
Get a folder or file chosen by the user  
Usage example  
  
```vb
 With CreateObject("VBScripting.Includer") 
     Execute .read("Chooser")
 End With 

 Dim choose : Set choose = New Chooser 
 MsgBox choose.folder 
 MsgBox choose.file 
```
  
Browse for file <a href="http://stackoverflow.com/questions/21559775/vbscript-to-open-a-dialog-to-select-a-filepath"> reference</a>.  
Browse for folder <a href="http://ss64.com/vb/browseforfolder.html"> reference</a>.  
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Property|File|None|a file path|Opens a Choose File dialog and returns the path of a file chosen by the user. Returns an empty string if no folder was selected. Note: The title bar text will say Choose File to Upload.|
|Property|Folder|None|a folder path|Opens a Browse For Folder dialog and returns the path of a folder chosen by the user. Returns an empty string if no folder was selected.|
|Property|FolderTitle|None|the folder title|Opens a Browse For Folder dialog and returns the title of a folder chosen by the user. The title for a normal folder is just the folder name. For a special folder like %UserProfile%, it may be something entirely different. Returns an empty string if no folder was selected.|
|Property|FolderObject|None|an object|Opens a Browse For Folder dialog and returns a Shell.Application BrowseForFolder object for a folder chosen by the user. This object has methods Title and Self.Path, corresponding to this class's FolderTitle and FolderPath, respectively. This method is recommended for when you need both the FolderTitle and FolderPath but only want the user to have to choose once. If no folder was selected, then TypeName(folderObj) = "Nothing" is True.|
|Method|SetWindowTitle|a string|N/A|Sets the title of the Browse For Folder window: i.e. the text below the titlebar.|
|Method|SetWindowOptions|a hex value|N/A|Sets the behavior or behaviors for the Browse For Folder window. The parameter is one or more of the BIF_ constants:  e.g. obj.BIF_EDITBOX + obj.BIF_NONEWFOLDER.|
|Method|AddWindowOptions|a hex value|N/A|Adds a behavior or behaviors to the Browse For Folder window. The parameter is one or more of the BIF_ constants:  e.g. obj.BIF_EDITBOX + obj.BIF_NONEWFOLDER.|
|Property|BIF_RETURNONLYFSDIRS|None|&H0001|None|
|Property|BIF_DONTGOBELOWDOMAIN|None|&H0002|None|
|Property|BIF_STATUSTEXT|None|&H0004|None|
|Property|BIF_RETURNFSANCESTORS|None|&H0008|None|
|Property|BIF_NONEWFOLDER|None|&H0200|None|
|Property|BIF_BROWSEFORCOMPUTER|None|&H1000|None|
|Property|BIF_BROWSEFORPRINTER|None|&H2000|None|
|Method|SetRootPath|a folder path|N/A|Sets the root folder that the Browse For Folder window will allow browsing. Environment variables are allowed. See also the UnwiselyEnableSendKeys method.|
|Method|UnwiselyEnableSendKeys|None|N/A|Optional. Not recommended. Enables sending keystrokes to the Choose File to Upload dialog in order to open at the RootFolder. There is a risk whenever using the WScript.Shell SendKeys method that keystrokes will be sent to the wrong window.|
|Method|WiselyDisableSendKeys|None|N/A|Default setting. Disables SendKeys. The Choose File to Upload dialog will open to the last place a file was selected, regardless of the RootFolder setting.|
|Method|SetPatience|time in seconds|N/A|Sets the maximum time in seconds that the File method waits for the Choose File to Upload dialog to appear before abandoning attempts to open the dialog at the folder specified by RootFolder. Applies only when SendKeys is enabled. Default is 5 (seconds).|
|Property|DialogHasOpened|a string or an object|a boolean|Waits for the specified dialog to appear, then returns False if the specified doesn't appear within the time specified by SetPatience, by default 5 (seconds). Parameter is either a string to match with the title bar text, as when browsing for a file, or else a WshScriptExec object, as when browsing for a folder. Used internally and by the unit test.|
|Method|SetBFFileTimeout|an integer|N/A|Sets the time in seconds after which the Browse For File (Choose File to Upload) dialog will be terminated if a file has not been chosen. A timeout of 0 will allow the dialog to remain open indefinitely. Intended to allow improved testing reliability. Default is 0.|
|Method|SetMaxExecLifetime|WShellExec object, exe, milliseconds|N/A|Terminates a WShellExec process (the Browse for File window for example) after the specified time in milliseconds. Timeout of 0 prevents termination. An example of the exe: "mshta.exe".|

## CommandParser
Command Parser  
  
Runs a specified command and searches the output for a phrase  
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Method|SetCommand|newCmd|N/A|Sets the command to run whose output will be searched. Required before calling GetResult.|
|Method|SetSearchPhrase|newSearchPhrase|N/A|Sets a phase to search for in the command's output. Required before calling GetResult.|
|Property|GetResult|None|a boolean|Runs the sepecified command and returns True if the specified phrase is found in the command output.|
|Method|SetStartPhrase|newStartPhrase|N/A|Sets a unique phrase to identify the output line after which the search begins. Optional. By defualt the output is searched from the beginning.|
|Method|SetStopPhrase|newStopPhrase|N/A|Sets a unique phrase to identify the line that follows the last line of the search. Optional. By defualt, the output is searched to the end.|

## DocGenerator
Generate html and markdown documentation for VBScript code based on well-formed comments.  
Usage Example  
```vb
 With CreateObject("VBScripting.Includer")
     Execute .read("DocGenerator")
 End With
 With New DocGenerator
     .SetTitle "VBScript Utility Classes Documentation"
     .SetDocName "TheDocs.html"
     .SetFilesToDocument "*.vbs | *.wsf | *.wsc"
     .SetScriptFolder = "..\..\class"
     .SetDocFolder = "..\.."
     .Generate
     .View
 End With
```
  
<h5> Example of well-formed comments before a Sub statement </h5>  
 Note: A remark is required for Methods (Subs).  
  
```vb
'Method: SubName
'Parameters: varName, varType
'Remark: Details about the parameters.
```
<h5> Example of well-formed comments before a Property or Function statement </h5>  
 Note: A Returns (or Return or Returns: or Return:) is required with a Property or Function.  
  
```vb
'Property: PropertyName
'Returns: a string
'Remark: A remark is not required for a Property or Function.
```
<h5> Notes for the comment syntax at the beginning of a script </h5>  
Use a single quote (') for general comments <br />  
- lines without html will be wrapped with p tags <br />  
- lines with html will not be wrapped with p tags <br />  
- use a single quote by itself for an empty line <br />  
- Wrap VBScript code with <code>pre</code> tags, separating multiple lines with &lt;br /&gt;. <br />  
- Wrap other code with <code>code</code> tags, separating multiple lines with &lt;br /&gt;. <br />  
  
Use three single quotes for remarks that should not appear in the documentation <br />  
  
Use four single quotes (''''), if the script doesn't contain a class statement, to separate the general comments at the beginning of the file from the rest of the file.  
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Method|SetScriptFolder|a folder|N/A|Required. Must be set before calling the Generate method. Sets the folder containing the scripts to include in the generated documentation. Environment variables OK. Relative paths OK.|
|Method|SetDocFolder|a folder|N/A|Required. Must be set before calling the Generate method. Sets the folder of the documentation file. Environment variables OK. Relative paths OK.|
|Method|SetDocName|a filename|N/A|Required. Must be set before calling the Generate method. Specifies the name of the documentation file, including the filename extension (.html suggested).|
|Method|SetTitle|a string|N/A|Required. Must be set before calling the Generate method. Sets the title for the documentation.|
|Method|SetFilesToDocument|wildcard(s)|N/A|Optional. Specifies which files to document: default is <strong> *.vbs </strong>. Separate multiple wildcards with " | ".|
|Method|Generate|None|N/A|Generate comment-based documentation for the scripts in the specified folder.|
|Method|View|None|N/A|Open the documentation file for viewing|
|Property|Colorize|-|-|Gets or sets whether a &lt;pre&gt; code block in the markdown (.md) document (assumed to be VBScript) is colorized. If False (experimental, with GFM), the code lines will not wrap. Default is True|

## DocGeneratorCS
 DocGeneratorCS class  
  
 Generates html and markdown documentation for C# code from compiler-generated xml files based on three-slash (///) code comments.<br />  
 Four base tags are supported: summary, parameters, returns, and remarks.<br />  
 Within these tags, html tags are supported. While not all html tags are supported by markdown, they should at least be tolerated, subject to the Note below.  
 Note: Html tags may result in malformed markdown table rows when there is whitespace between adjacent tags.  
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Property|XmlFolder|-|-|Required. Gets or sets the folder containing the .xml files autogenerated by the C# compiler. Relative paths and environment variables are supported.|
|Property|OutputFile|-|-|Required. Gets or sets the path and base name of the output files, not including  the .html and .md filename extensions. Older versions, if any, will be overwritten. Relative paths and environment variables are supported.|
|Method|Generate|None|N/A|Generates html and markdown code documentation. Requires .xml files to have been generated by the C# compiler.|
|Method|ViewHtml|None|N/A|Opens the html document with the default viewer.|
|Method|ViewMarkdown|None|N/A|Opens the markdown document with the default viewer.|

## EncodingAnalyzer
Provides various properties to analyze a file's encoding  
Usage example  
```vb
With CreateObject("VBScripting.Includer")
    Execute .read("EncodingAnalyzer")
End With
 
With New EncodingAnalyzer.SetFile(WScript.Arguments(0))
    MsgBox "isUTF16LE: " & .isUTF16LE
End With
```
  
Stackoverflow references: <a href="http://stackoverflow.com/questions/3825390/effective-way-to-find-any-files-encoding"> 1</a>, <a href="http://stackoverflow.com/questions/1410334/filesystemobject-reading-unicode-files"> 2</a>.  
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Property|SetFile|a filespec|an object self reference|Required. Specifies the file whose encoding is to be determined. Relative paths are permitted, relative to the current directory.|
|Property|isUTF16LE|None|a boolean|Returns a boolean indicating whether the file specified by SetFile is Unicode Little Endian, <strong> aka Unicode</strong>.|
|Property|isUTF16BE|None|a boolean|Returns a boolean indicating whether the file specified by SetFile is Unicode Big Endian.|
|Property|isUTF7|None|a boolean|Returns a boolean indicating whether the file specified by SetFile is UTF7.|
|Property|isUTF8|None|a boolean|Returns a boolean indicating whether the file specified by SetFile is UTF8.|
|Property|isUTF32|None|a boolean|Returns a boolean indicating whether the file specified by SetFile is UTF32.|
|Property|isAscii|None|a boolean|Returns a boolean indicating whether the file specified by SetFile is Ascii.|
|Property|GetType|None|a string|Returns one of the following strings according the format of the file set by SetFile: Ascii, UTF16LE, UTF16BE, UTF7, UTF8, UTF32.|
|Property|GetCurrentDirectory|None|a folder|Returns the current directory|
|Method|SetCurrentDirectory|a folder|N/A|Sets the current directory.|
|Property|GetByte|BOM byte number|an integer|Returns the Ascii value, 0 to 255, of the byte specified. The parameter must be an integer: one of 0, 1, 2, or 3. These represent the first four bytes in the file, the Byte Order Mark (BOM).|

## EscapeMd
EscapeMd.vbs  
Escapes markdown special characters.  
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Property|EscapeMd|unescaped string|escaped string|Returns a string with Markdown special characters escaped.|

## GUIDGenerator
Generate a unique GUID  
Usage example  
```vb
 With CreateObject("VBScripting.Includer")
     Execute .read("GUIDGenerator")
 End With
 InputBox "",, New GUIDGenerator
```
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Property|Generate|None|a GUID|Returns a unique GUID. Generate is the default property for the class, so the property name is optional. A sample GUID: {928507A9-7958-4E6E-A0B1-C33A5D4D602A}|
|Method|SetUppercase|None|N/A|Configure the Generate property to return uppercase, the default.|
|Method|SetLowercase|None|N/A|Configure the Generate property to return lowercase|

## HTAApp
HTAApp class  
Supports the VBSApp class, providing .hta functionality.  
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Method|Sleep|an integer|N/A|Pauses execution of the script or .hta for the specified number of milliseconds.|
|Method|PrepareToSleep|None|N/A|Required before calling the Sleep method when AlwaysPrepareToSleep is False in HTAApp.config.|
|Property|GetFilespec|None|a string|Returns the filespec of the calling .hta file.|
|Property|GetArgs|None|an array|Returns the mshta.exe command line args as an array, including the .hta filespec, which has index 0.|

## Includer
  
The Includer object helps with dependency management, and can be used in a .wsf, .vbs, or .hta script.  
  
How it works: The Read method returns the contents of a .vbs class file--or any other text file.  
  
Usage example  
```vb
 With CreateObject("VBScripting.Includer")
     Execute .read("WMIUtility.vbs") '.vbs may be omitted
     Execute .read("TextStreamer")
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
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Property|GetObj|className|An object|Returns an object based on the VBScript class with the specified name. Requires a .wsc Windows Script Component file in \class\wsc. See StringFormatter.wsc for an example.|
|Property|Read|a file|the file contents|Returns the contents of the specified file, which may be expressed either as an abolute path, or as a relative path relative to the <code> class</code> folder. The file name extension may be omitted for .vbs files.|
|Property|ReadFrom|file, path|file contents|Returns the contents of the specified file, which may be expressed either as an abolute path, or as a relative path relative to the path specified. The file name extension may be omitted for .vbs files.|
|Property|LibraryPath|None|a folder path|Returns the resolved, absolute path of the folder that contains Includer.wsc, which is the reference for relative paths passed to the Read and ReadFrom methods.|

## KeyDeleter
Deletes a registry key and all of its subkeys.  
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Method|DeleteKey|root, key|N/A|Deletes the specified registry key and all of its subkeys. Use one of the root constants for the first parameter.|
|Property|HKCR|None|&H80000000|Provides a value suitable for the first parameter of the DeleteKey method.|
|Property|HKCU|None|&H80000001|Provides a value suitable for the first parameter of the DeleteKey method.|
|Property|HKLM|None|&H80000002|Provides a value suitable for the first parameter of the DeleteKey method.|
|Property|HKU|None|&H80000003|Provides a value suitable for the first parameter of the DeleteKey method.|
|Property|HKCC|None|&H80000005|Provides a value suitable for the first parameter of the DeleteKey method.|
|Property|Result|None|an integer|Returns a code indicating the result of the most recent DeleteKey call. Codes can be looked up in <a href="https://msdn.microsoft.com/en-us/library/aa393978(v=vs.85).aspx">WbemErrEnum</a>|
|Property|Delete|a boolean|a boolean|Gets or sets the boolean that controls whether the key is actually deleted.|

## MathConstants
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Property|Pi|None|3.14159...|None|
|Property|DEGRAD|None|Pi/180|Used to convert degrees to radians|
|Property|RADEG|None|180/Pi|Used to convert radians to degrees|

## PrivilegeChecker
Default property Privileged returns True if the calling script has elevated privileges.  
Usage example  
```vb
 With CreateObject("VBScripting.Includer") 
     Execute .read("PrivilegeChecker") 
 End With 
 Dim pc : Set pc = New PrivilegeChecker 
 If pc Then 
     WScript.Echo "Privileges are elevated" 
 Else 
     WScript.Echo "Privileges are not elevated" 
 End If 
```
  
Reference: <a href="http://stackoverflow.com/questions/4051883/batch-script-how-to-check-for-admin-rights/21295806"> stackoverflow.com</a>  
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Property|Privileged|None|a boolean|Returns True if the calling script is running with elevated privileges, False if not. Privileged is the default property.|

## RegExFunctions
Regular Expression functions - a work in progress  
  
Usage example  
```vb
  With CreateObject("VBScripting.Includer")
      Execute .read("RegExFunctions")
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
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Property|Pattern|wildcard|a regex expression|Returns a regex expression equivalent to the specified wildcard expression(s). Delimit multiple wildcards with |.|
|Property|re|None|an object reference|Returns a reference to the RegExp object instance|
|Method|SetPattern|a regex pattern|N/A|Required before calling FirstMatch or GetSubMatches. Sets the pattern of the RegExp object instance|
|Method|SetTestString|a string|N/A|Required before calling FirstMatch or GetSubMatches. Specifies the string against which the regex pattern will be tested.|
|Method|SetIgnoreCase|a boolean|N/A|Optional. Specifies whether the regex object will ignore case. Default is False.|
|Method|SetGlobal|a boolean|N/A|Optional. Specifies whether the pattern should match all occurrences in the search string or just the first one. Default is False.|
|Property|GetSubMatches|None|an object|Returns the RegExp SubMatches object for the specified pattern and test string. The matches can be accessed with a For Each loop. See general usage comments. Work in progress. You must handle errors in case there are no matches.|
|Property|FirstMatch|None|a string|Regarding the string specified by SetTestString, returns the first substring in the string that matches the regex pattern specified by SetPattern.|

## RegistryUtility
Provides functions relating to the Windows&reg; registry  
  
Usage example  
```vb
  With CreateObject("VBScripting.Includer") 
      Execute .read("RegistryUtility") 
  End With 
  Dim reg : Set reg = New RegistryUtility 
  Dim key : key = "SOFTWARE\Microsoft\Windows NT\CurrentVersion" 
  MsgBox reg.GetStringValue(reg.HKLM, key, "ProductName") 
```
  
Set valueName to vbEmpty or "" (two double quotes) to specify a key's default value.  
  
StdRegProv docs <a href="https://msdn.microsoft.com/en-us/library/aa393664(v=vs.85).aspx"> online</a>.  
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Method|SetPC|a computer name|N/A|Optional. A dot (.) can be used for the local computer (default), in place of the computer name.|
|Property|GetStringValue|rootKey, subKey, valueName|a string|Returns the value of the specified registry location. The specified registry entry must be of type string (REG_SZ).|
|Method|SetStringValue|rootKey, subKey, valueName, value|N/A|Writes the specified REG_SZ value to the specified registry location. Writing to HKLM or HKCR requires elevated privileges.|
|Property|GetExpandedStringValue|rootKey, subKey, valueName|a string|Returns the value of the specified registry location. The specified registry entry must be of type REG_EXPAND_SZ.|
|Method|SetExpandedStringValue|rootKey, subKey, valueName, value|N/A|Writes the specified REG_EXPAND_SZ value to the specified registry location. Writing to HKLM or HKCR requires elevated privileges.|
|Property|HKLM|None|&H80000002|Represents HKEY_LOCAL_MACHINE. For use with the rootKey parameter.|
|Property|HKCU|None|&H80000001|Represents HKEY_CURRENT_USER. For use with the rootKey parameter.|
|Property|HKCR|None|&H80000000|Represents HKEY_CLASSES_ROOT. For use with the rootKey parameter.|
|Property|GetPC|None|a string|Returns the name of the current computer. <strong> .</strong> (dot) indicates the local computer.|
|Property|GetRegValueType|rootKey, subKey, valueName|an integer|Returns a registry key value type integer.|
|Method|EnumValues|rootKey, subKey, aNames, aTypes|N/A|Enumerates the value names and their types for the specified key. The aNames and aTypes parameters are populated with arrays of key value name strings and type integers, respectively. Wraps the StdRegProv EnumValues method, effectively fixing its <a href="https://groups.google.com/forum/#!topic/microsoft.public.win32.programmer.wmi/10wMqGWIfms"> lonely Default Value bug</a>, except that with HKCR and HKLM, elevated privileges are required or else aNames and aValues may be null if the default value is the only value.|
|Property|REG_SZ|None|1|Returns a registry value type constant.|
|Property|REG_EXPAND_SZ|None|2|Returns a registry value type constant.|
|Property|REG_BINARY|None|3|Returns a registry value type constant.|
|Property|REG_DWORD|None|4|Returns a registry value type constant.|
|Property|REG_MULTI_SZ|None|7|Returns a registry value type constant.|
|Property|REG_QWORD|None|11|Returns a registry value type constant.|
|Property|GetRegValueTypeString|rootKey, subKey, valueName|a string|Returns a registry key value type string suitable for use with WScript.Shell RegWrite method argument #3. That is, one of "REG_SZ", "REG_EXPAND_SZ", "REG_BINARY", or "REG_DWORD".|

## ShellConstants
Constants for use with WScript.Shell.Run  
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Property|RunHidden|None|0|Window opens hidden. <br /> For use with Run method parameter #2|
|Property|RunNormal|None|1|Window opens normal. <br /> For use with Run method parameter #2|
|Property|RunMinimized|None|2|Window opens minimized. <br /> For use with Run method parameter #2|
|Property|RunMaximized|None|3|Window opens maximized. <br /> For use with Run method parameter #2|
|Property|Synchronous|None|True|Script execution halts and waits for the called process to exit. <br /> For use with Run method parameter #3|
|Property|Asynchronous|None|False|Script execution proceeds without waiting for the called process to exit. <br /> For use with Run method parameter #3|

## SpecialFolders
An enum and wrapper for WScript.Shell.SpecialFolders  
Usage example  
```vb
     With CreateObject("VBScripting.Includer") 
         Execute .read("SpecialFolders") 
     End With 
   
     Dim sf : Set sf = New SpecialFolders 
     MsgBox sf.GetPath(sf.AllUsersDesktop) 'C:\Users\Public\Desktop 
```
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Property|GetPath|a special folder alias|a folder path|Returns the absolute path of the specified special folder. This is the default property, so the property name is optional.|
|Property|GetAliasList|None|a string|Returns a comma + space delimited list of the aliases of all the special folders.|
|Property|GetAliasArray|None|an array of strings|Returns an array of the aliases of all the special folders.|
|Property|AllUsersDesktop|None|a string|Returns a special folder alias having the exact same characters as the property name|
|Property|AllUsersStartMenu|None|a string|Returns a special folder alias having the exact same characters as the property name|
|Property|AllUsersPrograms|None|a string|Returns a special folder alias having the exact same characters as the property name|
|Property|AllUsersStartup|None|a string|Returns a special folder alias having the exact same characters as the property name|
|Property|Desktop|None|a string|Returns a special folder alias having the exact same characters as the property name|
|Property|Favorites|None|a string|Returns a special folder alias having the exact same characters as the property name|
|Property|Fonts|None|a string|Returns a special folder alias having the exact same characters as the property name|
|Property|MyDocuments|None|a string|Returns a special folder alias having the exact same characters as the property name|
|Property|NetHood|None|a string|Returns a special folder alias having the exact same characters as the property name|
|Property|PrintHood|None|a string|Returns a special folder alias having the exact same characters as the property name|
|Property|Programs|None|a string|Returns a special folder alias having the exact same characters as the property name|
|Property|Recent|None|a string|Returns a special folder alias having the exact same characters as the property name|
|Property|SendTo|None|a string|Returns a special folder alias having the exact same characters as the property name|
|Property|StartMenu|None|a string|Returns a special folder alias having the exact same characters as the property name|
|Property|Startup|None|a string|Returns a special folder alias having the exact same characters as the property name|
|Property|Templates|None|a string|Returns a special folder alias having the exact same characters as the property name|

## StringFormatter
 StringFormatter.vbs is the script for StringFormatter.wsc  
  
Provides string formatting functions  
  
Three instantiation examples:  
```vb
 With CreateObject("VBScripting.Includer") 
      Execute .read("StringFormatter") 
      Dim fm : Set fm = New StringFormatter 
 End With 
```
or   
```vb
 With CreateObject("VBScripting.Includer") 
      Dim fm : Set fm = .GetObj("StringFormatter") 
 End With 
```
or   
```vb
 Dim fm : Set fm = CreateObject("VBScripting.StringFormatter") 
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
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Property|Format|array|a string|Returns a formatted string. The parameter is an array whose first element contains the pattern of the returned string. The first %s in the pattern is replaced by the next element in the array. The second %s in the pattern is replaced by the next element in the array, and so on. Variant subtypes tested OK with %s include string, integer, and single. Format is the default property for the class, so the property name is optional. If there are too many or too few %s instances, then an error will be raised.|
|Method|SetSurrogate|a string|N/A|Optional. Sets the string that the Format method will replace with the specified array element(s), %s by default.|
|Property|Pluralize|count, noun|a string|Returns a string that may or may not be pluralized, depending on the specified count. If the noun has irregular pluralization, pass in a two-element array: <code> Split("person people")</code>. Otherwise, you may pass in either a singular noun as a string, <code> red herring</code>, or else a two-element array, <code> Split("red herring | red herrings", "|")</code>.|
|Method|SetZeroSingular|None|N/A|Optional. Changes the default behavior of considering a count of zero to be plural.|
|Method|SetZeroPlural|None|N/A|Optional. Restores the default behavior of considering a count of zero to be plural.|

## TestingFramework
A lightweight testing framework  
Usage example  
 ```vb
     With CreateObject("VBScripting.Includer") 
         Execute .read("VBSValidator") 
         Execute .read("TestingFramework") 
     End With 
     Dim val : Set val = New VBSValidator 'class under test 
     With New TestingFramework 
         .describe "VBSValidator class" 
         .it "should return False when IsBoolean is given a string" 
             .AssertEqual val.IsBoolean("sdfjke"), False 
         .it "should raise an error when EnsureBoolean is given a string" 
             Dim nonBool : nonBool = "a string" 
             On Error Resume Next 
                 val.EnsureBoolean(nonBool) 
                 .AssertErrorRaised 
                 Dim errDescr : errDescr = Err.Description 'capture the error information 
                 Dim errSrc : errSrc = Err.Source 
             On Error Goto 0 
     End With 
```
  
 See also VBSTestRunner  
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Method|describe|unit description|N/A|Sets the description for the unit under test. E.g. .describe "DocGenerator class"|
|Method|it|an expectation|N/A|Sets the specification, a.k.a. spec, which is a description of some expectation to be met by the unit under test. E.g. .it "should return an integer"|
|Property|GetSpec|None|a string|Returns the specification string for the current spec.|
|Method|ShowPendingResult|None|N/A|Flushes any pending results. Generally for internal use, but may occasionally be helpful prior to an ad hoc StdOut comment, so that the comment shows up in the output in its proper place.|
|Method|AssertEqual|actual, expected|N/A|Asserts that the specified two variants, of any subtype, are equal.|
|Method|AssertErrorRaised|None|N/A|Asserts that an error should be raised by one or more of the preceeding statements. The statement(s), together with the AssertErrorRaised statement, should be wrapped with an <br /> <pre style='white-space: nowrap;'> On Error Resume Next <br /> On Error Goto 0 </pre> block.|
|Method|DeleteFiles|an array|N/A|Deletes the specified files. The parameter is an array of filespecs. Relative paths may be used.|
|Property|MessageAppeared|None|a boolean|None|
|Method|ShowSendKeysWarning|None|N/A|Shows a SendKeys warning: a warning message to not make mouse clicks or key presses.|
|Method|CloseSendKeysWarning|None|N/A|Closes the SendKeys warning.|

## TextStreamer
Open a file as a text stream for reading, writing, or appending.  
<h5> Methods for use with the text stream that is returned by the Open method: </h5>  
<p> <em> Reading methods: </em> Read, ReadLine, ReadAll <br /> <em> Writing methods: </em> Write, WriteLine, WriteBlankLines <br /> <em> Reading or Writing methods: </em> Close, Skip, SkipLine <br /> <em> Reading or writing properties: </em> AtEndOfLine, AtEndOfStream, Column, Line </p>  
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Property|Open|None|an object|Returns a text stream object according to the specified settings (methods beginning with Set...)|
|Method|SetFile|a filespec|N/A|Specifies the file to be opened by the text streamer. Can include environment variable names. The default file is a random-named .txt file on the desktop.|
|Method|SetFolder|a folder|N/A|Specifies the folder of the file to be opened by the text streamer. Can include environment variables. Default is %UserProfile%\Desktop|
|Method|SetFileName|a file name|N/A|Specifies the file name, including extension, of the file to be opened by the text streamer. Default is a randomly named .txt file.|
|Method|SetForReading|None|N/A|Prepares the text stream to be opened for reading|
|Method|SetForWriting|None|N/A|Prepares the text stream to be opened for writing|
|Method|SetForAppending|None|N/A|Prepares the text stream to be opened for appending (default)|
|Method|SetCreateNew|None|N/A|Allows a new file to be created (default)|
|Method|SetDontCreateNew|None|N/A|Prevents a new file from being created if the file doesn't already exist|
|Method|SetAscii|None|N/A|Sets the expectation that the file will be Ascii (default)|
|Method|SetUnicode|None|N/A|Sets the expectation that the file will be Unicode|
|Method|SetSystemDefault|None|N/A|Uses Ascii or Unicode according to the system default|
|Method|View|None|N/A|Opens the file for viewing|
|Method|CloseViewer|None|N/A|Close the file viewer. From the docs: Use the Terminate method only as a last resort since some applications do not clean up properly. As a general rule, let the process run its course and end on its own. The Terminate method attempts to end a process using the WM_CLOSE message. If that does not work, it kills the process immediately without going through the normal shutdown procedure.|
|Method|SetViewer|filespec|N/A|Sets the filespec of an alternate file viewer to use with the View method.The default viewer is Notepad.|
|Method|Delete|None|N/A|Deletes the streamer file|
|Method|Run|None|N/A|Open/Run the file, assuming it has an executable file extension.|
|Property|GetFile|None|a filespec|Returns the filespec of the file that is open or set to be opened by the text streamer. Environment variables are not expanded.|
|Property|GetFileName|None|a file name|Returns the file name of the file that is open or set to be opened by the text streamer. Environment variables are not expanded.|
|Property|GetFolder|None|a folder|Returns the folder of the file that is open or set to be opened by the text streamer. Environment variables are not expanded.|
|Property|GetCreateMode|None|a boolean|Gets the current CreateMode setting. Returns one of these stream constants: bDontCreateNew or bCreateNew.|
|Property|GetStreamMode|None|an integer|Gets the current StreamMode setting. Returns one of these stream constants: iForReading, iForWriting, iForAppending|
|Property|GetStreamFormat|None|a tristate boolean|Gets the current StreamFormat setting. Returns one of these stream constants: tbAscii, tbUnicode, tbSystemDefault|

## TimeFunctions
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Method|SetFirstDOW|an integer|N/A|Specifies the first day of the week. Parameter can be one of the VBScript constants vbSunday, vbMonday, ...|
|Property|LetDOWBeAbbreviated|a boolean|N/A|Specifies whether day-of-the-week strings should be abbreviated: Default is False.|
|Property|TwoDigit|a number|a two-char string|Returns a two-char string that may have a leading 0, given a numeric integer/string/variant of length one or two|
|Property|DOW|a date|a day of the week|Returns a day of the week string, e.g. Monday, given a VBS date|
|Property|GetFormattedDay|a date|a date string|Returns a formatted day string; e.g. 2016-09-15-Sat|
|Property|GetFormattedTime|a date|a date string|Returns a formatted 24-hr time string: e.g. 13:38:45 or 00:45:32|

## ValidFileName
Provides for modifying a string to remove characters that are not suitable for use in a Windows&reg; file name.  
Usage Example  
```vb
     With CreateObject("VBScripting.Includer") 
         Execute .read("ValidFileName") 
     End With 
  
     MsgBox GetValidFileName("test\ing") 'test-ing 
```
  
ValidFileName.vbs provides an example of introductory comments in a script that lacks a Class statement: With DocGenerator.vbs, a line beginning with '''' (four single quotes) may be used instead of a Class statement, in order to end the introductory comments section.  
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Property|GetValidFileName|a file name candidate|a valid file name|Returns a string suitable for use as a file name: Removes <strong> \ / : * ? " < > | %20 # </strong> and replaces them with a hyphen/dash (-)|

## VBSApp
VBSApp class  
Intended to support identical handling of class procedures by .vbs/.wsf files and .hta files.  
This can be useful when writing a class that might be used in both types of "apps".  
Four ways to instantiate  
For .vbs/.wsf scripts,  
 ```vb
  Dim app : Set app = CreateObject("VBScripting.VBSApp") 
  app.Init WScript 
```
For .hta applications,  
 ```vb
  Dim app : Set app = CreateObject("VBScripting.VBSApp") 
  app.Init document 
```
If the script may be used in .vbs/.wsf scripts or .hta applications  
 ```vb
  With CreateObject("VBScripting.Includer") 
      Execute .read("VBSApp") 
  End With 
  Dim app : Set app = New VBSApp 
```
Alternate method for both .hta and .vbs/.wsf,  
 ```vb
  Set app = CreateObject("VBScripting.VBSApp") 
  If "HTMLDocument" = TypeName(document) Then 
      app.Init document 
  Else app.Init WScript 
  End If 
```
Examples  
 ```vb
  'test.vbs "arg one" "arg two" 
  With CreateObject("VBScripting.Includer") 
      Execute .read("VBSApp") 
  End With 
  Dim app : Set app = New VBSApp 
  MsgBox app.GetName 'test.vbs 
  MsgBox app.GetArg(1) 'arg two 
  MsgBox app.GetArgsCount '2 
  app.Quit 
```
  
 ```vb
  <!-- test.hta "arg one" "arg two" --> 
  <hta:application icon="msdt.exe"> 
      <script language="VBScript"> 
          With CreateObject("VBScripting.Includer") 
              Execute .read("VBSApp") 
          End With 
          Dim app : Set app = New VBSApp 
          MsgBox app.GetName 'test.hta 
          MsgBox app.GetArg(1) 'arg two 
          MsgBox app.GetArgsCount '2 
          app.Quit 
      </script> 
  </hta:application> 
```
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Property|GetArgs|None|array of strings|Returns an array of command-line arguments.|
|Property|GetArgsString|None|a string|Returns the command-line arguments string. Can be used when restarting a script for example, in order to retain the original arguments. Each argument is wrapped wih double quotes. The return string has a leading space, by design, unless there are no arguments.|
|Property|GetArg|an integer|a string|Returns the command-line argument having the specified zero-based index.|
|Property|GetArgsCount|None|an integer|Returns the number of arguments.|
|Property|GetFullName|None|a string|Returns the filespec of the calling script or hta.|
|Property|GetFileName|None|a string|Returns the name of the calling script or hta, including the filename extension.|
|Property|GetBaseName|None|a string|Returns the name of the calling script or hta, without the filename extension.|
|Property|GetExtensionName|None|a string|Returns the filename extension of the calling script or hta.|
|Property|GetParentFolderName|None|a string|Returns the folder that contains the calling script or hta.|
|Property|GetExe|None|a string|Returns "mshta.exe" to hta files, and "wscript.exe" or "cscript.exe" to scripts, depending on the host.|
|Method|RestartWith|#1: host; #2: switch; #3: elevating|N/A|Restarts the script/app with the specified host (typically "wscript.exe", "cscript.exe", or "mshta.exe") and retaining the command-line arguments. Paramater #2 is a cmd.exe switch, "/k" or "/c". Parameter #3 is a boolean, True if restarting with elevated privileges. If userInteractive, first warns user that the User Account Control dialog will open.|
|Method|SetUserInteractive|boolean|N/A|Sets userInteractive value. Setting to True can be useful for debugging. Default is True.|
|Property|GetUserInteractive|None|boolean|Returns the userInteractive setting. This setting also may affect the visibility of selected console windows.|
|Method|SetVisibility|0 (hidden) or 1 (normal)|N/A|Sets the visibility of selected command windows. SetUserInteractive also affects this setting. Default is True.|
|Property|GetVisibility|None|0 (hidden) or 1 (normal)|Returns the current visibility setting. SetUserInteractive also affects this setting.|
|Method|Quit|None|N/A|Gracefully closes the hta/script.|
|Method|Sleep|an integer|N/A|Pauses execution of the script or .hta for the specified number of milliseconds.|
|Property|WScriptHost|None|"wscript.exe"|Can be used as an argument for the method RestartIfNotPrivileged.|
|Property|CScriptHost|None|"cscript.exe"|Can be used as an argument for the method RestartIfNotPrivileged.|
|Property|GetHost|None|"wscript.exe" or "cscript.exe" or "mshta.exe"|Returns the current host. Can be used as an argument for the method RestartIfNotPrivileged.|

## VBSArguments
Functions related to VBScript command-line arguments  
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Property|GetArgumentsString|None|a string containing all command-line arguments|For use when restarting a script, in order to retain the original arguments. Each argument is wrapped wih quotes, which are stripped off as they are read back in. The return string has a leading space, by design, unless there are no arguments|

## VBSArrays
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Property|Uniques|an array|an array|Returns an array with no duplicate items, given an array that may have some.|
|Property|RemoveFirstElement|an array of strings|an array of strings|Returns a array without the first element of the specified array.|
|Property|CollectionToArray|a collection of strings|array of strings|Can be used to convert the WScript.Arguments object to an array, for example.|

## VBSClipboard
Clipboard procedures  
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Method|SetClipboardText|a string|N/A|Copies the specified string to the clipboard. Uses clip.exe, which shipped with Windows&reg; Vista / Server 2003 through Windows 10.|
|Property|GetClipboardText|None|a string|Returns text from the clipboard|

## VBSEnvironment
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Property|Expand|a string|a string|Expands environment variable(s); e.g. convert %UserProfile% to C:\Users\user42|
|Property|Collapse|a string|a string|Collapses a string that may contain one or more substrings that can be shortened to an environment variable.|
|Method|CreateUserVar|varName, varValue|N/A|Create or set a user environment variable|
|Method|SetUserVar|varName, varValue|N/A|Set or create a user environment variable|
|Property|GetUserVar|a variable name|the variable value|Returns the value of the specified user environment variable|
|Method|RemoveUserVar|varName|N/A|Removes a user environment variable|
|Method|CreateProcessVar|varName, varValue|N/A|Create a process variable|
|Method|SetProcessVar|varName, varValue|N/A|Sets or creates a process environment variable|
|Property|GetProcessVar|varName|the variable value|Returns the value of the specified environment variable|
|Method|RemoveProcessVar|varName|N/A|Removes the specified process environment variable|
|Property|GetDefaults|None|an array|Returns an array of common environment variables pre-installed with some versions of Windows&reg;. Not exhaustive.|

## VBSEventLogger
Logs messages to the Application event log.  
  
Wraps the LogEvent method of the WScript.Shell object.  
  
To see a log entry, type EventVwr at the command prompt to open the Event Viewer, expand Windows Logs, and select Application. The log Source will be WSH. Or you can use the CreateCustomView method to create an entry in the Event Viewer's Custom Views section.  
  
Usage example:  
 ```vb
  With CreateObject("VBScripting.Includer") 
      Execute .read("VBSEventLogger") 
  End With 
   
  Dim logger : Set logger = New VBSEventLogger 
  logger.log logger.INFORMATION, "message 1" 
  logger logger.INFORMATION, "message 2" 
  logger 4, "message 3" 
  logger 1, "error message" 
   
  logger.CreateCustomView 'create a custom view in the Event Viewer 
  logger.OpenViewer 'open EventVwr.msc 
```
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Method|Log|eventType, message|N/A|Adds an event entry to a log file with the specified message. This is the default method, so the method name is optional.|
|Method|CreateCustomView|None|N/A|Creates a Custom View in the Event Viewer, eventvwr.msc, named WSH Logs. The User Account Control dialog will open, in order to confirm elevation of privileges. Based on VBSEventLoggerCustomView.xml.|
|Method|OpenViewer|None|N/A|Opens the Windows&reg; Event Viewer, eventvwr.msc|
|Property|SUCCESS|None|0|Returns a value for use as an "eventType" parameter|
|Property|ERROR|None|1|Returns a value for use as an "eventType" parameter|
|Property|WARNING|None|2|Returns a value for use as an "eventType" parameter|
|Property|INFORMATION|None|4|Returns a value for use as an "eventType" parameter|
|Property|AUDIT_SUCCESS|None|8|Returns a value for use as an "eventType" parameter|
|Property|AUDIT_FAILURE|None|16|Returns a value for use as an "eventType" parameter|
|Method|OpenConfigFolder|None|N/A|Opens the Event Viewer configuration folder, by default "%ProgramData%\Microsoft\Event Viewer". The Views subfolder contains the .xml files defining the custom views.|
|Method|OpenLogFolder|None|N/A|Opens the folder with the .evtx files that contain the event logs, by default "%SystemRoot%\System32\Winevt\Logs". Application.evtx holds the WSH data.|

## VBSExtracter
For extracting a string from a text file, given a regular expression  
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Method|SetPattern|a regex pattern|N/A|Required. Specifies the text to be extracted. Non-regex expressions containing any of the regex special characters <strong>(  )  .  $  +  [  ?  \  ^  {  |</strong> must preceed the special character with a <strong>\</strong>|
|Method|SetFile|filespec|N/A|Required. Specifies the file to extract text from.|
|Method|SetIgnoreCase|a boolean|N/A|Set whether to ignore case when matching text. Default is False.|
|Property|Extract|None|a string|Returns the first string that matches the specified regex pattern. Returns an empty string if there is no match. Before calling this method, you must specify the file and the pattern: see SetPattern and SetFile.|
|Property|Extract0|None|a string|Deprecated for not spanning multiple lines. Formerly named Extract. Returns the string that matches the specified regex pattern. Returns an empty string if there is no match. Before calling this method, you must specify the file and the pattern: see SetPattern and SetFile.|

## VBSFileSystem
General utility functions  
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Property|SBaseName|None|a file name, no extension|Returns the name of the calling script, without the file name extension.|
|Property|SName|None|a file name|Returns the name of the calling script, including file name extension|
|Property|SFullName|None|a filespec|Returns the filespec of the calling script|
|Property|SFolderName|None|a folder|Returns the parent folder of the calling script.|
|Property|MakeFolder|a path|a boolean|Create a folder, and if necessary create also its parent, grandparent, etc. Returns False if the folder could not be created.|
|Property|Parent|a folder, file, or registry key|the item's parent|Returns the parent of the folder or file or registry key, or removes a trailing backslash. The parent need not exist.|
|Method|SetReferencePath|a path|N/A|Optional. Specifies the base path from which relative paths should be referenced. By default, the reference path is the parent folder of the calling script. See also Resolve and ResolveTo.|
|Property|Resolve|a relative path|a resolved path|Resolves a relative path (e.g. "../lib/WMI.vbs"), to an absolute path (e.g. "C:\Users\user42\lib\WMI.vbs"). The relative path is by default relative to the parent folder of the calling script, but this behavior can be changed with SetReferencePath. See also property ResolveTo.|
|Property|ResolveTo|relativePath, absolutePath|a resolved path|Resolves the specified relative path, e.g. "../lib/WMI.vbs", relative to the specified absolute path, and returns the resolved absolute path, e.g. "C:\Users\user42\lib\WMI.vbs". Environment variables are allowed.|
|Property|Expand|a string|an expanded string|Expands environment strings. E.g. %WinDir% => C:\Windows|
|Method|Elevate|command, arguments, folder|N/A|Runs the specified command with elevated privileges, with the specified arguments and working folder|
|Property|FoldersAreTheSame|folder1, folder2|a boolean|Determines whether the two specified folders are the same. If so, returns True.|
|Method|DeleteFile|filespec|N/A|Deletes the specified file.|
|Method|SetForceDelete|boolean|N/A|Controls the behavior of the DeleteFile method: Specify True to force a file deletion. Optional. Default is False.|

## VBSHoster
Manage which script host is hosting the currently running script  
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Method|EnsureCScriptHost|None|N/A|Restart the script hosted with CScript if it isn't already hosted with CScript.exe|
|Method|SetSwitch|/k or /c|N/A|Optional. Specifies a switch for %ComSpec% for use with the EnsureCScriptHost method: controls whether the command window, if newly created, remains open (/k). Useful for troubleshooting, in order to be able to read error messages. Unnecessary if starting the script from a console window, because /c is the default.|
|Method|SetDefaultHostWScript|None|N/A|Sets wscript.exe to be the default script host. The User Account Control dialog will open for permission to elevate privileges.|
|Method|SetDefaultHostCScript|None|N/A|Sets cscript.exe to be the default script host. The User Account Control dialog will open for permission to elevate privileges.|

## VBSLogger
A lightweight VBScript logger  
Instantiation   
```vb
     With CreateObject("VBScripting.Includer") 
         Execute .read("VBSLogger") 
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
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Method|Log|a string|N/A|Opens the log file, writes the specified string, then closes the log file. This is the default method for the VBSLogger class.|
|Method|SetLogFolder|a folder path|N/A|Optional. Customize the log folder. The folder will be created if it does not exist. Environment variables are allowed. See GetDefaultLogFolder.|
|Method|Open|None|N/A|Opens the log file for writing. The log file is opened and remains open for writing. While it is open, other processes/scripts will be unable to write to it.|
|Method|Write|a string|N/A|Writes the specified string to the log file.|
|Method|Close|None|N/A|Closes the log file text stream, enabling other process to write to it.|
|Method|View|None|N/A|Opens the log file for viewing. Notepad is the default editor. See SetViewer.|
|Method|SetViewer|a filespec|N/A|Optional. Customize the program that the View method uses to view log files. Default: Notepad.|
|Method|ViewFolder|None|N/A|Open the log folder|
|Property|WordPad|None|a filespec|Can be used as the argument for the SetViewer method in order to open files with WordPad when the View method is called.|
|Property|GetDefaultLogFolder|None|a folder|Retrieves the default log folder, %AppData%\VBScripts\logs|
|Property|GetLogFilePath|None|a filespec|Retreives the filespec for the log file, with environment variables expanded. Default: &lt;GetDefaultLogFolder&gt;\YYYY-MM-DD-DayOfWeek.txt|

## VBSPower
Power functions: shutdown, restart, logoff, sleep, and hibernate.  
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Property|Shutdown|None|a boolean|Shuts down the computer. Returns True if the operation completes with no errors.|
|Property|Restart|None|a boolean|Restarts the computer. Returns True if the operation completes with no errors.|
|Property|Logoff|None|a boolean|Logs off the computer. Returns True if the operation completes with no errors.|
|Method|Sleep|None|N/A|Puts the computer to sleep. Requires <a href="https://docs.microsoft.com/en-us/sysinternals/downloads/psshutdown"> PsTools</a> download and PsShutdown.exe to be located somewhere on your %Path%. Recovery from sleep is faster than from hibernation, but uses more power.|
|Method|Hibernate|None|N/A|Puts the computer into hibernation. Will not work if hibernate is disabled in the Control Panel, in which case the EnableHibernation method may be used to reenable hibernation. Hibernate is more power-efficient than sleep, but recovery is slower. If the computer wakes after pressing a key or moving the mouse, then it was sleeping, not in hibernation. Recovery from hibernation typically requires pressing the power button.|
|Method|EnableHibernation|None|N/A|Enables hibernation. The User Account Control dialog will open to request elevated privileges.|
|Method|DisableHibernation|None|N/A|Disables hibernation. The User Account Control dialog will open to request elevated privileges.|
|Method|SetForce|force|N/A|Optional. Setting this to True forces the Shutdown or Restart, discarding unsaved work. Default is False. Logoff always forces apps to close.|
|Method|SetDebug|a boolean|N/A|Used for testing. True prevents the computer from actually shutting down, etc., during testing. Default is False.|

## VBSTestRunner
Run a test or group of tests  
Usage example  
 ```vb
    'test-launcher.vbs 
    'run this file from a console window; e.g. cscript //nologo test-launcher.vbs 
   
     With CreateObject("VBScripting.Includer") 
         Execute .read("VBSTestRunner") 
     End With 
   
     With New VBSTestRunner 
         .SetSpecFolder "../spec" 'location of test files relative to test-launcher.vbs 
         .Run 
     End With 
```
  
See also TestingFramework  
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Method|Run|None|N/A|Initiate the specified tests|
|Method|SetSpecFolder|a folder|N/A|Optional. Specifies the folder containing the test files. Can be a relative path, relative to the calling script. Default is the parent folder of the calling script.|
|Method|SetSpecPattern|a regular expression|N/A|Optional. Specifies which file types to run. Default is *.spec.vbs. Standard wildcard notation with | delimiter.|
|Method|SetSpecFile|a file|N/A|Optional. Specifies a single file to test. Include the filename extension. E.g. SomeClass.spec.vbs. A relative path is OK, relative to the spec folder. If no spec file is specified, all test files matching the specified pattern will be run. See SetSpecPattern.|
|Method|SetSearchSubfolders|a boolean|N/A|Optional. Specifies whether to search subfolders for test files. True or False. Default is False.|
|Method|SetPrecision|0, 1, or 2|N/A|Optional. Sets the number of decimal places for reporting the elapsed time. Default is 2.|
|Method|SetRunCount|an integer|N/A|Optional. Sets the number of times to run the test(s). Default is 1.|

## VBSTimer
A timer  
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Property|Split|None|a rounded number (Single)|Returns the seconds elapsed since object instantiation or since calling the Reset method. Split is the default Property.|
|Method|SetPrecision|0, 1, or 2|N/A|Sets the number of decimal places to round the Split function return value. Default is 2.|
|Property|GetPrecision|None|0, 1, or 2|Returns the current precision.|
|Method|Reset|None|N/A|Sets the timer to zero.|

## VBSTroubleshooter
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Method|LogAscii|a string|N/A|Write to the log the Ascii codes for each character in the specified string.|

## VBSValidator
A working example of how validation can be accomplished.  
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Property|GetClassName|None|the class name|Returns                           "VBSValidator". Useful for verifying Err.Source in a unit test.|
|Property|IsBoolean|a boolean candidate|a boolean|Returns True if the parameter is a boolean subtype; False if not.|
|Method|EnsureBoolean|a boolean candidate|N/A|Raises an error if the parameter is not a boolean|
|Property|IsInteger|an integer candidate|a boolean|Returns True if the parameter is an integer subtype; False if not.|
|Method|EnsureInteger|an integer candidate|N/A|Raises an error if the parameter is not an integer|
|Property|ErrDescrBool|None|a string|" is not a boolean." Useful for verifying Err.Description in a unit test.|
|Property|ErrDescrInt|None|a string|" is not an integer." Useful for verifying Err.Description in a unit test.|

## WindowsUpdatesPauser
Pause Windows Updates to get more bandwidth. Don't forget to resume.  
For configuration settings, see the .config file in %LocalAppData% that has the same base name as the calling script/hta.  
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Method|PauseUpdates|None|N/A|Pauses Windows Updates.|
|Method|ResumeUpdates|None|N/A|Resumes Windows Updates.|
|Property|GetStatus|None|a string|Returns Metered or Unmetered. If Metered, then Windows Updates has paused to save money, incidentally not soaking up so much bandwidth. If TypeName(GetStatus) = "Empty", then the status could not be determined, possibly due to a bad network name (internal name: profileName).|
|Property|GetAppName|None|a string|Returns the base name of the calling script|
|Property|GetProfileName|None|a string|Returns the name of the network. The name is set by editing WindowsUpdatesPauser.config|
|Property|GetServiceType|None|a string|Returns the service type|
|Method|OpenConfigFile|None|N/A|Opens the .config file|

## WMIUtility
Examples of the Windows Management Instrumentation object  
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Property|TerminateProcessById|process id|a boolean|Terminates any Windows&reg process with the specified id. Returns True if the process was found, False if not.|
|Property|TerminateProcessByIdAndName|id, name|a boolean|Terminates a process with the specified id and name. Returns True if the process was found, False if not.|
|Method|TerminateProcessByIdAndNameDelayed|id, name, milliseconds|N/A|Terminates a process with the specified id (integer), name (string, e.g. notepad.exe), and delay (integer: milliseconds), asynchronously.|
|Property|GetProcessIDsByName|a process name|a boolean|Returns an array of process ids that have the specified name. The process name is what would appear in the Task Manager's Details tab. <br /> E.g. <code> notepad.exe</code>.|
|Property|GetProcessesWithNamesLike|a string like jav%|an array of process names|None|
|Property|IsRunning|a process name|a boolean|Returns a boolean indicating whether at least one instance of the specified process is running. <br /> E.g. <code> wmi.IsRunning("notepad.exe") 'True or False</code>.|
|Property|partitions|None|a collection|Returns a collection of partition objects, each with the following methods: Caption, Name, DiskIndex, Index, PrimaryPartition, Bootable, BootPartition, Description, Type, Size, StartingOffset, BlockSize, DeviceID, Access, Availability, ErrorMethodology, HiddenSectors, Purpose, Status|
|Property|disks|None|a collection|Returns a collection of disk objects, each with these methods: FileSystem, DeviceID|
|Property|cpu|None|an object|Returns an object with these methods: Architecture, Description|
|Property|os|None|an object|Return an OS object with these methods: Name, Version, Manufacturer, WindowsDirectory, Locale, FreePhysicalMemory, TotalVirtualMemorySize, FreeVirtualMemory, SizeStoredInPagingFiles|
|Property|pc|None|an object|Returns a PC object with these methods: Name, Manufacturer, Model, CurrentTimeZone, TotalPhysicalMemory|
|Property|Bios|None|an object|Returns a BIOS object with this method: Version|

## WoWChecker
Provides an object whose default property, isWoW, returns a boolean indicating whether the calling script was itself called by a SysWoW64 (32-bit) .exe file.  
  
How it works: .exe files in %SystemRoot%\System32 and %SystemRoot%\SysWoW64 are compared by size or checksum. If the files are the same, then the calling script must be running in a 32-bit process.  
  
Usage examples  
```vb
 MsgBox New WoWChecker.BySize.isWoW 
 MsgBox New WoWChecker.isWoW 
 With New WoWChecker : .BySize : MsgBox .isWoW : End With 
 With New WoWChecker.BySize : MsgBox .isWoW : End With 
 MsgBox New WoWChecker 
```
  
| Procedure | Name | Parameter | Return | Comment |
| :-------- | :--- | :-------- | :----- | :------ |
|Property|OSIs64Bit|None|a boolean|Returns a boolean that indicates whether the Windows OS is 64-bit.|
|Property|isWoW|None|a boolean|Returns a boolean that indicates whether the calling script was itself called by a SysWoW64 (32-bit) .exe file. This is the class default property.|
|Property|isSysWoW64|None|a boolean|Wraps isWoW: Same as calling isWoW.|
|Property|isSystem32|None|a boolean|Returns the opposite of isSysWoW64|
|Property|BySize|None|an object self reference|Optional. Specifies that the .exe files will be compared by size. BySize will not distinguish between the 32- and 64-bit .exe files if they are the same size, which is unlikely but possible. ByCheckSum is therefore more reliable.|
|Property|ByCheckSum|None|an object self reference|Selected by default. Specifies that the .exe files will be compared by checksum. ByCheckSum uses CertUtil, which ships with Windows&reg; 7 through 10, and can be manually installed on older versions.|
