## To Do

- Refactor the DocGenerator class for readability, best practices, etc.  
- Add a class to support local, remote, and class configuration data.  
- Move the `RestartIfNotPrivileged` method from the `DotNetCompiler` class 
to the `PrivilegeChecker` class or to some other class, for reusability.  
- Use the `VBSArgument` class's `GetArgumentsString` property when restarting a script: 
in `DotNetCompiler`, `PrivilegeChecker`, `VBSHoster`.
- Add an Expand all button to the docs, for searchability.  
- In the `VBSPower` class, consider adding a Suspend method,
which would either call the `Sleep` or `Hibernate` method,
depending on whether hibernate is enabled. 
Consider raising an error if the hibernate method is called when hibernate is disabled.
- `WindowsUpdatePauser` class: Improve handling of switching networks. 
Desired behavior: change the settings of the currently connected Wi-Fi network,
not the one specified in the `.config` file.  

#### Refactor all classes, if appropriate, to ...

- Support supressing user interactivity. 
When user interactivity is suppressed, make raised errors unobtrusive.
- Better support for `.hta` files. 
This is necessary because the `WScript` object is not available in `.hta` files,
so, for example, WScript.ScriptName does not give the filespec of the source file in an `.hta`.
Also, the command-line arguments are accessed differently in an `.hta` script.
See DotNetCompiler.vbs Sub Class_Initialize for a rudimentary example of
how these concerns might be addressed.
- Support having a changeable working directory, where applicable. 
Current behavior: for many classes, the working directory is assumed to be the folder 
that contains the calling script. Mitigating factor: usually this is OK.  
- Support Visual Studio intellisense for native objects: Don't use wrapped native objects.  
