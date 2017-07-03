## To Do

- Better support for `.hta` files. 
  This is desirable because the `WScript` object is not available in 
  `.hta` files. So, for example, WScript.ScriptName cannot give the 
  filespec of the source file with an `.hta`. 
  Also, the command-line arguments are accessed differently in an 
  `.hta` script/app.
  Implement this by refactoring all (?) other classes to use the VBSApp class.
  Remove dependencies from the VBSApp class.
  - In order to minimize the potential for dependency conficts,
    consider deprecating `VBSArguments`, `VBSHoster`, and
   `PrivilegeChecker` and moving their functions to `VBSApp`.
    Keeping the classes separate if possible, will support bloat
    reduction, but keep in mind the requirement to support both
    `.hta`s and scripts.
  - Consider removing the VBSArrays dependency from the VBSApp class, 
    if necessary.
  - Consider removing the VBSExtracter dependency from the VBSApp class,
    if necessary.

- Improve dependency management.
  - Minimize dependencies.
    - Deprecate `VBSNatives`, `StreamConstants`, `VBSFileSystem`.
  - Within each class, place all dependencies on other classes in the
    `CLass_Initialize` method and also in the introductory comments.

- Add a class to support local, cloud, class, and app/script 
  config data. Use this to refactor the `WindowsUpdatesPauser` class.

- In the `VBSPower` class, consider adding a Suspend method,
  which would either call the `Sleep` or `Hibernate` method,
  depending on whether hibernate is enabled. 
  Consider raising an error if the hibernate method is called when 
  hibernate is disabled.

- `WindowsUpdatePauser` class: Improve handling of switching networks. 
  Desired behavior: change the settings of the currently connected 
  Wi-Fi network, not the one specified in the `.config` file.  

- Support supressing user interactivity. When user interactivity 
is suppressed, make raised errors unobtrusive.

- Support having a changeable working directory, where applicable. 
  Current behavior: for many classes, the working directory is 
  assumed to be the folder that contains the calling script. 
  Mitigating factor: usually this is OK.  

- Support Visual Studio intellisense for native objects: 
  Don't use wrapped native objects such as WScript.Script, 
  Scripting.FileSystemObject, etc. 
  For complete list, see VBSNatives class. 
  Instead, instantiate them as needed within each class.  

#### Documentation

- Refactor the DocGenerator class for readability, 
  best practices, etc.  

- Add an Expand all button to the docs, for searchability.  

- Write the docs in Markdown format too, not just HTML, 
  or otherwise support viewing viewing in GitHub.

- Add code highlighting to the docs.

---

### Done

- Move the `RestartIfNotPrivileged` method from the 
`DotNetCompiler` class to the `VBSApp` class 
  for reusability.  
