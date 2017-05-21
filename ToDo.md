## To Do

#### Refactor applicable classes to ...

- Support supressing user interactivity, except for errors. Mitigating factor: most classes already do.  
- Better support for `.hta` files. (The `WScript` object is not available in `.hta` files).  
- Support having a changeable working directory, where applicable. Current behavior: for many classes, the working directory is assumed to be the folder containing the calling script. Mitigating factor: usually this is OK.  
- Support Visual Studio intellisense for native objects: Don't use wrapped native objects.  

#### Bug fixes

- `WindowsUpdatePauser` class: Improve handling of switching networks. Desired behavior: change the settings of the currently connected Wi-Fi network, not the one specified in the `.config` file.  
- Fix failing test in `DotNetCompiler.spec.elev.vbs`.  


#### Also ...
- Add a class to support local, remote, and class configuration data.  
- Move the `RestartIfNotPrivileged` method from the `DotNetCompiler` class to the `PrivilegeChecker` class or to some other class, for reusability.  
- Use the `VBSArgument` class's `GetArgumentsString` property when restarting a script: in `DotNetCompiler`, `PrivilegeChecker`, `VBSHoster`.
- Add an Expand all button to the docs, for searchability.  
- In the `VBSPower` class, add Suspend method (?), which would either call the `Sleep` or `Hibernate` method, depending on whether hibernate is enabled.  
- Refactor the DocGenerator class for readability, best practices, etc.