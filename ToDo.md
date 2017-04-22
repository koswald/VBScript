## To Do

#### Refactor applicable classes to ...

- Support supressing user interactivity, except for errors. Mitigating factor: most classes already do.  
- Better support .hta files (WScript object is not available in .hta's).  
- Support having a changeable working directory, where applicable. Current behavior: for many classes, the working directory is assumed to be the folder containing the calling script. Mitigating factor: usually this is OK.  
- Support Visual Studio intellisense for native objects. (Don't use wrapped native objects.)  
- Change ExecuteGlobal(.read .. to Execute(.read, and remove error handling if appropriate.

#### Bug fixes

- WindowsUpdatePauser class: Improve handling of switching networks. Desired behavior: change the settings of the currently connected Wi-Fi network, not the one specified in the .config file.

#### Also ...
- Add a class to support local, remote, and class configuration data.  
- Move the .RestartIfNotPrivileged method from the DotNetCompiler class to the PrivilegeChecker class or to some other class, for reusability.  
- Use the VBSArgument's GetArgumentsString Property when restarting a script: in DotNetCompiler, PrivilegeChecker, VBSHoster.
- Add an Expand all button to the docs, for searchability.

