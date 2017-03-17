## To Do

#### Refactor each class ...

to support supressing user interactivity.  
to support .hta files (WScript object is not available in .hta's).  
to support having a changeable working directory, if applicable.  
to support Visual Studio intellisense for native objects (don't use wrapped native objects)
to change ExecuteGlobal(.read .. to Execute(.read, and remove error handling if appropriate

#### Also ...
add a class to support local, remote, and class configuration data
Move the .RestartIfNotPrivileged method from the DotNetCompiler class to the PrivilegeChecker class; use the VBSArgument' GetArgumentsString Property

