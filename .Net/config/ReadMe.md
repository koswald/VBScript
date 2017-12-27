###### The `config` folder

# ReadMe.md

### exeLocations.bat

The locations of `csc.exe` and `regasm.exe` 
are configurable in [exeLocations].bat
and [DotNetCompiler.config]. 

### CreateEventSource.vbs

`CreateEventSource.vbs` is called automatically by running 
`Setup.vbs` in the project's root folder, adding the event 
log source `VBScripting` to the Application log. 
It can be run again as a quick way to verify that the 
source has been added.

### Recommended git configuration

If and when you change configurations files, it is 
recommended that you don't check in the change 
into the remote `git` repository.  

The following command is recommended to be run 
from git bash for that purpose, before staging the change(s).

```
git update-index --assume-unchanged **/*.config **/**/*.config .Net/config/exeLocations.bat .Net/rsp/_common.rsp
```

To see the affected files, run

```
git ls-files -v | grep '^h'
```

To undo the index update, run the `update-index` command as above except with `--no-assume-unchanged`


[exeLocations]: ./exeLocations.bat
[DotNetCompiler.config]: ../../class/DotNetCompiler.config
[CreateEventSource]: ./CreateEventSource.vbs
[ReadMe]: ../build/ReadMe.md