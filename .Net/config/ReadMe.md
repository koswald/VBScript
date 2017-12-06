###### The `config` folder

# ReadMe.md

### exeLocations.bat

The locations of `csc.exe` and `regasm.exe` 
are configurable in [exeLocations].bat
and [DotNetCompiler.config].

### CreateEventSource.vbs

After compiling and registering the libraries (see 
the `build` folder [ReadMe]), running [CreateEventSource].vbs 
is recommended in order for errors to be 
logged with the `VBScripting` source name, rather 
than with the somewhat misleading `WSH` source name.

### Recommended git configuration

If and when you change configurations files, it is 
recommended that you don't check in the change 
into the remote `git` repository.  

The following command is recommended to be run 
from git bash for that purpose, before staging the change(s).

```
git update-index --assume-unchanged **/*.config **/**/*.config .Net/config/exeLocations.bat .Net/key/generate-key-pair.vbs .Net/rsp/_common.rsp
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