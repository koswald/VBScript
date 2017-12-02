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

[exeLocations]: ./exeLocations.bat
[DotNetCompiler.config]: ../../class/DotNetCompiler.config
[CreateEventSource]: ./CreateEventSource.vbs
[ReadMe]: ../build/ReadMe.md