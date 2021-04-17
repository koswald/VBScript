# The `config` folder

## exeLocations.bat

The locations of `csc.exe` and `regasm.exe` are configurable in [exeLocations.bat]. It is expected that it would need to be changed only for rare circumstances.

## Recommended git configuration

If and when you change configurations files, it is recommended that you don't check in the change into the remote `git` repository.  

The following command is recommended to be run from git bash for that purpose, before staging the change(s).

``` bash
git update-index --assume-unchanged **/*.config .NET/config/exeLocations.bat .NET/rsp/_common.rsp
```

To see the affected files, run

``` bash
git ls-files -v | grep '^h'
```

To undo the index update, run the `update-index` command as above except with `--no-assume-unchanged`

> Note: There is a private ignore file at `.git\info\exclude` in the project folder.

[exeLocations.bat]: ./exeLocations.bat
[ReadMe]: ../build/ReadMe.md
