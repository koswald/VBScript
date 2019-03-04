
#### The Lonely Default Value Bug

A [longstanding bug] in the WMI [StdRegProv] EnumValues method exhibits itself when enumerating a registry key that has only one value, the default value. In this situation, the resulting names and values arrays may be null.  

A solution for this bug has been included in `Fixer.hta` and in the RegistryUtility class, which has an EmumValues method alternative to the StdRegProv method of the same name. It is recommended to use the RegistryUtility class method EnumValues rather than the StdRegProv method. Keep in mind that using this method on HKLM and HKCR will require elevated privileges for the fix to take effect.  

[longstanding bug]: https://groups.google.com/forum/#!topic/microsoft.public.win32.programmer.wmi/10wMqGWIfms
[StdRegProv]: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/regprov/stdregprov