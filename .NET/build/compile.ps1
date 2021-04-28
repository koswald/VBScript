<#
    .Description
    Compile the C# files using the 32-bit .NET Framework C# version 5 compiler, but don't register them.

    .Notes
    For this project, .cs files are compiled by the 32-bit compiler. The resulting binaries are later registered for 32- and 64-bit apps. This file was used for testing a procedure that registered the assemblies on a per-user basis. The procedure was abandoned because per-user, HKey_Current_User configurations are not available to processes running with elevated privileges.

    .Link
    https://stackoverflow.com/questions/46217402/how-to-extract-a-certain-part-of-a-string-in-powershell
#>

# set the current directory to this script's location
$PSScriptRoot | Set-Location

# extract from a file the path for the 32-bit csc.exe (CSharp Compiler)
$content = Get-Content "..\config\exeLocations.bat"
<# pattern explanation:
[\w:\\]? matches a possible C:\ or D:\, etc. 
[%\w]+ matches Windows or %SystemRoot%, for example 
v\d{1}... matches v4.0.1234 or v4.5.1234567, for example #>
$pattern32 = "[\w:\\]?[%\w]+\\Microsoft.NET\\Framework\\v\d{1}\.\d{1}\.\d{4,7}" 
$framework32Path = [regex]::matches($content, $pattern32).value
$framework32Path = [System.Environment]::ExpandEnvironmentVariables($framework32Path)

# get the base names of all of the C# files in the parent folder
Get-ChildItem ..\*.cs | ForEach-Object {
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($_)

    # compile
    "`nCompiling $baseName.cs"
    & "$framework32Path\csc.exe" '@..\rsp\_common.rsp' "@..\rsp\$baseName.rsp" ..\$baseName.cs
}

"`nCompiling project files is finished."
"Press a key to exit"
[void] $host.UI.RawUI.ReadKey()
