
'Register a .wsc file for 64-bit and 32-bit apps.

'Drop a file onto this script or a shortcut to it,
'or use the command line with absolute path,
'relative path, or environment variable.

'get argument and validate
With WScript.Arguments
    If 0 = .Count Then
        Err.Raise 1,, "Argument expected: the .wsc file to register."
    End If
    file = Resolve(Expand(.item(0)))
End With

'prepare to register
args = "/c" & _
    " %SystemRoot%\System32\regsvr32.exe """ & file & """ & " & _
    " %SystemRoot%\SysWoW64\regsvr32.exe """ & file & """"

'elevate privileges and register
With CreateObject("Shell.Application")
    .ShellExecute "cmd", args,, "runas"
End With

Function Expand(path)
    With CreateObject("WScript.Shell")
        Expand = .ExpandEnvironmentStrings(path)
    End With
End Function

Function Resolve(relativePath)
    With CreateObject("Scripting.FileSystemObject")
        Resolve = .GetAbsolutePathName(relativePath)
        If Not .FileExists(Resolve) Then
            Err.Raise 2,, "File doesn't exist: " & Resolve
        ElseIf Not "wsc" = LCase(.GetExtensionName(Resolve)) Then
            Err.Raise 3,, "Expected .wsc file. Actual file: " & Resolve
        End If
    End With
End Function
