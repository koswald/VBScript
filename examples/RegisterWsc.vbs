'Register a .wsc file for 64-bit and 32-bit apps.
'Use as a drop target or from the command line.
'Relative paths are supported.
'Environment variables are supported.

'Use /u to unregister

Option Explicit : Setup
With WScript.Arguments
    If 0 = .Count Then
        Err.Raise 1,, "Argument expected: the .wsc file to register."
    End If
    Dim i, item, file
    For i = 0 To .Count - 1
        If "/u" = LCase(.item(i)) Then
            uninstalling = True
            flags = "/u /n"
        Else
            item = Resolve(Expand(.item(i)))
            If fso.FileExists(item) Then file = item
        End If
    Next
    If IsEmpty(file) Then Err.Raise 2,, "None of the command line arguments specify an existing file."
End With

Dim args : args = format(Array("/c " & _
    "%SystemRoot%\System32\regsvr32.exe %s /s /i:""%s"" scrobj.dll & " & _
    "%SystemRoot%\SysWow64\regsvr32.exe %s /s /i:""%s"" scrobj.dll", _
    flags, file, flags, file _
  ))

With CreateObject("Shell.Application")
    .ShellExecute "cmd", args,, "runas"
End With

Teardown

Sub Teardown
    Set fso = Nothing
    Set sh = Nothing
    Set format = Nothing
End Sub

Dim sh, fso, format
Dim flags
Dim uninstalling

Sub Setup
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set sh = CreateObject("WScript.Shell")
    Set format = CreateObject("VBScripting.StringFormatter")
    flags = ""
    uninstalling = False
End Sub

Function Expand(path)
    Expand = sh.ExpandEnvironmentStrings(path)
End Function
Function Resolve(relativePath)
    Resolve = fso.GetAbsolutePathName(relativePath)
    If Not fso.FileExists(Resolve) Then
        Err.Raise 3,, "File doesn't exist: " & Resolve
    ElseIf Not "wsc" = LCase(fso.GetExtensionName(Resolve)) Then
        Err.Raise 4,, "Expected .wsc file. Actual file: " & Resolve
    End If
End Function
