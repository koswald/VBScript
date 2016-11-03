
'Used to start .hta file of the same name:
'  1) to elevate privileges
'  2) to ensure that .hta is called with the 64-bit mshta.exe

With New Fixer : .Main : End With

Class Fixer

    Dim sh, sa, fso
    Dim args, msg, token

    Sub Class_Initialize
        token = "Ensure_x64_exe_starts_.hta_file"
        Set sh = CreateObject("WScript.Shell")
        Set sa = CreateObject("Shell.Application")
        Set fso = CreateObject("Scripting.FileSystemObject")
        msg = fso.GetBaseName(WScript.ScriptName) & ".hta must be started with the .bat file of the same name, in order to ensure that the 64-bit mshta.exe is used to start it."
    End Sub

    Sub Main

        'validate command-line argument

        If 0 = WScript.Arguments.Count Then QuitMessage("No arguments.") : Exit Sub
        If WScript.Arguments(0) <> token Then QuitMessage("Invalid token: " & WScript.Arguments(0)) : Exit Sub

        'start the .hta file with elevated privileges,
        'using the 64-bit mshta: this will work only if this script itself was started with a 64-bit wscript.exe or cscript.exe, hence the token
        'assume .hta file has the same name as this file
        'assume .hta file is in the same folder as this file

        dir = fso.GetParentFolderName(WScript.ScriptFullName)
        cmd = sh.ExpandEnvironmentStrings("%SystemRoot%\System32\mshta.exe")
        args = """" & dir & "\" & fso.GetBaseName(WScript.ScriptName) & ".hta"" " & WScript.Arguments(0)
        sa.ShellExecute cmd, args,, "runas"
    End Sub

    Sub QuitMessage(str)
        sh.PopUp "Error: " & str & vbLf & vbLf & msg, 30, WScript.ScriptName, vbInformation
    End Sub

    Sub Class_Terminate
        Set sa = Nothing
        Set sh = Nothing
        Set fso = Nothing
    End Sub

End Class

