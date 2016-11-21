
'start .hta file of the same name
'  1) with elevated privileges
'  2) with the 64-bit mshta.exe

With New Fixer : .Main : End With

Class Fixer
    Dim sa, fso
    Dim args, msg, token

    Sub Class_Initialize
        token = "Ensure_x64_exe_starts_.hta_file"
        Set sa = CreateObject("Shell.Application")
        Set fso = CreateObject("Scripting.FileSystemObject")
        msg = fso.GetBaseName(WScript.ScriptName) & ".hta must be started with the .bat file of the same name, in order to ensure that the 64-bit mshta.exe is used."
    End Sub

    Sub Main
        'validate the command-line argument

        If 0 = WScript.Arguments.Count Then Err.Raise 1, WScript.ScriptName, msg
        If WScript.Arguments(0) <> token Then Err.Raise 1, WScript.ScriptName, msg

        'start the .hta file with elevated privileges, using the 64-bit mshta.exe; this will work only if this script itself was started with a 64-bit wscript.exe or cscript.exe, hence the token
        'assume .hta file has the same name as this file and is in the same folder as this file

        dir = fso.GetParentFolderName(WScript.ScriptFullName)
        args = """" & dir & "\" & fso.GetBaseName(WScript.ScriptName) & ".hta"" " & WScript.Arguments(0)
        sa.ShellExecute "mshta", args,, "runas"
    End Sub

    Sub Class_Terminate
        Set sa = Nothing
        Set fso = Nothing
    End Sub
End Class

