
'register or unregister the .dll by dragging it onto this script

With New DLLRegistrar
    .Register '.Register or .unRegister
End With

Class DLLRegistrar

    Private fso, sa
    Private exeFolder, filespec, baseName, scriptName
    Private args, unregisterString, L, msg

    Sub Class_Initialize
        exeFolder = "C:\Windows\Microsoft.NET\Framework\v4.0.30319"
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set sa = CreateObject("Shell.Application")
        scriptName = WScript.ScriptName
        unregisterString = ""
        L = vbLf & vbTab
    End Sub

    Sub Unregister
        unregisterString = " /unregister"
        Register
    End Sub

    Sub Register

        'validate the command line argument, exeFolder

        If 0 = WScript.Arguments.Count Then Err.Raise 1, scriptName, "A command line argument is required:" & L & "The filespec of a .dll file"
        filespec = WScript.Arguments(0)
        If Not "dll" = LCase(fso.GetExtensionName(filespec)) Then Err.Raise 1, scriptName, "The file type is incorrect:" & L & "A .dll file is required."
        baseName = fso.GetBaseName(filespec)
        If Not fso.FolderExists(exeFolder) Then Err.Raise 1, scriptName, "Couldn't find the .NET executables folder," & L & exeFolder

        'build the argument(s)

        args = "/c cd """ & fso.GetParentFolderName(WScript.ScriptFullName) & """"
        args = args & " & echo. "
        args = args & " & """ & exeFolder & "\RegAsm"""
        'args = args & " /tlb:" & baseName & ".tlb" 'create a type library
        args = args & " /codebase" 'if not putting .dll in the GAC
        args = args & " """ & filespec & """"
        args = args & unregisterString
        args = args & " & echo. & pause "

        'give an opt out

        If Len(cmd) > 254 Then
            'string length exceeds InputBox limit, so use MsgBox
            msg = "Verify arguments"
            If vbCancel = MsgBox(args, vbOKCancel, msg & " - " & WScript.ScriptName) Then Exit Sub
        Else
            msg = "Verify/modify arguments"
            If "" = InputBox(msg, scriptName, args) Then Exit Sub
        End If

        'run the command with elevated privileges

        sa.ShellExecute "cmd", args,, "runas"
    End Sub

    Sub Class_Terminate
        Set fso = Nothing
        Set sa = Nothing
    End Sub
End Class
