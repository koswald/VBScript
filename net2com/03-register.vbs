
'register or unregister the .dll by dragging it onto this script

'http://stackoverflow.com/questions/4198583/how-do-i-register-a-com-dll-written-in-c-sharp-with-regsvr32

With New DLLRegistrar
    '.SetBitness 32 'default = 64
    .Register '.Register or .unRegister
End With

Class DLLRegistrar

    Private fso, sa
    Private exeFolder, exeFolder64, exeFolder32
    Private filespec, baseName, scriptName
    Private args, unregisterString, L, msg

    Sub Class_Initialize
        exeFolder64 = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319"
        exeFolder32 = "C:\Windows\Microsoft.NET\Framework\v4.0.30319"
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set sa = CreateObject("Shell.Application")
        scriptName = WScript.ScriptName
        unregisterString = ""
        L = vbLf & vbTab
        SetBitness 64
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
        args = args & " & """ & exeFolder & "\RegAsm.exe"""
        'args = args & " /tlb:" & baseName & ".tlb" 'create a type library
        args = args & " /codebase" 'if not putting .dll in the GAC
        args = args & " """ & filespec & """"
        args = args & unregisterString
        args = args & " & echo. & pause "

        'give an opt out

        If Len(args) > 254 Then
            'string length exceeds InputBox limit, so use MsgBox
            msg = "Verify arguments"
            If vbCancel = MsgBox(args, vbOKCancel, msg & " - " & WScript.ScriptName) Then Exit Sub
        Else
            'use InputBox: give an oportunity to modify
            msg = "Verify/modify arguments"
            args = InputBox(msg, scriptName, args)
            If "" = args Then Exit Sub 'user clicked Cancel or pressed Esc
        End If

        'run the command with elevated privileges

        sa.ShellExecute "cmd", args,, "runas"
    End Sub

    Sub SetBitness(bitness)
        If 64 = bitness Then
            exeFolder = exeFolder64
        ElseIf 32 = bitness Then
            exeFolder = exeFolder32
        End If
    End Sub

    Sub Class_Terminate
        Set fso = Nothing
        Set sa = Nothing
    End Sub
End Class
