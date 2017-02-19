
'drag a .cs file onto this script to compile a .dll (use .SetDLL) or .exe (class default)

'http://stackoverflow.com/questions/4198583/how-do-i-register-a-com-dll-written-in-c-sharp-with-regsvr32

With New NETCOMCompiler
    '.SetDebug
    '.SetNoWarn
    .SetDLL
    .Compile
End With

Class NETCOMCompiler

    Private fso, sh
    Private filespec, baseName, exeFolder, ext
    Private args, cmd, L, scriptName, msg

    Sub Class_Initialize
        args = ""
        exeFolder = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319"
        AddRef "C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.0\System.Speech.dll"
        ext = "exe" 'default file extension
        'args = args & " /nologo"
        'args = args & " /platform:x86"
        'args = args & " /platform:x64"
        'args = args & " /platform:anycpu" 'default
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set sh = CreateObject("WScript.Shell")
        scriptName = WScript.ScriptName
        L = vbLf & vbTab
    End Sub

    Sub Compile

        'validate command-line argument, exeFolder

        If 0 = WScript.Arguments.Count Then Err.Raise 1, scriptName, "A command-line argument is required:" & L & "the filespec of the .cs file to be compiled"
        filespec = WScript.Arguments(0)
        If Not fso.FileExists(filespec) Then Err.Raise 1, scriptName, "Couldn't find the file " & filespec
        If Not "cs" = LCase(fso.GetExtensionName(filespec)) Then Err.Raise 1, scriptName, "A .cs file is required."
        If Not fso.FolderExists(exeFolder) Then Err.Raise 1, scriptName, "Couldn't find the .NET executables folder, " & L & exeFolder & L & "Check the config file, " & configFile

        'build the commmand string

        baseName = fso.GetBaseName(filespec)
        cmd = "%ComSpec% /c cd """ & fso.GetParentFolderName(WScript.ScriptFullName) & """"
        cmd = cmd & " & echo."
        cmd = cmd & " & """ & exeFolder & "\csc"" /out:" & baseName & "." & ext & args & " """ & filespec & """"
        cmd = cmd & " & echo. & echo OK to ignore warning CS1699 and BC41008. & echo."
        cmd = cmd & " & pause"

        'give an opt out

        If Len(cmd) > 254 Then
            'string length exceeds InputBox limit, so use MsgBox
            msg = "Verify command"
            If vbCancel = MsgBox(cmd, vbOKCancel, msg & " - " & WScript.ScriptName) Then Exit Sub
        Else
            msg = "Verify/modify command"
            cmd = InputBox(msg, scriptName, cmd)
            If "" = cmd Then Exit Sub
        End If

        'run the command

        sh.Run cmd
    End Sub

    Sub SetNoWarn : args = args & " /warn:0" : End Sub
    Sub SetDebug : args = args & " /debug" : End Sub
    Sub AddRef(ref) : args = args & " /r:""" & ref & """" : End Sub


    Sub SetDLL
        args = args & " /target:library"
        ext = "dll"
    End Sub

    Sub Class_Terminate
        Set fso = Nothing
        Set sh = Nothing
    End Sub
End Class
