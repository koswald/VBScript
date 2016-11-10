
'Setup the VBScript utilities

'Registers the dependency manager scriptlet, includer.wsc, if desired, runs the tests

With New VBSSetupUtility
    .Setup
End With

Class VBSSetupUtility

    Private sa, sh, fso, parent, includer, launcher

    Sub Class_Initialize

        includer = "class\includer.wsc"
        launcher = "examples\TestLauncher.vbs"

        Set sa = CreateObject("Shell.Application")
        Set sh = CreateObject("WScript.Shell")
        Set fso = CreateObject("Scripting.FileSystemObject")
        parent = fso.GetParentFolderName(WScript.ScriptFullName)
    End Sub

    Sub Setup

        'verify that this script was called from the batch file of
        'the same name, in order to ensure that it was started by
        'the 64-bit executable, if available, regardless of whether
        'the host machine opens .vbs files with the 64-bit .exe

        Dim msg : msg = "Please use Setup.bat to launch the Setup.vbs script"

        If 0 = WScript.Arguments.Count Then Err.Raise 1, WScript.ScriptName, msg
        If Not "Ensure_64-bit_executable" = WScript.Arguments(0) Then Err.Raise 1, WScript.ScriptName, msg

        'verify that we can find the scriptlet

        If Not fso.FileExists(parent & "\" & includer) Then
            Err.Raise 1, WScript.ScriptName, "Couldn't find the required scriptlet: " & includer
        End If

        'register the scriptlet for both x86 (32-bit) and x64 (64-bit)

        sa.ShellExecute "cmd", "/c regsvr32 /s """ & parent & "\" & includer & """ & %SystemRoot%\SysWow64\regsvr32 /s """ & parent & "\" & includer & """", "", "runas"

        'test the setup by running the tests, if desired

        s = "Do you want to run the tests?"

        If vbCancel = MsgBox(s, vbOKCancel + vbQuestion + vbSystemModal, WScript.ScriptName) Then Exit Sub

        sh.Run "%ComSpec% /k cscript.exe //nologo " & launcher

    End Sub

    Sub Class_Terminate
        Set sa = Nothing
        Set sh = Nothing
        Set fso = Nothing
    End Sub

End Class