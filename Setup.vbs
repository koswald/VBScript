
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

        'verify that we can find the scriptlet

        If Not fso.FileExists(parent & "\" & includer) Then
            Err.Raise 1, WScript.ScriptName, "Couldn't find the required scriptlet: " & includer
        End If

        'register the scriptlet

        sa.ShellExecute "regsvr32", "/s " & parent & "\" & includer, "", "runas"

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