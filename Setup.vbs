
'Setup the VBScript utilities

'Registers the dependency manager and, if desired, runs the tests

With New VBSSetupUtility
    .Setup
End With

Class VBSSetupUtility

    Private oSA, oSh, oFSO

    Sub Class_Initialize
        Set oSA = CreateObject("Shell.Application")
        Set oSH = CreateObject("WScript.Shell")
        Set oFSO = CreateObject("Scripting.FileSystemObject")
    End Sub

    Property Get sh : Set sh = oSh : End Property
    Property Get fso : Set fso = oFSO : End Property
    Property Get sa : Set sa = oSA : End Property

    Sub Setup

        'register the required scriptlet, includer.wsc, for dependency management

        Dim s : s = "Setup needs to register the VBScript dependency-management " & _
            "helper, includer.wsc, so the user account control dialog will open."

        If vbCancel = MsgBox(s, vbOKCancel + vbInformation, WScript.ScriptName) Then
            Exit Sub
        End If
        Dim thisFolder : thisFolder = fso.GetParentFolderName(WScript.ScriptFullName)
        Dim includerFile : includerFile = thisFolder & "\class\includer.wsc"
        If Not fso.FileExists(includerFile) Then 
            Err.Raise 1, WScript.ScriptName, "Couldn't find the includer file " & includerFile
        End If
        sa.ShellExecute "regsvr32", "/s " & includerFile, "", "runas"

        'test the setup by running the tests, if desired

        s = "Do you want to run the tests?"

        If vbCancel = MsgBox(s, vbOKCancel + vbQuestion, WScript.ScriptName) Then Exit Sub

        sh.Run "%ComSpec% /k cscript.exe //nologo examples\TestLauncher.vbs"

    End Sub

    Sub Class_Terminate 'fires when object instance goes out of scope
        Set oSA = Nothing
        Set oSh = Nothing
        Set oFSO = Nothing
    End Sub

End Class