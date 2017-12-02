
'Setup the VBScript utilities

'Registers the dependency manager scriptlet, includer.wsc, 
'and if desired, launches the standard tests

Option Explicit
Const includer = "class\includer.wsc" 'scriptlet relative path
Const launcher = "examples\test launchers\TestLauncherStandard.vbs"
With New VBSSetupUtility
    .Setup
End With
Class VBSSetupUtility
    Sub Setup

        'verify that we can find the scriptlet
        If Not fso.FileExists(scriptlet) Then
            Err.Raise 1,, "Couldn't find the required scriptlet: " & scriptlet
        End If

        'register the scriptlet for both x86 (32-bit) and x64 (64-bit)
        Dim args : args = "/c " & _
            "%SystemRoot%\System32\regsvr32 /s """ & scriptlet & """ & " & _
            "%SystemRoot%\SysWow64\regsvr32 /s """ & scriptlet & """"
        sa.ShellExecute "cmd", args,, "runas" 'elevate privileges

        'test the setup by running the tests, if desired
        Dim msg : msg = "Setup is finished." & vbLf & vbLf & "Before closing, Setup can run the standard tests, which may take about 30 seconds."
        Dim mode : mode = vbOKCancel + vbInformation + vbSystemModal
        Dim caption : caption = WScript.ScriptName
        If vbCancel = MsgBox(msg, mode, caption) Then Exit Sub
        sh.Run "%ComSpec% /k cscript.exe //nologo """ & launcher & """"
    End Sub

    Private scriptlet
    Private sa, sh, fso 'objects
    Private parent 'absolute, resolved folder spec

    Sub Class_Initialize
        Set sa = CreateObject("Shell.Application")
        Set sh = CreateObject("WScript.Shell")
        Set fso = CreateObject("Scripting.FileSystemObject")
        parent = fso.GetParentFolderName(WScript.ScriptFullName)
        scriptlet = fso.GetAbsolutePathName(parent & "\" & includer)
    End Sub

    Sub Class_Terminate
        Set sa = Nothing
        Set sh = Nothing
        Set fso = Nothing
    End Sub
End Class