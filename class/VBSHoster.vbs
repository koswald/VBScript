
'Manage which script host is hosting the currently running script

Class VBSHoster

    Private oVBSNatives, oVBSArguments, switch

    Private Sub Class_Initialize 'event fires on object instantiation
        With CreateObject("includer")
            Execute(.read("VBSNatives"))
            Execute(.read("VBSArguments"))
        End With

        Set oVBSNatives = New VBSNatives
        Set oVBSArguments = New VBSArguments

        SetSwitch "/c"
    End Sub

    Property Get n : Set n = oVBSNatives : End Property

    'if restarting the script, you likely want to pass along the original arguments...

    Private Property Get GetArgsString
        GetArgsString = oVBSArguments.GetArgumentsString
    End Property

    'Method EnsureCScriptHost
    'Remark: Restart the script hosted with CScript if it isn't already hosted with CScript.exe

    Sub EnsureCScriptHost
        If Not "cscript.exe" = LCase(Right(WScript.FullName,11)) Then
            RestartWith("cscript.exe")
        End If
    End Sub

    'restart the script with the specified host

    Private Sub RestartWith(host)
        Select Case LCase(host)
        Case "cscript.exe"

            'notify the user that something kinda weird is going on and how to fix it

            MsgBox "This should work, but ideally, " _
                & WScript.ScriptName & " should be started from a " _
                & "command window with cscript." & vbLf & vbLf _
                & "E.g. cscript //nologo " & WScript.ScriptName _
                , vbInformation + vbSystemModal _
                , WScript.ScriptName
        End Select

        n.sh.Run "%ComSpec% " & switch & " " & host & " //nologo " & WScript.ScriptName & GetArgsString
        WScript.Quit
    End Sub

    'Method SetSwitch
    'Parameter: /k or /c
    'Remark: Optional. Specifies a switch for %ComSpec% for use with the EnsureCScriptHost method: controls whether the command window, if newly created, remains open (/k). Useful for troubleshooting, in order to be able to read error messages. Unnecessary if starting the script from a console window, because /c is the default.

    Sub SetSwitch(newSwitch)
        switch = newSwitch
    End Sub

    'Method SetDefaultHostWScript
    'Remark: Sets wscript.exe to be the default script host. The User Account Control dialog will open for permission to elevate privileges.

    Sub SetDefaultHostWScript : n.sa.ShellExecute "wscript.exe", "//h:wscript", "", "runas" : End Sub

    'Method SetDefaultHostCScript
    'Remark: Sets cscript.exe to be the default script host. The User Account Control dialog will open for permission to elevate privileges.

    Sub SetDefaultHostCScript : n.sa.ShellExecute "wscript.exe", "//h:cscript", "", "runas" : End Sub

End Class
