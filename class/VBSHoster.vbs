'Manage which script host is hosting the currently running script
'
Class VBSHoster

    'Method EnsureCScriptHost
    'Remark: Restart the script hosted with CScript if it isn't already hosted with CScript.exe
    Sub EnsureCScriptHost
        If Not "cscript.exe" = LCase(Right(WScript.FullName,11)) Then
            SetSwitch "/k"
            RestartWith("cscript.exe")
        End If
    End Sub

    'Method SetSwitch
    'Parameter: /k or /c
    'Remark: Optional. Specifies a switch for %ComSpec% for use with the EnsureCScriptHost method: controls whether the command window, if newly created, remains open (/k). Useful for troubleshooting, in order to be able to read error messages. Unnecessary if starting the script from a console window, because /c is the default.
    Sub SetSwitch(newSwitch)
        switch = newSwitch
    End Sub

    'Method SetDefaultHostWScript
    'Remark: Sets wscript.exe to be the default script host. The User Account Control dialog will open for permission to elevate privileges.
    Sub SetDefaultHostWScript : sa.ShellExecute "wscript.exe", "//h:wscript", "", "runas" : End Sub

    'Method SetDefaultHostCScript
    'Remark: Sets cscript.exe to be the default script host. The User Account Control dialog will open for permission to elevate privileges.
    Sub SetDefaultHostCScript : sa.ShellExecute "wscript.exe", "//h:cscript", "", "runas" : End Sub

    'Property GetDefaultHost
    'Returns: a string
    'Remark: Returns "wscript.exe" or "cscript.exe", according to which .exe opens .vbs files by default.
    Property Get GetDefaultHost
        If "Open" = sh.RegRead(DefaultHostKey) Then
            GetDefaultHost = "wscript.exe"
        ElseIf "Open2" = sh.RegRead(DefaultHostKey) Then
            GetDefaultHost = "cscript.exe"
        Else GetDefaultHost = "Default host could not be determined."
        End If
    End Property

    'current default host key
    'if value is Open2 then current host is cscript; if Open then wscript
    'intended for use with the RegRead method of WScript.Shell
    Property Get DefaultHostKey
        DefaultHostKey = "HKLM\SOFTWARE\Classes\VBSFile\Shell\"
    End Property

    Private args
    Private sh, sa
    Private switch
    Private format

    Private Sub Class_Initialize 'event fires on object instantiation
        With CreateObject("VBScripting.Includer")
            Execute .read("VBSArguments")
            Execute .read("StringFormatter")
        End With
        Set args = New VBSArguments
        Set format = New StringFormatter
        Set sh = CreateObject("WScript.Shell")
        Set sa = CreateObject("Shell.Application")
        SetSwitch "/c"
    End Sub

    'get the command-line arguments
    Private Property Get GetArgsString
        GetArgsString = args.GetArgumentsString
    End Property

    'restart the script with the specified host, preserving the arguments
    Private Sub RestartWith(host)
        sh.Run format (Array( _
            "%ComSpec% %s %s //nologo ""%s"" %s", _
            switch, host, WScript.ScriptName, GetArgsString _
        ))
        WScript.Quit
    End Sub

    Sub Class_Terminate
        Set sh = Nothing
        Set sa = Nothing
    End Sub
End Class
