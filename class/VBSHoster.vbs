' Manage which script host is hosting the currently running script: cscript.exe or wscript.exe.
'
' Not suitable for .hta scripts. For .hta scripts, use the VBSApp class.
'
' If Windows Terminal is installed, a suggested setting in %LocalAppData%\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json: <code>"windowingBehavior": "useAnyExisting"</code> or <code>"windowingBehavior": "useExisting"</code>. The same setting in the Windows Terminal GUI: Settings &#124; Startup &#124; New instance behavior &#124; Attach to the most recently used window (or Attach to the most recently used window on this desktop). This applies to the RestartWith method's default behavior. The RestartWith method is used by the TestingFramework class when a test file is double-clicked in Windows Explorer.
'
Class VBSHoster

    Private args 'VBSArguments object
    Private format 'StringFormatter object
    Private sh 'WScript.Shell object
    Private sa 'Shell.Application object
    Private fso 'Scripting.FileSystemObject object
    Private parent 'string: parent folder of the calling script
    Private switch_ 'string: /c or /k or the powershell equivalent; contains the value of the Switch property.
    Private shell_
    Private methodExistsTest_
    Private powershell 'pwsh.exe filespec, if available,  or "powershell"

    Private Sub Class_Initialize
        Dim incl 'VBScripting.Includer object

        Set incl = CreateObject( "VBScripting.Includer" )
        Execute incl.Read( "VBSArguments" )
        Set args = New VBSArguments
        Set sh = CreateObject( "WScript.Shell" )
        Set sa = CreateObject( "Shell.Application" )
        Set fso = CreateObject( "Scripting.FileSystemObject" )
        Set format = CreateObject( "VBScripting.StringFormatter" )
        parent = fso.GetParentFolderName( WScript.ScriptFullName )
        SetSwitch "/c"
        MethodExistsTest = False

        Execute incl.Read( "Configurer" )
        With New Configurer

            'Prepare the default shell: use pwsh.exe filespec, or if it is not installed, fallback to "powershell". Use Windows Terminal if installed.
            powershell = .PowerShell
            If IsEmpty( .WT ) Then
                Shell = ""
            Else Shell = "wt -d """ & parent & """ "
            End If
            If .PsFallback = powershell Then
                Shell = Shell & .PsFallback
            Else Shell = Shell & """" & powershell & """"
            End If

            'If specified, a shell key in VBSHoster.configure takes precedence over the default shell.
            .LoadClassConfig me
            If .Exists( "shell" ) Then
                Shell = .Item( "shell" )
            End If

        End With
    End Sub

    'Method EnsureCScriptHost
    'Remark: Restart the script hosted with cscript.exe if it isn't already hosted with cscript.exe.
    Sub EnsureCScriptHost
        If MethodExistsTest Then Exit Sub
        If Not "cscript.exe" = LCase(Right(WScript.FullName,11)) Then
            SetSwitch "/k"
            RestartWith( "cscript.exe" )
        End If
    End Sub

    'Method SetSwitch
    'Parameter: /k or /c
    'Remark: Optional. Specifies a switch for %ComSpec% for use with the EnsureCScriptHost method: controls whether the command window, if newly created, remains open (/k). Useful for troubleshooting, in order to be able to read error messages. Unnecessary if starting the script from a console window, because /c is the default. If pwsh or powershell (or wt pwsh, etc.) is the Shell, then the equivalent string is substituted.
    Sub SetSwitch( newSwitch )
        If IsEmpty( newSwitch ) Then Exit Sub
        Switch = newSwitch
    End Sub
    Property Let Switch( newSwitch )
        switch_ = LCase( newSwitch )
        OnShellChange
    End Property
    Property Get Switch
        Switch = switch_
    End Property
    Private Sub OnShellChange
        Dim IsPwsh, IsPowershell 'booleans
        IsPwsh = CBool( InStr( LCase( Shell ), "pwsh" ))
        IsPowershell = CBool( InStr( LCase( Shell ), "powershell" ))
        If IsPwsh Or IsPowershell Then
            If "/k" = Switch Then
                switch_ = "-NoExit -Command"
            ElseIf "/c" = Switch Then
                switch_ = "-Command"
            End If
        End If
    End Sub

    'Method SetDefaultHostWScript
    'Remark: Sets wscript.exe to be the default script host. If privileges are not already elevated, then the User Account Control dialog will open for permission to elevate privileges.
    Sub SetDefaultHostWScript
        If MethodExistsTest Then Exit Sub
        sa.ShellExecute "wscript.exe", "//h:wscript", "", "runas"
    End Sub

    'Method SetDefaultHostCScript
    'Remark: Sets cscript.exe to be the default script host. If privileges are not already elevated, then the User Account Control dialog will open for permission to elevate privileges.
    Sub SetDefaultHostCScript
        If MethodExistsTest Then Exit Sub
        sa.ShellExecute "wscript.exe", "//h:cscript", "", "runas"
    End Sub

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

    'get the command-line arguments
    Private Property Get GetArgsString
        GetArgsString = args.GetArgumentsString
    End Property

    Property Get RestartCommand(host)
        RestartCommand = format (Array( _
            "%s %s %s //nologo ""%s"" %s", _
            Shell, Switch, host, WScript.ScriptFullName, GetArgsString _
        ))
    End Property

    'Method RestartWith
    'Parameter: a string: the host .exe
    'Remarks: Restarts the .vbs or .wsf script with the specified host, "cscript.exe" or "wscript.exe". By default, Windows Terminal will be used, if available. Also by default, pwsh.exe (PowerShell) will be used if available. A custom or unusual pwsh.exe install path can be specified if necessary in the file <code>.configure</code> in the project root folder. Use <code> class/VBSHoster.configure</code> to specify another shell configuration. <br /> Examples:<br /><code>shell, cmd</code><br /><code>shell, powershell</code><br /><code>shell, pwsh</code><br /><code>shell, wt cmd</code><br /><code>shell, wt pwsh</code><br /><code>shell, wt "%ProgramFilesX86%\PowerShell\7\pwsh.exe"</code><br /><code>shell, %ProgramFilesX86%\PowerShell\7\pwsh.exe</code><br />This setting can be overridden by the Shell property. See also the RestartUsing method of the <a href="#vbsapp"> VBSApp class</a>.
    Sub RestartWith(host)
        If IsEmpty(host) Then Exit Sub
        sh.CurrentDirectory = parent
        sh.Run RestartCommand(host)
        WScript.Quit
    End Sub

    'Property Shell
    'Parameter: a string
    'Returns: a string
    'Remarks: Gets or sets the shell used when restarting a script (see the RestartWith method). Examples: cmd, powershell, pwsh, wt pwsh. Overrides the shell read from <code>VBSHoster.configure</code>.
    Public Property Let Shell( newValue )
        shell_ = newValue
        OnShellChange
    End Property
    Public Property Get Shell
        Shell = shell_
    End Property

    Property Let MethodExistsTest( newValue )
        methodExistsTest_ = newValue
    End Property
    Property Get MethodExistsTest
        MethodExistsTest = methodExistsTest_
        methodExistsTest_ = False
    End Property

    Sub Class_Terminate
        Set sh = Nothing
        Set sa = Nothing
        Set fso = Nothing
    End Sub

End Class
