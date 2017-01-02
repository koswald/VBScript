
'Power functions: shutdown, restart, logoff, sleep, and hibernate.

Class VBSPower

    Private debug, sh, sa
    Private force

    Sub Class_Initialize
        SetDebug False
        SetForce False
        Set sh = CreateObject("WScript.Shell")
        Set sa = CreateObject("Shell.Application")
    End Sub

    'Property Shutdown
    'Returns a boolean
    'Remark: Shuts down the computer. Returns True if the operation completes with no errors.
    Property Get Shutdown
        Shutdown = WMIPowerAction(ACTION_SHUTDOWN, force)
    End Property

    'Property Restart
    'Returns a boolean
    'Remark: Restarts the computer. Returns True if the operation completes with no errors.
    Property Get Restart
        Restart = WMIPowerAction(ACTION_RESTART, force)
    End Property

    'Property Logoff
    'Returns a boolean
    'Remark: Logs off the computer. Returns True if the operation completes with no errors.
    Property Get Logoff
        Logoff = WMIPowerAction(ACTION_LOGOFF, force)
    End Property

    'Private Function WMIPowerAction
    'Parameters: action, force
    'Returns a boolean
    'Remark: Uses WMI (Windows Management Instumentation) to perform the specified power action on the computer: see action constants. Forces the action, or not, as specified by the second parameter, except Windows 10 always forces. Returns True if the operation completes with no errors.
    Private Function WMIPowerAction(ByVal action, force)
        If debug Then Exit Function
        Dim OS
        On Error Resume Next
            If CBool(force) Then action = action + 4
            For Each OS in GetObject("winmgmts:{(Shutdown)}").ExecQuery("Select * from Win32_OperatingSystem")
                OS.Win32Shutdown(action)
            Next
            WMIPowerAction = Not CBool(Err) 'Returns False on error
        On Error Goto 0
    End Function

    Property Get ACTION_LOGOFF : ACTION_LOGOFF = 0 : End Property
    Property Get ACTION_RESTART : ACTION_RESTART = 2 : End Property
    Property Get ACTION_SHUTDOWN : ACTION_SHUTDOWN = 8 : End Property

    'Method Sleep
    'Remark: Puts the computer to sleep. Requires <a href="https://technet.microsoft.com/en-us/sysinternals/psshutdown.aspx"> PsShutdown.exe</a> from Windows Sysinternals to be located somewhere on your %Path%. Recovery from sleep is faster than from hibernation, but uses more power.
    Sub Sleep
        If debug Then Exit Sub
        sh.Run "psshutdown -d -t 0"
    End Sub

    'Method Hibernate
    'Remark: Puts the computer into hibernation. Will not work if hibernate is disabled in the Control Panel, in which case the EnableHibernation method may be used to reenable hibernation. Hibernate is more power-efficient than sleep, but recovery is slower. If the computer wakes after pressing a key or moving the mouse, then it was sleeping, not in hibernation. Recovery from hibernation typically requires pressing the power button.
    Sub Hibernate
        If debug Then Exit Sub
        sh.Run "%SystemRoot%\System32\rundll32.exe powrprof.dll, SetSuspendState 0,1,0"
    End Sub

    'Method EnableHibernation
    'Remark: Enables hibernation. The User Account Control dialog will open to request elevated privileges.
    Sub EnableHibernation : SetHibernationEnabled True : End Sub

    'Method DisableHibernation
    'Remark: Disables hibernation. The User Account Control dialog will open to request elevated privileges.
    Sub DisableHibernation : SetHibernationEnabled False : End Sub

    'Private Method SetHibernationEnabled
    Private Sub SetHibernationEnabled(ByVal enabled)
        If debug Then Exit Sub
        If enabled Then enabled = "on" Else enabled = "off"
        sa.ShellExecute "cmd", "/c powercfg.exe /hibernate " & enabled,, "runas"

        'when the Hibernate method is called immediately following,
        'some kind of pause is necessary;
        'ideally there would be a way to make the above command
        'synchronous, and there is a way, but it's too complicated
        WScript.Sleep 3000
    End Sub

    'Method SetForce
    'Parameter: force
    'Remark: Optional. Setting this to True forces the Shutdown or Restart, discarding unsaved work. Default is False. Logoff always forces apps to close.
    Sub SetForce(newForce) : force = newForce : End Sub

    'Method SetDebug
    'Parameter: a boolean
    'Remark: Used for testing. Prevents the computer from actually shutting down, etc., during testing. Default is False.
    Sub SetDebug(newDebug) : debug = newDebug : End Sub

    Sub Class_Terminate
        Set sh = Nothing
        Set sa = Nothing
    End Sub
End Class
