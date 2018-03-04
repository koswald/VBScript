
'test IdlePreventer.dll

Set ip = CreateObject("VBScripting.IdlePreventer")

Set ni = CreateObject("VBScripting.NotifyIcon")
ni.SetIconByDllFile "%SystemRoot%\System32\imageres.dll", 96, largeIcon
ni.Text = "Presentation mode - OFF" 'on-hover tooltip
ni.AddMenuItem "Allow idle", GetRef("AllowIdle")
ni.AddMenuItem "Prevent idle", GetRef("PreventIdle")
ni.AddMenuItem "Exit", GetRef("CloseAndExit")
Const ALLOW_IDLE = 0, PREVENT_IDLE = 1 'menu indexes
ni.DisableMenuItem ALLOW_IDLE
ni.Visible = True

ListenForCallbacks
Sub ListenForCallbacks
    While True
        WScript.Sleep 200
    Wend
End Sub

Sub AllowIdle
    ip.AllowIdle
    ni.DisableMenuItem ALLOW_IDLE
    ni.EnableMenuItem PREVENT_IDLE
    ni.Text = "Presentation mode OFF"
    ni.SetIconByDllFile "%SystemRoot%\System32\imageres.dll", 96, largeIcon
End Sub
Sub PreventIdle
    ip.PreventIdle
    ni.DisableMenuItem PREVENT_IDLE
    ni.EnableMenuItem ALLOW_IDLE
    ni.Text = "Presentation mode ON"
    ni.SetIconByDllFile "%SystemRoot%\System32\shell32.dll", 272, largeIcon
End Sub
Sub CloseAndExit
    ip.Dispose
    Set ip = Nothing
    ni.Dispose
    Set ni = Nothing
    WScript.Quit
End Sub

Dim ip, ni
Const largeIcon = True, smallIcon = False
