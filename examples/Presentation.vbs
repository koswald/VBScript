'System tray icon with option to
'prevent the computer from going to sleep

Option Explicit
Dim ir, ni, sh
Set sh = CreateObject("WScript.Shell")
Const largeIcon = True, smallIcon = False
Const ALLOW_IDLE = 0, PREVENT_IDLE = 1 'menu indexes

Set ir = CreateObject("VBScripting.IdleTimer")

Set ni = CreateObject("VBScripting.NotifyIcon")
ni.SetIconByDllFile "%SystemRoot%\System32\imageres.dll", 101, largeIcon
ni.Text = "Presentation mode is off" 'on-hover tooltip
ni.AddMenuItem "Allow sleep", GetRef("AllowSleep")
ni.AddMenuItem "Prevent sleep", GetRef("PreventSleep")
ni.AddMenuItem "Sleep now", GetRef("SleepNow")
ni.AddMenuItem "Lock workstation", GetRef("LockWorkstation")
ni.AddMenuItem "Edit script", GetRef("EditAndRestartScript")
ni.AddMenuItem "Exit", GetRef("CloseAndExit")
ni.DisableMenuItem ALLOW_IDLE
ni.Visible = True

ListenForCallbacks
Sub ListenForCallbacks
    While True
        WScript.Sleep 200
    Wend
End Sub

Sub AllowSleep
    ir.AllowSleep
    ni.DisableMenuItem ALLOW_IDLE
    ni.EnableMenuItem PREVENT_IDLE
    ni.Text = "Presentation mode is off"
    ni.SetIconByDllFile "%SystemRoot%\System32\imageres.dll", 101, largeIcon
End Sub
Sub PreventSleep
    ir.PreventSleep
    ni.DisableMenuItem PREVENT_IDLE
    ni.EnableMenuItem ALLOW_IDLE
    ni.Text = "Presentation mode is on"
    ni.SetIconByDllFile "%SystemRoot%\System32\imageres.dll", 102, largeIcon
End Sub
Sub SleepNow
    With CreateObject("VBScripting.VBSPower")
        .Sleep
    End With
End Sub
Sub LockWorkstation
    sh.Run "%SystemRoot%\System32\rundll32.exe user32.dll,LockWorkStation"
End Sub
Sub EditAndRestartScript
    Const IN_USE = 0
    Dim editor : Set editor = sh.Exec("notepad """ & WScript.ScriptFullName & """")
    While editor.Status = IN_USE
        WScript.Sleep 200
    Wend
    sh.Run """" & WScript.ScriptFullName & """"
    CloseAndExit
End Sub
Sub CloseAndExit
    ir.Dispose
    Set ir = Nothing
    ni.Dispose
    Set ni = Nothing
    Set sh = Nothing
    WScript.Quit
End Sub
