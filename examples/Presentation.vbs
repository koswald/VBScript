'System tray icon with option to
'prevent the computer from going to sleep

Option Explicit
Const ALLOW_SLEEP = 0, PREVENT_SLEEP = 1 'menu indexes
Const largeIcon = True, smallIcon = False
Dim it : Set it = CreateObject("VBScripting.IdleTimer")
Dim ni : Set ni = CreateObject("VBScripting.NotifyIcon")
Dim sh : Set sh = CreateObject("WScript.Shell")
ni.AddMenuItem "Allow sleep", GetRef("AllowSleep")
ni.AddMenuItem "Prevent sleep", GetRef("PreventSleep")
ni.AddMenuItem "Exit", GetRef("CloseAndExit")
ni.Visible = True
AllowSleep
ListenForCallbacks

Sub ListenForCallbacks
    While True
        WScript.Sleep 200
    Wend
End Sub
Sub AllowSleep
    it.SystemRequired = False
    it.DisplayRequired = False
    ni.DisableMenuItem ALLOW_SLEEP
    ni.EnableMenuItem PREVENT_SLEEP
    ni.Text = "Presentation mode is off"
    ni.SetIconByDllFile "%SystemRoot%\System32\imageres.dll", 101, largeIcon
End Sub
Sub PreventSleep
    it.SystemRequired = True
    it.DisplayRequired = True
    ni.DisableMenuItem PREVENT_SLEEP
    ni.EnableMenuItem ALLOW_SLEEP
    ni.Text = "Presentation mode is on"
    ni.SetIconByDllFile "%SystemRoot%\System32\imageres.dll", 102, largeIcon
End Sub
Sub CloseAndExit
    it.Dispose
    Set it = Nothing
    ni.Dispose
    Set ni = Nothing
    Set sh = Nothing
    WScript.Quit
End Sub
