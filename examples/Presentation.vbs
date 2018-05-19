Option Explicit
purpose = "Show a system tray icon with options to prevent the computer and/or monitor from going to sleep."
helpMessage = "When presentation mode is on, the computer and monitor are typically prevented from going into a suspend (sleep) state or hibernation. The computer may still be put to sleep by other applications or by user actions such as closing a laptop lid or pressing a sleep button or power button." & vbLf & vbLf & "Phone charger mode is the same as presentation mode except that the monitor is turned off, initially."
Call Setup

Sub NormalMode
    watcher.Watch = False
    notifyIcon.DisableMenuItem normalModeMenuIndex
    notifyIcon.EnableMenuItem presentationModeMenuIndex
    notifyIcon.SetIconByDllFile icon(ico_normFile), icon(ico_normIndex), icon(ico_normType)
    notifyIcon.Text = "Presentation mode is off"
    PublishStatus "Normal"
    csTimer.Stop
End Sub
Sub PresentationMode
    watcher.Watch = True
    notifyIcon.EnableMenuItem normalModeMenuIndex
    notifyIcon.DisableMenuItem presentationModeMenuIndex
    notifyIcon.SetIconByDllFile icon(ico_presentFile), icon(ico_presentIndex), icon(ico_presentType)
    PublishStatus "Presentation"
    stopwatch.Reset
    csTimer.Start
End Sub
Sub ChargerMode
    shell.Run "rundll32 user32.dll,LockWorkStation",, synchronous
    WScript.Sleep 1000
    watcher.MonitorOff
    WScript.Sleep 1000
    PresentationMode
End Sub
Sub PublishStatus(newStatus)
    status = newStatus
    Dim stream : Set stream = fso.OpenTextFile(statusFile, ForWriting, CreateNew)
    stream.WriteLine newStatus
    stream.Close
    Set stream = Nothing
End Sub
Sub SetDurationUI
    sa.MinimizeAll
    Dim currentValue : currentValue = Round(csTimer.IntervalInHours, 4)
    Dim response : response = InputBox(format(Array("Enter the desired duration of Presentation mode / Phone charger mode, in hours.%sCurrent value: %s", vbLf & vbLf, currentValue)), WScript.ScriptName, currentValue)
    sa.UndoMinimizeAll
    If "" = response Then Exit Sub
    csTimer.IntervalInHours =  response
    If "Presentation" = status Then
        PresentationMode 'reset timers
    End If
End Sub
Sub Help
    shell.PopUp helpMessage, 80, WScript.ScriptName, vbInformation + vbSystemModal
End Sub
Sub ListenForCallbacks
    While True
        If "Presentation" = status Then
            notifyIcon.Text = format(Array("Presentation mode is on%sNormal mode resumes in %s min.", vbLf, Round(csTimer.Interval/60000 - stopwatch/60, 0)))
        End If
        WScript.Sleep 200
    Wend
End Sub

'icon options
Const icon1 = "%SystemRoot%\System32\powercpl.dll|5|False|%SystemRoot%\System32\powercpl.dll|6|False"
Const icon2 = "%SystemRoot%\System32\imageres.dll|101|False|%SystemRoot%\System32\imageres.dll|102|False"
Const icon3 = "%SystemRoot%\System32\imageres.dll|96|True|%SystemRoot%\System32\deskmon.dll|0|True"
Const icon4 = "%SystemRoot%\System32\hgcpl.dll|1|False|%SystemRoot%\System32\hgcpl.dll|0|False"
Const icon5 = "%SystemRoot%\System32\DDORes.dll|19|False|%SystemRoot%\System32\DDORes.dll|15|False"
Const icon6 = "%SystemRoot%\System32\DDORes.dll|19|True|%SystemRoot%\System32\DDORes.dll|15|True"
Const ico_normFile = 0, ico_normIndex = 1, ico_normType = 2, ico_presentFile = 3, ico_presentIndex = 4, ico_presentType = 5

Const synchronous = True
Const largeIcon = True, smallIcon = False
Const PresentationState = 3, NormalState = 0
Const ForWriting = 2, CreateNew = True
Dim watcher, notifyIcon, shell, fso, csTimer, sa, stopwatch, includer, format
Dim normalModeMenuIndex, presentationModeMenuIndex
Dim purpose, helpUrl, helpMessage
Dim statusFile, status
Dim icon

Sub Setup
    Set shell = CreateObject("WScript.Shell")
    Dim dataFolder : dataFolder = shell.ExpandEnvironmentStrings("%AppData%\VBScripting")
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(dataFolder) Then fso.CreateFolder dataFolder
    Set format = CreateObject("VBScripting.StringFormatter")
    statusFile = format(Array("%s\%s.status", dataFolder, fso.GetBaseName(WScript.ScriptName)))

    Set notifyIcon = CreateObject("VBScripting.NotifyIcon")
    notifyIcon.AddMenuItem "Normal mode", GetRef("NormalMode")
        normalModeMenuIndex = 0
    notifyIcon.AddMenuItem "Presentation mode", GetRef("PresentationMode")
        presentationModeMenuIndex = 1
    notifyIcon.AddMenuItem "Phone charger mode", GetRef("ChargerMode")
    notifyIcon.AddMenuItem "Set duration", GetRef("SetDurationUI")
    notifyIcon.AddMenuItem "Help", GetRef("Help")
    notifyIcon.AddMenuItem "Exit", GetRef("CloseAndExit")
    notifyIcon.Visible = True

    Set csTimer = CreateObject("VBScripting.Timer")
    csTimer.IntervalInHours = 1.5
    csTimer.AutoReset = False
    Set csTimer.Callback = GetRef("NormalMode")

    Set watcher = CreateObject("VBScripting.Watcher")
    Set sa = CreateObject("Shell.Application")
    Set includer = CreateObject("VBScripting.Includer")
    Execute includer.Read("VBSStopwatch")
    Set stopwatch = New VBSStopwatch

    icon = Split(icon1, "|")
    NormalMode
    ListenForCallbacks
End Sub
Sub CloseAndExit
    PublishStatus "Normal"
    watcher.Dispose
    Set watcher = Nothing
    notifyIcon.Dispose
    Set notifyIcon = Nothing
    csTimer.Dispose
    Set csTimer = Nothing
    Set shell = Nothing
    Set fso = Nothing
    Set sa = Nothing
    Set includer = Nothing
    Set stopwatch = Nothing
    WScript.Quit
End Sub
