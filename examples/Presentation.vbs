Option Explicit
purpose = "Show a system tray icon with options to prevent the computer and/or monitor from going to sleep."
helpMessage = "When presentation mode is on, the computer and monitor are typically prevented from going into a suspend (sleep) state or hibernation. The computer may still be put to sleep by other applications or by user actions such as closing a laptop lid or pressing a sleep button or power button."
Call Main

Sub NormalMode
    watcher.Watch = False
    notifyIcon.DisableMenuItem normalModeMenuIndex
    notifyIcon.EnableMenuItem presentationModeMenuIndex
    notifyIcon.SetIconByDllFile "%SystemRoot%\System32\powercpl.dll", 5, largeIcon
    notifyIcon.Text = "Presentation mode is off"
    status = "Normal"
    PublishStatus
    csTimer.Stop
End Sub
Sub PresentationMode
    watcher.Watch = True
    notifyIcon.EnableMenuItem normalModeMenuIndex
    notifyIcon.DisableMenuItem presentationModeMenuIndex
    notifyIcon.SetIconByDllFile "%SystemRoot%\System32\powercpl.dll", 6, largeIcon
    status = "Presentation"
    PublishStatus
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
Sub PublishStatus
    Dim stream : Set stream = fso.OpenTextFile(statusFile, ForWriting, CreateNew)
    stream.WriteLine status
    stream.Close
    Set stream = Nothing
End Sub
Sub SetDurationUI
    sa.MinimizeAll
    Dim currentValue : currentValue = Round(csTimer.IntervalInHours, 4)
    Dim response : response = InputBox("Enter the desired duration of Presentation mode / Phone charger mode, in hours." & vbLf & vbLf & "Current value: " & currentValue, WScript.ScriptName, currentValue)
    sa.UndoMinimizeAll
    If "" = response Then Exit Sub
    csTimer.IntervalInHours =  response
    If "Presentation" = status Then
        PresentationMode 'reset timers
    End If
End Sub
Sub Help
    shell.PopUp helpMessage, 40, WScript.ScriptName, vbInformation + vbSystemModal
End Sub
Sub ListenForCallbacks
    While True
        If "Presentation" = status Then
            notifyIcon.Text = "Presentation mode is on" & vbLf & "Normal mode resumes in " & Round(csTimer.Interval/60000 - stopwatch/60, 0) & " min."
        End If
        WScript.Sleep 200
    Wend
End Sub

Const synchronous = True
Const largeIcon = True, smallIcon = False
Const PresentationState = 3, NormalState = 0
Const ForWriting = 2, CreateNew = True
Dim watcher, notifyIcon, shell, fso, csTimer, sa, stopwatch, includer
Dim normalModeMenuIndex, presentationModeMenuIndex
Dim purpose, helpUrl, helpMessage
Dim statusFile, status

Sub Main
    Set watcher = CreateObject("VBScripting.Watcher")
    Set notifyIcon = CreateObject("VBScripting.NotifyIcon")
    Set csTimer = CreateObject("VBScripting.Timer")
    Set shell = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set sa = CreateObject("Shell.Application")
    Set includer = CreateObject("VBScripting.Includer")
    Execute includer.Read("VBSStopwatch")
    Set stopwatch = New VBSStopwatch
    Dim folder : folder = shell.ExpandEnvironmentStrings("%AppData%\VBScripting")
    If Not fso.FolderExists(folder) Then fso.CreateFolder folder
    statusFile = folder & "\" & fso.GetBaseName(WScript.ScriptName) & ".status"

    normalModeMenuIndex = 0
    presentationModeMenuIndex = 1
    notifyIcon.AddMenuItem "Normal mode", GetRef("NormalMode")
    notifyIcon.AddMenuItem "Presentation mode", GetRef("PresentationMode")
    notifyIcon.AddMenuItem "Phone charger mode", GetRef("ChargerMode")
    notifyIcon.AddMenuItem "Set duration", GetRef("SetDurationUI")
    notifyIcon.AddMenuItem "Help", GetRef("Help")
    notifyIcon.AddMenuItem "Exit", GetRef("CloseAndExit")
    notifyIcon.Visible = True

    csTimer.IntervalInHours = 1.5
    csTimer.AutoReset = False
    Set csTimer.Callback = GetRef("NormalMode")

    NormalMode
    ListenForCallbacks
End Sub
Sub CloseAndExit
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
