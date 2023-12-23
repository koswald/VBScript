Option Explicit
purpose = "Show a notification area icon with a menu option to prevent the computer and monitor from going to sleep."

helpMessage = "When presentation mode is on, the computer and monitor are typically prevented from going into a suspend (sleep) state or hibernation. The computer may still be put to sleep by other applications or by user actions such as closing a laptop lid or pressing a sleep button or power button." & vbLf & vbLf & "Phone charger mode is the same as presentation mode except that the workstation is locked, initially."

requires = "Sleep (menu item) functionality requires psshutdown from " & link1
Const link1 = "https://learn.microsoft.com/en-us/sysinternals/downloads/psshutdown"

Setup
csTimer.IntervalInHours = 1 'default presentation mode timeout
icon = Split( icon3, "|" )
NormalMode
ListenForCallbacks

Dim sh, fso, sa 'native WScript objects
Dim watcher, notifyIcon, csTimer, stopwatch, includer, format 'objects from github.com/koswald/vbscript
Dim normalModeMenuIndex, presentationModeMenuIndex 'integer: notification icon menu index
Dim purpose, helpUrl, helpMessage 'strings
Dim statusFile 'filespec of the file to which status is published
Dim status 'string: '"Presentation" or "Normal"
Dim icon 'array: filespec, index, and icon type (large/True or small/False) for Presentation and Normal modes
Dim requires 'string: used for internal documentation

'icon options
Const icon2 = "%SystemRoot%\System32\imageres.dll|101|False|%SystemRoot%\System32\imageres.dll|102|False" 'green & yellow shields
Const icon3 = "%SystemRoot%\System32\imageres.dll|96|True|%SystemRoot%\System32\deskmon.dll|0|True" 'monitor with moon & monitor without
Const icon4 = "%SystemRoot%\System32\hgcpl.dll|1|False|%SystemRoot%\System32\hgcpl.dll|0|False" 'dark LED & green LED
Const icon5 = "%SystemRoot%\System32\DDORes.dll|19|False|%SystemRoot%\System32\DDORes.dll|15|False" 'dark flat screen & bright flat screen / small icons
Const icon6 = "%SystemRoot%\System32\DDORes.dll|19|True|%SystemRoot%\System32\DDORes.dll|15|True" 'dark flat screen & bright flat screen / large icons
Const icon7 = "%SystemRoot%\System32\comres.dll|8|False|%SystemRoot%\System32\comres.dll|12|False" 'checkmark on green shield & checkmark on gold shield
Const ico_normFile = 0, ico_normIndex = 1, ico_normType = 2, ico_presentFile = 3, ico_presentIndex = 4, ico_presentType = 5 'icon array indexes

Const synchronous = True 'sh.Run constant, arg #3
Const hidden = 0 'sh.Run constant, arg #2
Const ForWriting = 2, CreateNew = True 'OpenTextFile args #2 and #3

Sub Setup
    Set sh = CreateObject( "WScript.Shell" )
    dataFolder = sh.ExpandEnvironmentStrings("%AppData%\VBScripting")
    Set fso = CreateObject( "Scripting.FileSystemObject" )
    If Not fso.FolderExists(dataFolder) Then fso.CreateFolder dataFolder
    Set format = CreateObject( "VBScripting.StringFormatter" )
    statusFile = format(Array("%s\%s.status", dataFolder, fso.GetBaseName(WScript.ScriptName)))

    Set notifyIcon = CreateObject( "VBScripting.NotifyIcon" ) 'Err.Number &H80131040
    notifyIcon.AddMenuItem "Normal mode", GetRef( "NormalMode" )
        normalModeMenuIndex = 0
    notifyIcon.AddMenuItem "Presentation mode", GetRef( "PresentationMode" )
        presentationModeMenuIndex = 1
    notifyIcon.AddMenuItem "Phone charger mode", GetRef( "ChargerMode" )
    notifyIcon.AddMenuItem "Set duration", GetRef( "SetDurationUI" )
    notifyIcon.AddMenuItem "Start screensaver", GetRef( "StartScreenSaver" )
    notifyIcon.AddMenuItem "Lock workstation   (Windows key + L)", GetRef( "LockWorkStation" )
    notifyIcon.AddMenuItem "Sleep", GetRef( "Sleep" )
    notifyIcon.AddMenuItem "Turn off monitor", GetRef( "MonitorOff" )
    notifyIcon.AddMenuItem "Edit " & WScript.ScriptName, GetRef( "EditScript" )
    notifyIcon.AddMenuItem "Edit " & WScript.ScriptName & " elevated", GetRef( "EditScriptElevated" )
    notifyIcon.AddMenuItem "Help", GetRef( "Help" )
    notifyIcon.AddMenuItem "Exit " & WScript.ScriptName, GetRef( "CloseAndExit" )
    notifyIcon.Visible = True

    Set csTimer = CreateObject( "VBScripting.Timer" )
    csTimer.AutoReset = False
    Set csTimer.Callback = GetRef( "NormalMode" )

    Set watcher = CreateObject( "VBScripting.Watcher" )
    Set sa = CreateObject( "Shell.Application" )
    Set includer = CreateObject( "VBScripting.Includer" )
    Execute includer.Read( "VBSStopwatch" )
    Set stopwatch = New VBSStopwatch

    Dim datafolder
End Sub

Sub NormalMode
    watcher.Watch = False
    notifyIcon.DisableMenuItem normalModeMenuIndex
    notifyIcon.EnableMenuItem presentationModeMenuIndex
    notifyIcon.SetIconByDllFile icon(ico_normFile), icon(ico_normIndex), icon(ico_normType)
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
    LockWorkStation
    PresentationMode
End Sub

Sub StartScreenSaver
    sh.Run "%SystemRoot%\system32\scrnsave.scr"
End Sub

Sub LockWorkStation
    sh.Run "rundll32 user32.dll,LockWorkStation",, synchronous
End Sub

Sub EditScript
    sh.Run format(Array("notepad ""%s""", WScript.ScriptFullName))
End Sub

Sub EditScriptElevated
    sa.ShellExecute "notepad", WScript.ScriptFullName,, "runas"
End Sub

Sub PublishStatus(newStatus)
    status = newStatus
    Set stream = fso.OpenTextFile(statusFile, ForWriting, CreateNew)
    stream.WriteLine newStatus
    stream.Close
    Set stream = Nothing

    Dim stream 'text stream for writing
End Sub

Sub SetDurationUI
    currentValue = Round(csTimer.IntervalInHours, 4)
    prompt = format(Array(" Enter the desired duration in hours %s of Presentation mode / Phone charger mode. %s Current value: %s", vbLf, vbLf & vbLf, currentValue))
    caption = WScript.ScriptName
    suggestedValue = currentValue
    sa.MinimizeAll
    response = InputBox(prompt, caption, suggestedValue)
    While Not IsNumeric(response)
        sh.PopUp "Presentation mode duration must be numeric.", 4, WScript.ScriptName, vbSystemModal + vbInformation
        response = InputBox(prompt, caption, suggestedValue)
    Wend
    sa.UndoMinimizeAll
    If "" = response Then Exit Sub
    csTimer.IntervalInHours = response
    If "Presentation" = status Then
        PresentationMode 'reset timers
    End If

    Dim currentValue 'current duration of Presentation mode in hours, if it were to be activated or reactivated
    Dim response 'InputBox return value
    Dim prompt, caption, suggestedValue 'InputBox arguments
End Sub

Sub Sleep
    On Error Resume Next
        sh.Run "psshutdown -d -t 0", hidden
        If Err Then
            ' show error message with link that can be easily copied
            sa.MinimizeAll
            InputBox requires, WScript.ScriptName, link1
            sa.UndoMinimizeAll
        End If
    On Error Goto 0
End Sub

Sub MonitorOff
    With CreateObject( "VBScripting.Admin" )
        .MonitorOff
    End With
End Sub

Sub Help
    sh.PopUp helpMessage, 80, WScript.ScriptName, vbInformation + vbSystemModal
End Sub

Sub ListenForCallbacks
    While True
        intervalInMinutes = csTimer.Interval/60000
        elapsedMinutes = stopwatch/60
        If "Presentation" = status Then
            notifyIcon.Text = format(Array(" Presentation mode is on %s Normal mode resumes in %s min.", vbLf, Round(intervalInMinutes - elapsedMinutes, 0)))
        Else notifyIcon.Text = "Presentation mode is off"
        End If
        WScript.Sleep 2000
    Wend

    Dim elapsedMinutes 'how long Presentation mode has been activated
    Dim intervalInMinutes 'C# timer's current setting for the max. time that Presentation mode will last before reverting to normal mode
End Sub

Sub CloseAndExit
    PublishStatus "Normal"
    watcher.Dispose
    Set watcher = Nothing
    notifyIcon.Dispose
    Set notifyIcon = Nothing
    csTimer.Dispose
    Set csTimer = Nothing
    Set sh = Nothing
    Set fso = Nothing
    Set sa = Nothing
    Set includer = Nothing
    Set stopwatch = Nothing
    Set format = Nothing
    WScript.Quit
End Sub
