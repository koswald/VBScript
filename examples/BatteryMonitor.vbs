'Battery monitor
'with system tray icon and warning messages
Option Explicit
Setup
Main

Sub Main
    notifier.AddMenuItem "Exit Battery Monitor", GetRef("Quit")
    notifier.Visible = True
    WaitForCallbacks
End Sub

Sub WaitForCallbacks
    While True
        ResetIcon
        WScript.Sleep 500
    Wend
End Sub
Sub ResetIcon
    index = IconIndex
    If index = previousIndex Then Exit Sub
    notifier.SetIconByDllFile resFile, index, iconSize
    previousIndex = index
End Sub
Function IconIndex
    Dim percent : percent = Charge
    Dim status : If PluggedIn Then status = "Plugged in" Else status = "Not plugged in"
    notifier.Text = format(Array( _
        "Battery charge %s%" & vbLf & _
        "%s", percent, status))
    resFile = resFile1
    If PluggedIn Then
        IconIndex = 2
        resFile = resFile2
    ElseIf percent < 11 Then
        IconIndex = 9
    ElseIf percent < 31 Then
        IconIndex = 10
    ElseIf percent < 51 Then
        IconIndex = 11
    ElseIf percent < 71 Then
        IconIndex = 12
    Else IconIndex = 13
    End If
    If percent < 50 _
    And Not PluggedIn _
    And (SecondsSinceLastWarning > 600 _
    Or neverBeenWarned) Then
        ShowWarning
    End If
End Function
Sub ShowWarning
    neverBeenWarned = False
    stopwatch.Reset
    sh.Run format(Array( _
        "mshta ""%s\BatteryStatus.hta""", _
        sh.CurrentDirectory))
End Sub
Function Charge
    Charge = Battery.EstimatedChargeRemaining
End Function
Function PluggedIn
    PluggedIn = Battery.BatteryStatus = 2
End Function
Function SecondsSinceLastWarning
    SecondsSinceLastWarning = stopwatch.Split
End Function
Function Battery
    Set Battery = wmi.Battery
End Function

Const resFile1 = "%SystemRoot%\System32\wpdshext.dll"
Const resFile2 = "%SystemRoot%\System32\powercpl.dll"
Const smallIcon = False
Const largeIcon = True
Dim resFile
Dim iconSize
Dim index, previousIndex
Dim neverBeenWarned
Dim notifier, wmi, format, stopwatch
Dim sh
Sub Setup
    Set notifier = CreateObject("VBScripting.NotifyIcon")
    Set includer = CreateObject("VBScripting.Includer")
    Set format = CreateObject("VBScripting.StringFormatter")
    Set sh = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    sh.CurrentDirectory = fso.GetParentFolderName(WScript.ScriptFullName)
    Execute includer.Read("WMIUtility")
    Set wmi = New WMIUtility
    Execute includer.Read("VBSStopwatch")
    Set stopwatch = New VBSStopwatch
    neverBeenWarned = True
    previousIndex = -1
    iconSize = smallIcon

    Set includer = Nothing
    Set fso = Nothing
    Dim includer, fso
End Sub
Sub Quit
    notifier.Dispose
    Set notifier = Nothing
    Set format = Nothing
    Set wmi = Nothing
    Set sh = Nothing
    Set stopwatch = Nothing
    WScript.Quit
End Sub
