'demonstrate the functionality of the WindowsUpdatesPauser class 

With CreateObject("includer")
    Execute(.read("WindowsUpdatesPauser"))
End With

Dim wup : Set wup = New WindowsUpdatesPauser

With WScript.Arguments
    If 0 = .Count Then
        InfoBox "One command-line argument is required:" & vbLf & _
            "/pause, /resume, or /getstatus"
    ElseIf "/pause" = .item(0) Then
        wup.PauseUpdates
    ElseIf "/resume" = .item(0) Then
        wup.ResumeUpdates
    ElseIf "/getstatus" = .item(0) Then
        ShowStatus
    End If
End With

Sub InfoBox(msg)
    MsgBox msg, vbInformation, WScript.ScriptName
End Sub

Sub ShowStatus
    Dim status : status = wup.GetStatus
    If "Metered" = status Then
        InfoBox "Windows Updates are paused (" & wup.GetProfileName & ")."
    ElseIf "Unmetered" = status Then
        InfoBox "Windows Updates are enabled (" & wup.GetProfileName & ")."
    Else
        'if userInteractive,
        'an error message should appear,
        'offering to open the .config file;

        'whether or not userInteractive,
        'makes a log entry that status
        'could not be determined,
        'probably due to invalid network name
    End If
End Sub