With CreateObject("includer")
    Execute(.read("WindowsUpdatesPauser"))
End With

Dim wup : Set wup = New WindowsUpdatesPauser

With WScript.Arguments
    If 0 = .Count Then
        MsgBox "One command-line argument is required.", vbInformation, WScript.ScriptName
    ElseIf "/pause" = .item(0) Then
        wup.PauseUpdates
    ElseIf "/resume" = .item(0) Then
        wup.ResumeUpdates
    ElseIf "/getstatus" = .item(0) Then
        ShowStatus
    End If
End With

Sub ShowStatus
    Dim status : status = wup.GetStatus
    If "Metered" = status Then
        MsgBox "Windows Updates are paused."
    ElseIf "Unmetered" = status Then
        MsgBox "Windows Updates are enabled."
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