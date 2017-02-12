With CreateObject("includer")
    Execute(.read("WindowsUpdatesPauser"))
End With

Dim wup : Set wup = New WindowsUpdatesPauser

With WScript.Arguments
    If "/pause" = .item(0) Then
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
    End If
End Sub