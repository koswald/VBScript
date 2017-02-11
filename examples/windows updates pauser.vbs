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
        MsgBox "status: " & wup.GetStatus
    End If
End With