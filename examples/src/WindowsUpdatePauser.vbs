'script for "WindowsUpdatePauser.hta"

Sub PauseUpdates
    wup.PauseUpdates
    UpdateStatus wup.GetStatus
End Sub

Sub ResumeUpdates
    wup.ResumeUpdates
    UpdateStatus wup.GetStatus
End Sub

Sub UpdateStatus(status)
    If "Metered" = status Then 'updates are paused
        Enable iResume
        SetTitle iResume, "Resume Windows Updates (" & wup.GetProfileName & ")"
        Disable iPause
        SetTitle iPause, ""
    ElseIf "Unmetered" = status Then 'updates are enabled
        Enable iPause
        SetTitle iPause, "Pause Windows Updates (" & wup.GetProfileName & ")"
        Disable iResume
        SetTitle iResume, ""
    Else
        SelfClose
    End If
End Sub

Sub OpenConfigFile
    wup.OpenConfigFile
    SelfClose
End Sub

'Return the html input element/object specified
Function GetInput(i)
    Set GetInput = document.getElementsByTagName( "input" )(i)
End Function
Sub Enable(i)
    GetInput(i).disabled = False
End Sub
Sub Disable(i)
    GetInput(i).disabled = True
End Sub
'Set the tooltip that appears when the mouse hovers over the specified button
Sub SetTitle(i, title)
    GetInput(i).title = title
End Sub
Sub SetClass(i, name)
    GetInput(i).className = name
End Sub
Sub SelfClose
    self.close
End Sub

Dim wup 'WindowsUpdatesPauser object
Const iPause = 0, iResume = 1, iConfig = 2, iClose = 3 'button index numbers

Sub Window_OnLoad
    Self.ResizeTo 238, 30
    Self.MoveTo 100, 0
    With CreateObject( "VBScripting.Includer" )
        Execute .Read( "WindowsUpdatesPauser" )
    End With
    Set wup = New WindowsUpdatesPauser
    document.title = wup.GetAppName
    SetTitle iClose, "Close the Windows Updates Pauser"
    SetTitle iConfig, "Open the .config file"
    SetClass iPause, "pause"
    SetClass iResume, "resume"
    SetClass iConfig, "config"
    SetClass iClose, "close"
    UpdateStatus wup.GetStatus
End Sub

