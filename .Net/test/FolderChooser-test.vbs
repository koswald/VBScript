
'test FolderChooser.dll

With CreateObject("VBScripting.FolderChooser")
    .Title = "Select a folder - " & WScript.ScriptName
    .InitialDirectory = "%UserProfile%"
    .InitialDirectory = ".."
    .InitialDirectory = "../../.."
    folder = .FolderName
    If "" = folder Then
        ShowMsg "Dialog cancelled"
    Else
        ShowMsg "Chosen folder: " & vbLf & folder
    End If
End With

Sub ShowMsg(msg)
    MsgBox msg, vbInformation, WScript.ScriptName
End Sub

