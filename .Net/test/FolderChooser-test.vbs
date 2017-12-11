
'test FolderChooser.dll

With CreateObject("VBScripting.FolderChooser")
    .Title = "Select a folder - " & WScript.ScriptName
    .InitialDirectory = "..\.."
    folder = .FolderName
    If "" = folder Then
        MsgBox "Dialog cancelled"
    Else
        MsgBox "Chosen folder: " & vbLf & folder
    End If
End With

