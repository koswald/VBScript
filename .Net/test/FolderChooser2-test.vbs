With CreateObject( "VBScripting.FolderChooser2" )
    .Title = "Some title"
    .InitialDirectory = ".."
    .InitialDirectory = "%UserProfile%"
    folder = .FolderName
    If "" = folder Then
        MsgBox "Couldn't get a folder", vbInformation, WScript.ScriptName
    Else
        MsgBox "Chosen folder:" & vbLf & folder, vbInformation, WScript.ScriptName
    End If
End With
