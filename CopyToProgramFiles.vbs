With CreateObject("VBScripting.Includer")
    Execute .Read("FolderSender")
End With
With New FolderSender
    .SourceFolder = .Parent( WScript.ScriptFullName )
    .TargetFolder = "%ProgramFiles%\VBScripting"
    .Copy
End With
