'Script for AnyFolder.wsf

With CreateObject( "VBScripting.Includer" )
    Execute .Read( "VBSApp" )
End With
With CreateObject( "Scripting.FileSystemObject" )
    parent = .GetParentFolderName(WScript.ScriptFullName)
End With
Dim app : Set app = New VBSApp
With CreateObject( "WScript.Shell" )
    .Run """" & parent & "\AnyFolder.hta"" " & app.GetArgsString
End With
Set app = Nothing
