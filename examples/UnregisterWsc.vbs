'Call RegisterWsc.wsf with /u
With WScript.Arguments
    If .Count = 0 Then Err.Raise 449,, "Argument required: the .wsc file to unregister."
    file = .item(0)
End With
With CreateObject( "Scripting.FileSystemObject" )
    parent = .GetParentFolderName(WScript.ScriptFullName)
End With
With CreateObject( "WScript.Shell" )
    .CurrentDirectory = parent
    .Run "wscript RegisterWsc.wsf /u """ & file & """"
End With
