'Call RegisterWsc.vbs with /u
With WScript.Arguments
    If .Count = 0 Then Err.Raise 1,, "Argument required: the .wsc file to unregister."
    file = .item(0)
End With
With CreateObject("Scripting.FileSystemObject")
    parent = .GetParentFolderName(WScript.ScriptFullName)
End With
With CreateObject("WScript.Shell")
    .CurrentDirectory = parent
    .Run "wscript RegisterWsc.vbs /u """ & file & """"
End With
