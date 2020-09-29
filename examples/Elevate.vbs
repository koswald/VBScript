'Elevate privileges of the specified file or command.
'Use as a drop target or from the command line.
'A shortcut to this file can be placed in SendTo.

Option Explicit

With WScript.Arguments
    If .Count = 0 Then
        Err.Raise 1,, "Expected a command line argument: the file to open with elevated privileges."
    ElseIf .Count > 1 Then
        Err.Raise 2,, "Can only elevate one item at a time."
    End If
    Dim filespec : filespec = .item(0)
End With

With CreateObject("Scripting.FileSystemObject")
    If Not .FileExists(filespec) Then
        Err.Raise 3,, "Cannot find the file '" & filespec "'"
    End If
    Dim cmdArgs : cmdArgs = _
        "/c cd """ & .GetParentFolderName(filespec) & """ " & _
        "& start """" """ & filespec & """"
End With

With CreateObject("Shell.Application")
    .ShellExecute "cmd", cmdArgs,, "runas"
End With

