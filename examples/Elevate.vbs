' Elevate privileges of the specified file or command.
' A shortcut to this file can be placed in SendTo.
' Or use as a drop target or from the command line.
' Making use of the VBScripting.VBSApp object's GetArgsString method allows support for multiple arguments in the command. 

Option Explicit

With WScript.Arguments
    If .Count = 0 Then
        Err.Raise 1,, "Expected a command line argument: the file to open with elevated privileges."
    End If
    filespec = .item(0)
End With

With CreateObject("Scripting.FileSystemObject")
    Set app = CreateObject("VBScripting.VBSApp")
    app.Init WScript
    cmdArgs = _
        "/c cd """ & .GetParentFolderName(filespec) & """" & _
        " & start """" " & app.GetArgsString
End With

With CreateObject("Shell.Application")
    .ShellExecute "cmd", cmdArgs,, "runas"
End With

Dim filespec, app, cmdArgs