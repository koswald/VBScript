' Elevate privileges of the specified file or command.
' Place a shortcut to this file in %AppData%\Microsoft\Windows\SendTo;
' or use as a drop target or from the command line.

' Issues
' When starting pwsh elevated, with for example
'    elevate.vbs pwsh -NoExit -Command "$env:SomePath"
' the quotes are stripped out of the command.

Option Explicit

With WScript.Arguments
    If .Count = 0 Then
        Err.Raise 449,, "Expected a command line argument: the file or command to run with elevated privileges."
    End If
    Dim filespec : filespec = .item(0)
End With

With CreateObject( "Scripting.FileSystemObject" )
    ' Making use of the VBScripting.VBSApp object's GetArgsString method allows support for multiple arguments in the command.
    Dim app : Set app = CreateObject( "VBScripting.VBSApp" )
    app.Init WScript
    Dim cmdArgs : cmdArgs = _
        "/c cd """ & .GetParentFolderName(filespec) & """" & _
        " & start """" " & app.GetArgsString
End With

With CreateObject( "Shell.Application" )
    .ShellExecute "cmd", cmdArgs,, "runas"
End With

Set app = Nothing
