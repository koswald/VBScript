'Run all the batch (.bat) files in this folder, which compiles the .cs files and registers the resulting .dll files. Alternatively, one or more of the batch files can be dragged and dropped onto this file or entered from the command line:
'PS C:\VBScripting\.Net\build> build.vbs Admin.bat FileChooser.bat

'The User Account Control dialog will open for permission to elevate privileges, if privileges are not already elevated.

'Requires Setup.vbs to have been run,

Option Explicit
Dim fso 'Scripting.FileSystemObject object
Dim sh 'WScript.Shell object
Dim format 'StringFormatter object 
Dim stream 'text stream object for creating & writing to a temporary script file
Dim tempFile 'string: name of the temp script file
Dim parent 'string: this script's folder, the "build" folder
Dim powershell 'string suitable for starting a powershell.exe or pwsh.exe process
Const Force = True 'for DeleteFile method
Const synchronous = True 'for Run method

Initialize

If WScript.Arguments.Count Then
    RunSelected
Else RunAll
End If

Cleanup

Sub Initialize
    Dim script 'this script's filespec
    Const CreateNew = True, ForWriting = 2 'for the OpenTextFile method
    Set fso = CreateObject( "Scripting.FileSystemObject" )
    Set sh = CreateObject( "WScript.Shell" )
    With CreateObject( "VBScripting.Includer" )
        Execute .Read( "StringFormatter" )
        Set format = New StringFormatter
        Execute .Read( "Configurer" )
        Execute .Read( "VBSApp" )
    End With
    With New VBSApp
        .SetUserInteractive False
        .RestartUsing .GetExe, .DoExit, .DoElevate
    End With
    With New Configurer
        powershell = .PowerShell
    End With
    script = WScript.ScriptFullName
    parent = fso.GetParentFolderName(script)
    sh.CurrentDirectory = parent
    tempFile = LCase( "_build_.ps1" )
    Set stream = fso.OpenTextFile(tempFile, ForWriting, CreateNew)

End Sub

Sub RunSelected
    Dim i 'integer
    With WScript.Arguments
        For i = 0 To .Count - 1
            stream.WriteLine ".\" & fso.GetFileName(.item(i))
        Next
    End With
    RunScriptFile
End Sub

Sub RunAll
    Dim file 'file object
    For Each file In fso.GetFolder(parent).Files
        If "bat" = LCase(fso.GetExtensionName(file.Name)) _
        And Not tempFile = file.Name Then
            stream.WriteLine ".\" & file.Name
        End If
    Next
    RunScriptFile
End Sub

Sub RunScriptFile
    stream.Close
    sh.Run format( Array( _
        """%s"" -ExecutionPolicy Bypass -NoExit -Command Set-Location '%s'; .\%s", _
        powershell, parent, tempFile _
    )),, synchronous
End Sub

Sub Cleanup
    ' sh.Run "notepad " & tempFile,, synchronous 'view the script file before deleting
    fso.DeleteFile tempFile, Force
    Set fso = Nothing
    Set sh = Nothing
    Set format  = Nothing
    Set stream = Nothing
End Sub
