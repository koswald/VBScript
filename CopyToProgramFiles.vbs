'Copy project files to %ProgramFiles%.
'Use /DeleteTarget to delete the target folder before copying.
'User interactive.
'Assumes that Setup.vbs has not been run.
'This script should remain in the project root folder.

Option Explicit
Dim fso 'Scripting.FileSystemObject
Dim sa 'Shell.Application object
Dim sourceFolder, targetFolder 'folder paths
Dim scr 'this file's filespec
Dim failedToDeleteTarget 'boolean
Dim msg, i, caption 'for MsgBox
Dim errNumber, errDescription
Const Force = True 'for DeleteFolder

targetFolder = "%ProgramFiles%\VBScripting"
Set fso = CreateObject("Scripting.FileSystemObject")
scr = WScript.ScriptFullName
sourceFolder = fso.GetParentFolderName( scr )
Set sa = CreateObject("Shell.Application")

With CreateObject( "WScript.Shell" )
    .CurrentDirectory = sourceFolder
    targetFolder = .ExpandEnvironmentStrings( _
        targetFolder )
End With

With WScript.Arguments.Named
    failedToDeleteTarget = False
    If .Exists( "DeleteTarget" ) _
    And fso.FolderExists( targetFolder ) Then
        On Error Resume Next
        fso.DeleteFolder targetFolder, Force
        If Err Then 
            failedToDeleteTarget = True
            errNumber = Err.Number
            errDescription = Err.Description
        End If
        On Error Goto 0
    End If
End With
If failedToDeleteTarget Then
    msg = "Retry with elevated privileges?"
    msg = "Error attempting to delete the target folder" & _
        vbLf & targetFolder & vbLf & vbLf & _
        "Err.Description: " & vbTab & errDescription & vbLf & _
        "Err.Number:      " & vbTab & errNumber & vbLf & _
        "Hex(Err.Number): " & vbTab & Hex(errNumber) & vbLf & vbLf & _
        msg
    i = vbOkCancel + vbSystemModal + vbInformation 
    caption = WScript.ScriptName
    If vbCancel = MsgBox( msg, i, caption) Then
        Quit
    End If
    sa.ShellExecute "wscript", _
        """" & scr & """ /DeleteTarget",, "runas"
    Quit
End If

With fso.OpenTextFile("class\FolderSender.vbs")
    Execute .ReadAll
    .Close
End With
On Error Resume Next
With New FolderSender
    On Error Goto 0
    .SourceFolder = sourceFolder
    .TargetFolder = targetFolder
    .Copy
End With

Call Quit
Sub Quit
    Set fso = Nothing
    Set sa = Nothing
    WScript.Quit 
End Sub