'Copy project files to %ProgramFiles%.
'Use /DeleteTarget to delete the target folder before copying.
'Notes
'1) User interactive.
'2) Does not require Setup.vbs to have been run.
'3) This script should remain in the project root folder.

Option Explicit
Dim fso 'Scripting.FileSystemObject
Dim sa 'Shell.Application object
Dim sourceFolder, targetFolder 'folder paths
Dim scr 'this file's filespec
Dim failedToDeleteTarget, failedToCreateTarget 'booleans
Dim msg, i, caption 'for MsgBox
Dim errNumber, errDescription 'integer, string
Const Force = True 'for DeleteFolder

targetFolder = "%ProgramFiles%\VBScripting"
Set fso = CreateObject("Scripting.FileSystemObject")
scr = WScript.ScriptFullName
sourceFolder = fso.GetParentFolderName( scr ) 'see Note 3
Set sa = CreateObject("Shell.Application")

With CreateObject( "WScript.Shell" )
    .CurrentDirectory = sourceFolder
    targetFolder = .ExpandEnvironmentStrings( _
        targetFolder )
End With

'Delete the target folder if requested, and if not already deleted

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
    msg = "Error attempting to delete the target folder" & _
        vbLf & targetFolder & vbLf & vbLf & _
        "Err.Description: " & vbTab & errDescription & vbLf & _
        "Err.Number:      " & vbTab & errNumber & vbLf & _
        "Hex(Err.Number): " & vbTab & Hex(errNumber) & vbLf & vbLf & _
        "Retry with elevated privileges?" & vbLf & vbLf & _
        "If this message appears repeatedly, " & _
        "stop any scripts using the VBScripting .NET libraries, and try again."
    i = vbOkCancel + vbSystemModal + vbInformation 
    caption = WScript.ScriptName
    If vbCancel = MsgBox( msg, i, caption) Then
        Quit
    End If
    sa.ShellExecute "wscript", _
        """" & scr & """ /DeleteTarget",, "runas"
    Quit
End If

'Create the target folder if not already created

failedToCreateTarget = false
If Not fso.FolderExists(targetFolder) Then
    On Error Resume Next
    fso.CreateFolder targetFolder
    If Err Then
        failedToCreateTarget = True
        errNumber = Err.Number
        errDescription = Err.Description
    End If
    On Error Goto 0
End If
If failedToCreateTarget Then
    msg = "Error attempting to create the target folder" & _
        vbLf & targetFolder & vbLf & vbLf & _
        "Err.Description: " & vbTab & errDescription & vbLf & _
        "Err.Number:      " & vbTab & errNumber & vbLf & _
        "Hex(Err.Number): " & vbTab & Hex(errNumber) & vbLf & vbLf & _
        "Retry with elevated privileges?"
    i = vbOKCancel + vbSystemModal + vbInformation
    caption = WScript.ScriptName
    If vbCancel = MsgBox( msg, i, caption) Then
        Quit
    End If
    sa.ShellExecute "wscript", _
        """" & scr & """",, "runas"
    Quit
End If

'Invoke the CopyHere method

sa.Namespace(targetFolder).CopyHere sourceFolder

Call Quit

Sub Quit
    Set fso = Nothing
    Set sa = Nothing
    WScript.Quit 
End Sub