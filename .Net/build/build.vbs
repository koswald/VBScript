
'run all .bat files in the folder,
'or just the one(s) dropped onto this file

Option Explicit : Initialize

If WScript.Arguments.Count Then
    RunSelected
Else RunAll
End If

Cleanup

Sub RunSelected
    Dim i, cmds : cmds = ""
    For i = 0 To WScript.Arguments.Count - 1
        cmds = cmds & " & " & Wrap(fso.GetFileName(WScript.Arguments.item(i)))
    Next
    RunElevated(cmds)
End Sub

Sub RunAll
    Dim file, folder, cmds : cmds = ""
    Set folder = fso.GetFolder(buildDirectory)
    For Each file In folder.Files
        If "bat" = LCase(fso.GetExtensionName(file.Name)) Then cmds = cmds & " & " & Wrap(file.Name)
    Next
    Set folder = Nothing
    RunElevated(cmds)
End Sub

'Run the specified (concatenated) command(s) with elevated privileges
Sub RunElevated(commands)
    sa.ShellExecute "cmd", format(Array("/k cd ""%s"" %s", buildDirectory, commands)),, "runas"
End Sub    

'Wrap with double quotes
Function Wrap(str)
    Wrap = """" & str & """"
End Function

Dim fso, sa, format
Dim buildDirectory

Sub Initialize
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set sa = CreateObject("Shell.Application")
    Dim inc : Set inc = CreateObject("includer")
    Execute inc.read("StringFormatter")
    Set format = New StringFormatter
    buildDirectory = fso.GetParentFolderName(WScript.ScriptFullName)
End Sub

Sub Cleanup
    Set fso = Nothing
    Set sa = Nothing
End Sub
