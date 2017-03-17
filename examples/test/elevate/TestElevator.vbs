
'launch the tests with elevated privileges

'validate

msg = "Start " & WScript.ScriptName & " with the .bat file of the same name to ensure that it is started with the 64-bit exe."
With WScript.Arguments
    If 0 = .Count Then Err.Raise 1, WScript.ScriptName, vbLf & msg
    If Not "ensure_64-bit_exe" = .item(0) Then Err.Raise 1, WScript.ScriptName, vbLf & msg
    spec = "" : If 2 <= .Count Then spec = .item(1)
End With

'initialize

Set sa = CreateObject("Shell.Application")
Set fso = CreateObject("Scripting.FileSystemObject")
scriptParent = GetParent(WScript.ScriptFullName)
args = "/k cd """ & scriptParent & """ & cscript //nologo """ & scriptParent & "\TestLauncherElevated.vbs"" " & spec

'verify args

'If vbCancel = MsgBox(args, vbOKCancel + vbQuestion, "Verify args - " & WScript.ScriptName) Then Quit

'elevate

sa.ShellExecute "cmd", args,, "runas"

ReleaseObjectMemory

Sub Quit
    ReleaseObjectMemory
    WScript.Quit
End Sub

Sub ReleaseObjectMemory
    Set sa = Nothing
    Set fso = Nothing
End Sub

Function GetParent(item)
    GetParent = fso.GetParentFolderName(item)
End Function
