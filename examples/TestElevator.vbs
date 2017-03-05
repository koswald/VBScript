
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
parent = fso.GetParentFolderName(WScript.ScriptFullName)
Set fso = Nothing

'elevate

sa.ShellExecute "cmd", "/k cd """ & parent & """ & cscript //nologo """ & parent & "\TestLauncher.vbs"" " & spec,, "runas"

'clean up

Set sa = Nothing
