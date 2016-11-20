
'launch the tests with elevated privileges

'validate

msg = "Start " & WScript.ScriptName & " with the .bat file of the same name to ensure that it is started with the 64-bit exe."
If 0 = WScript.Arguments.Count Then Err.Raise 1, WScript.ScriptName, vbLf & msg
If Not "ensure_64-bit_exe" = WScript.Arguments(0) Then Err.Raise 1, WScript.ScriptName, vbLf & msg

'initialize

Set sa = CreateObject("Shell.Application")
Set fso = CreateObject("Scripting.FileSystemObject")
parent = fso.GetParentFolderName(WScript.ScriptFullName)
Set fso = Nothing

'elevate

sa.ShellExecute "cmd", "/k cd """ & parent & """ & cscript //nologo """ & parent & "\TestLauncher.vbs""",, "runas"

'clean up

Set sa = Nothing
