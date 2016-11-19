
'start ElevatedTestLauncher.vbs with elevated privileges

'validate

msg = vbLf & "Start " & WScript.ScriptName & " with the .bat file of the same name, to ensure that the 64-bit .exe is used, if available."
If 0 = WScript.Arguments.Count Then Err.Raise 1, WScript.ScriptName, msg
If Not "Ensure_64-bit_exe" = WScript.Arguments(0) Then Err.Raise 1, WScript.ScriptName, msg

'initialize

Set fso = CreateObject("Scripting.FileSystemObject")
Set sa = CreateObject("Shell.Application")
parent = fso.GetParentFolderName(WScript.ScriptFullName)

'launch elevated tests with elevated privileges

sa.ShellExecute "cmd", "/k cd """ & parent & """ & cscript //nologo """ & parent & "\ElevatedTestLauncher.vbs""", "", "runas"

'garbage collection

Set sa = Nothing
Set fso = Nothing