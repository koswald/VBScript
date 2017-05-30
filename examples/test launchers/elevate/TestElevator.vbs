
'launch the tests with elevated privileges

'validate

msg = "Start " & WScript.ScriptName & " with the .bat file of the same name to ensure that it is started with the 64-bit exe."
With WScript.Arguments
    If 0 = .Count Then Err.Raise 1, WScript.ScriptName, vbLf & msg
    If Not "ensure_64-bit_exe" = .item(0) Then Err.Raise 1, WScript.ScriptName, vbLf & msg
End With

'initialize

Set sa = CreateObject("Shell.Application")
Set fso = CreateObject("Scripting.FileSystemObject")
Set incl = CreateObject("includer")
Execute(incl.read("VBSClipboard"))
Execute(incl.read("VBSArguments"))
Set incl = Nothing
Set clip = New VBSClipboard
Set args = New VBSArguments
scriptParent = GetParent(WScript.ScriptFullName)

'to facilitate rerunning the test, the main command is copied
'to the clipboard, ready to be pasted onto the command line with
'a right click, and also echoed to the console where it can be
'manually copied

mainCommand = "cscript //nologo TestLauncherElevated.vbs " & args.GetArgumentsString
clip.SetClipboardText mainCommand
args = "/k echo " & mainCommand & " & cd """ & scriptParent & """ & " & mainCommand

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
