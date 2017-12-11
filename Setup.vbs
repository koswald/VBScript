
'Setup the VBScript utilities

'Registers the dependency manager scriptlet, includer.wsc, 
'and if desired, runs the standard tests

Option Explicit : Initialize

Const includer = "class\includer.wsc" 'dependency manager scriptlet
Const tests = "examples\test launchers\TestLauncherStandard.vbs"

'verify that we can find the scriptlet
If Not fso.FileExists(scriptlet) Then
    Err.Raise 1,, "Couldn't find the required scriptlet: " & scriptlet
End If

'register the scriptlet for both x86 (32-bit) and x64 (64-bit)
args = "/c " & _
    "%SystemRoot%\System32\regsvr32 /s """ & scriptlet & """ & " & _
    "%SystemRoot%\SysWow64\regsvr32 /s """ & scriptlet & """"
sa.ShellExecute "cmd", args,, "runas" 'elevate privileges

'test the setup by running the tests, if desired
msg = "Setup is finished." & vbLf & vbLf & "Before closing, Setup can run the standard tests, which may take about 30 seconds."
If vbOK = MsgBox(msg, vbOKCancel + vbInformation + vbSystemModal, WScript.ScriptName) Then
    sh.Run "%ComSpec% /k cscript.exe //nologo """ & tests & """"
End If

'Release object memory
Set sa = Nothing
Set sh = Nothing
Set fso = Nothing

Dim scriptlet, args, msg
Dim sa, sh, fso

Sub Initialize
    Set sa = CreateObject("Shell.Application")
    Set sh = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    scriptlet = fso.GetAbsolutePathName(includer)
End Sub
