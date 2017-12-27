
'Setup the VBScript utilities

'Registers the dependency manager scriptlet, includer.wsc

Option Explicit : Initialize

Const scriptlet = "class\includer.wsc" 'dependency manager scriptlet
Const buildFolder = ".Net\build"
Const sourceCreator = ".Net\config\CreateEventSource.vbs"
Const tests = "examples\test launchers\TestLauncherStandard.vbs"
Const runTests = True

Main
ReleaseObjectMemory

Sub Main
    'verify that we can find the scriptlet
    If Not fso.FileExists(scriptlet_) Then
        Err.Raise 1,, "Couldn't find the required scriptlet: " & scriptlet_
    End If

    'commands for registering the scriptlet for 32-bit or 64-bit,
    'according to system bitness
    args = format(Array( _
        "/k cd ""%s"" & echo. & " & _
        "echo Registering scriptlet & " & _
        "%SystemRoot%\System32\regsvr32 /s ""%s""", _
        projectFolder, scriptlet_ _
    ))

    'command for registering for 32-bit apps on 64-bit systems
    If fso.FolderExists(sh.ExpandEnvironmentStrings("%SystemRoot%\SysWow64")) Then
        args = format(Array("%s & echo. & " & _
            "echo Registering scriptlet for 32-bit apps & echo. & " & _
            "%SystemRoot%\SysWow64\regsvr32 /s ""%s""", _
            args, scriptlet_ _
        ))
    End If

    'commands for compiling and registering the VBS extensions
    args = format(Array("%s & cd ""%s""", args, buildFolder_))
    Dim file
    For Each file In fso.GetFolder(buildFolder_).Files
        If "bat" = fso.GetExtensionName(file) Then
            args = format(Array("%s & ""%s""", args, file.Name))
        End If
    Next

    'command for creating the event log source
    args = format(Array("%s & echo. & " & _
        "echo Creating the event log source VBScripting & " & _
        """%s"" /quiet", args, sourceCreator_ _
    ))

    'run the setup commands
    sh.Run "cmd " & args,, synchronous

    'run some tests, if desired
    If runTests Then
        msg = "Setup can run the standard tests, which may take about 30 seconds."
        mode = vbOKCancel + vbInformation + vbSystemModal
        If vbOK = MsgBox(msg, mode, WScript.ScriptName) Then
            sh.Run "%ComSpec% /k cscript.exe //nologo """ & tests_ & """"
        End If
    End If

End Sub

Function GetVBSObj(name)
    Dim stream : Set stream = fso.OpenTextFile("class\" & name & ".vbs", 1)
    Execute stream.ReadAll
    stream.Close
    Set stream = Nothing
    Execute "Set GetVBSObj = New " & name
End Function

Sub ReleaseObjectMemory
    Set sa = Nothing
    Set sh = Nothing
    Set fso = Nothing
End Sub

Const synchronous = True
Dim args, msg, mode
Dim projectFolder
Dim sa, sh, fso
Dim format
Dim scriptlet_, buildFolder_, sourceCreator_, tests_

Sub Initialize
    Set sa = CreateObject("Shell.Application")
    Set sh = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set format = me.GetVBSObj("StringFormatter")
    Dim pc : Set pc = GetVBSObj("PrivilegeChecker")
    Dim app : Set app = GetVBSObj("VBSApp")
    If Not pc Then 
        app.SetUserInteractive False
        app.RestartWith "wscript", "/c", True
    End If
    projectFolder = fso.GetParentFolderName(WScript.ScriptFullName)
    sh.CurrentDirectory = projectFolder
    scriptlet_ = fso.GetAbsolutePathName(scriptlet)
    tests_ = fso.GetAbsolutePathName(tests)
    buildFolder_ = fso.GetAbsolutePathName(buildFolder)
    sourceCreator_ = fso.GetAbsolutePathName(sourceCreator)
End Sub
