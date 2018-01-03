
'Setup the VBScript utilities

'Registers the dependency manager scriptlet, includer.wsc,
'and builds the VBScript extension libraries.

'The User Account Control dialog will open
'to verify elevation of privileges.

'Use /u to uninstall

Option Explicit : Initialize

Const scriptlet = "class\includer.wsc" 'dependency manager scriptlet
Const buildFolder = ".Net\build"

Main
ReleaseObjectMemory

Sub Main
    If installing Then
        ValidateScriptlet
        PrepWscRegistrationX64
        PrepWscRegistrationX86
        PrepDllRegistration
        FinalInstruction
        RunCommands
        CreateEventLogSource
    ElseIf uninstalling Then
        DeleteEventLogSource
        PrepDllRegistration
        PrepWscRegistrationX64
        PrepWscRegistrationX86
        FinalInstruction
        RunCommands
        DeleteScriptletKeys
    End If
End Sub

'verify that the scriptlet can be found
Sub ValidateScriptlet
    If Not fso.FileExists(scriptlet_) Then
        Err.Raise 1,, "Couldn't find the required scriptlet: " & scriptlet_
    End If
End Sub

'prepare commands for registering the scriptlet for 32-bit or 64-bit,
'according to system bitness
Sub PrepWscRegistrationX64
    args = format(Array( _
        "%s & echo. & " & _
        "echo %s scriptlet & " & _
        "%SystemRoot%\System32\regsvr32 %s /s ""%s""", _
        args, registerVerb, wscFlag, scriptlet_ _
    ))
End Sub

'prepare command for registering for 32-bit apps on 64-bit systems
Sub PrepWscRegistrationX86
    If fso.FolderExists(sh.ExpandEnvironmentStrings("%SystemRoot%\SysWow64")) Then
        args = format(Array("%s & echo. & " & _
            "echo %s scriptlet for 32-bit apps & echo. & " & _
            "%SystemRoot%\SysWow64\regsvr32 %s /s ""%s""", _
            args, registerVerb, wscFlag, scriptlet_ _
        ))
    End If
End Sub

'prepare commands for compiling and registering/unregistering the VBS extensions
Sub PrepDllRegistration
    args = format(Array("%s & cd ""%s""", args, buildFolder_))
    Dim file
    For Each file In fso.GetFolder(buildFolder_).Files
        If "bat" = fso.GetExtensionName(file) Then
            args = format(Array("%s & ""%s"" %s", args, file.Name, dllFlag))
        End If
    Next
    args = format(Array("%s & echo.", args))
End Sub

Sub FinalInstruction
    args = format(Array("%s & " & _
        "echo Close this command prompt window to finish %s.", _
        args, setupNoun))
End Sub

'run the setup/uninstall commands
Sub RunCommands
    sh.Run "cmd " & args,, synchronous
End Sub

Sub CreateEventLogSource
    On Error Resume Next
        Dim va : Set va = CreateObject("VBScripting.Admin")
        va.CreateEventSource va.EventSource
        Set va = Nothing
    On Error Goto 0
End Sub

Sub DeleteEventLogSource
    On Error Resume Next
        Dim va : Set va = CreateObject("VBScripting.Admin")
        va.DeleteEventSource va.EventSource
        Set va = Nothing
    On Error Goto 0
End Sub

'Remove the registry keys associated with the scriptlet;
'regsvr32.exe may show a success message on unregister
'without removing the registry keys.
Sub DeleteScriptletKeys
    'WScript.Shell RegDelete requires subkeys to be removed first
    Dim keys : keys = Array( _
        "HKEY_CLASSES_ROOT\CLSID\{ADCEC089-30DE-11D7-86BF-00606744568C}\InprocServer32\", _
        "HKEY_CLASSES_ROOT\CLSID\{ADCEC089-30DE-11D7-86BF-00606744568C}\ProgID\", _
        "HKEY_CLASSES_ROOT\CLSID\{ADCEC089-30DE-11D7-86BF-00606744568C}\ScriptletURL\", _
        "HKEY_CLASSES_ROOT\CLSID\{ADCEC089-30DE-11D7-86BF-00606744568C}\VersionIndependentProgID\", _
        "HKEY_CLASSES_ROOT\CLSID\{ADCEC089-30DE-11D7-86BF-00606744568C}\", _
        "HKEY_CLASSES_ROOT\Wow6432Node\CLSID\{ADCEC089-30DE-11D7-86BF-00606744568C}\InprocServer32\", _
        "HKEY_CLASSES_ROOT\Wow6432Node\CLSID\{ADCEC089-30DE-11D7-86BF-00606744568C}\ProgID\", _
        "HKEY_CLASSES_ROOT\Wow6432Node\CLSID\{ADCEC089-30DE-11D7-86BF-00606744568C}\ScriptletURL\", _
        "HKEY_CLASSES_ROOT\Wow6432Node\CLSID\{ADCEC089-30DE-11D7-86BF-00606744568C}\VersionIndependentProgID\", _
        "HKEY_CLASSES_ROOT\Wow6432Node\CLSID\{ADCEC089-30DE-11D7-86BF-00606744568C}\", _
        "HKEY_CLASSES_ROOT\includer\CLSID\", _
        "HKEY_CLASSES_ROOT\includer\")
    Dim i : For i = 0 To UBound(keys)
        On Error Resume Next
            sh.RegDelete keys(i)
        On Error Goto 0
    Next
End Sub

Sub ReleaseObjectMemory
    Set sa = Nothing
    Set sh = Nothing
    Set fso = Nothing
End Sub

Const synchronous = True
Dim args, msg, mode
Dim projectFolder
Dim installing, uninstalling, registerVerb, setupNoun
Dim wscFlag, dllFlag
Dim sa, sh, fso
Dim format
Dim scriptlet_, buildFolder_

Sub Initialize
    Set sa = CreateObject("Shell.Application")
    Set sh = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set format = New StringFormatter
    Dim pc : Set pc = New PrivilegeChecker

    'convert relative paths to absolute paths
    projectFolder = fso.GetParentFolderName(WScript.ScriptFullName)
    sh.CurrentDirectory = projectFolder
    scriptlet_ = fso.GetAbsolutePathName(scriptlet)
    buildFolder_ = fso.GetAbsolutePathName(buildFolder)

    'look for /u on the command line
    With WScript.Arguments
        uninstalling = False
        Dim i : For i = 0 To .Count - 1
            If "/u" = LCase(.item(i)) Then uninstalling = True
        Next
        Dim setupFlag 'flag for restarting this script
        If uninstalling Then
            setupFlag = "/u"
            registerVerb = "Unregistering"
            setupNoun = "uninstalling"
            wscFlag = "/u"
            dllFlag = "/unregister"
        Else
            setupFlag = ""
            registerVerb = "Registering"
            setupNoun = "setup"
            wscFlag = ""
            dllFlag = ""
            installing = True
        End If
    End With
    If Not pc Then

        'restart this script to elevate privileges
        Dim restartArgs : restartArgs = format(Array( _
            "/c cd ""%s"" & start wscript ""%s"" %s", _
            projectFolder, WScript.ScriptFullName, setupFlag _
        ))
        sa.ShellExecute "cmd", restartArgs,, "runas"
        ReleaseObjectMemory
        WScript.Quit
    End If

    'prepare initial cmd.exe arguments for setup/uninstall
    args = format(Array( "/k cd ""%s""", projectFolder))
End Sub

'This is a pared-down version of StringFormatter.vbs found in the "class" folder.
Class StringFormatter

    'Returns a formatted string. The parameter is an array whose first element contains the pattern of the returned string. The first %s in the pattern is replaced by the next element in the array. The second %s in the pattern is replaced by the next element in the array, and so on. Variant subtypes tested OK with %s include string, integer, and single. Format is the default property for the class, so the property name is optional. If there are too many or too few %s instances, then an error will be raised.
    Public Default Function Format(array_)
        Const startPosition = 1
        Const replacemtCount = 1
        Dim arr : arr = array_
        Dim i, pattern : pattern = arr(0)
        For i = 1 To UBound(arr)
            If Not CBool(InStr(pattern, surrogate)) Then Err.Raise 1,, "There are too few instances of " & surrogate & vbLf & "Pattern: " & arr(0)
            pattern = Replace(pattern, surrogate, arr(i), startPosition, replacemtCount)
        Next
        If InStr(pattern, surrogate) Then Err.Raise 1,, "There are too many instances of " & surrogate & vbLf & "Pattern: " & arr(0)
        Format = pattern
    End Function

    'Remark: Optional. Sets the string that the Format method will replace with the specified array element(s), %s by default.
    Sub SetSurrogate(newSurrogate)
        surrogate = newSurrogate
    End Sub

    Private surrogate

    Sub Class_Initialize
        SetSurrogate "%s"
    End Sub
End Class

'Adapted from http://stackoverflow.com/questions/4051883/batch-script-how-to-check-for-admin-rights/21295806
Class PrivilegeChecker

    'Returns True if the calling script is running with elevated privileges, False if not. Privileged is the default property.
    Public Default Function Privileged
        Dim sh : Set sh = CreateObject("WScript.Shell")
        Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
        Dim privileged_, unprivileged_, undefined_
        privileged_ = "privileged"
        unprivileged_ = "unprivilegd" 'intentionally misspelled for unique search results
        undefined_ = "undefined"
        Privileged = undefined_

        'create a temporary .bat file
        Dim tempFile : tempFile = sh.ExpandEnvironmentStrings("%temp%\" & fso.GetTempName & ".bat")
        Dim bf : Set bf = fso.OpenTextFile(tempFile, 2, True) 'create the batch file; open for writing
        bf.WriteLine "@echo off"
        bf.WriteLine "call :isAdmin"
        bf.WriteLine "if %errorlevel% == 0 ("
        bf.WriteLine "echo " & privileged_
        bf.WriteLine ") else ("
        bf.WriteLine "echo " & unprivileged_
        bf.WriteLine ")"
        bf.WriteLine "exit /b"
        bf.WriteLine ":isAdmin"
        bf.WriteLine "fsutil dirty query %systemdrive% >nul"
        bf.WriteLine "exit /b"
        bf.Close
        Set bf = Nothing

        'run the batch file and parse the output
        Dim pipe : Set pipe = sh.Exec("%ComSpec% /c """ & tempFile & """")
        Dim line
        While Not pipe.StdOut.AtEndOfStream
            line = pipe.StdOut.ReadLine
            If InStr(line, privileged_) Then
                Privileged = True
            ElseIf InStr(line, unprivileged_) Then
                Privileged = False
            End If
        Wend

        'cleanup
        Set pipe = Nothing
        fso.DeleteFile(tempFile)
        Set sh = Nothing
        Set fso = Nothing

        'raise an error if privileges are undefined
        If Privileged = undefined_ Then Err.Raise 1,, "The PrivilegeChecker could not determine privileges"
    End Function
End Class
