
'Setup the VBScript utilities

'Registers the windows script component, Includer.wsc,
'other project components, and builds the VBScript
'extension libraries (.dll files).

'The User Account Control dialog will open
'to verify elevation of privileges.

'Use /u to uninstall

Option Explicit : Initialize

Const componentFolder_ = "class\wsc"
Const buildFolder_ = ".Net\build"

Main
ReleaseObjectMemory

Sub Main
    If installing Then
        PrepWscRegistrationSystem32
        PrepWscRegistrationSysWoW64
        PrepDllRegistration
        PrepFinalInstruction
        RunBatchFile
        CreateEventLogSource
    ElseIf uninstalling Then
        DeleteEventLogSource
        PrepDllRegistration
        PrepWscRegistrationSystem32
        PrepWscRegistrationSysWoW64
        PrepFinalInstruction
        RunBatchFile
        DeleteScriptletKeys
    End If
    DeleteBatchFile
End Sub

'prepare for registering components (.wsc files),
'for 32-bit or 64-bit, according to system bitness
Sub PrepWscRegistrationSystem32
    batchStream.WriteLine "echo."
    Dim file : For Each file In fso.GetFolder(componentFolder).Files
        If "wsc" = LCase(fso.GetExtensionName(file)) Then
            batchStream.WriteLine format(Array( _
                "echo %s %s %s" & _
                "%SystemRoot%\System32\regsvr32 %s /s ""%s""", _
                registerVerb, fso.GetFileName(file), vbCrLf, wscFlag, file _
            ))
        End If
    Next
End Sub

'prepare for registering components,
'for 32-bit apps on 64-bit systems
Sub PrepWscRegistrationSysWoW64
    If wow Then Exit Sub 'not applicable to 32-bit systems
    batchStream.WriteLine "echo."
    Dim file : For Each file In fso.GetFolder(componentFolder).Files
        If "wsc" = LCase(fso.GetExtensionName(file)) Then
            batchStream.WriteLine format(Array( _
                "echo %s %s for 32-bit apps %s" & _
                "%SystemRoot%\SysWow64\regsvr32 %s /s ""%s""", _
                registerVerb, fso.GetFileName(file), vbCrLf, wscFlag, file _
            ))
        End If
    Next
End Sub

'prepare for compiling and registering/unregistering the VBS extension
Sub PrepDllRegistration
    batchStream.WriteLine "echo."
    batchStream.WriteLine format(Array("cd ""%s""", buildFolder))
    Dim file : For Each file In fso.GetFolder(buildFolder).Files
        If "bat" = fso.GetExtensionName(file) Then
            batchStream.WriteLine format(Array("call ""%s"" %s", file.Name, dllFlag))
        End If
    Next
End Sub

Sub PrepFinalInstruction
    batchStream.WriteLine "echo."
    batchStream.WriteLine format(Array( _
        "echo Close this window to finish %s. & pause > nul", _
        setupNoun))
End Sub

Sub RunBatchFile
    batchStream.Close
    If inspectBatchFile Then
        sh.Run "notepad " & batchFile
        If vbCancel = MsgBox("Click OK to proceed with VBScript Utilities " & setupNoun & " after inspecting the batch file.", vbInformation + vbOKCancel + vbSystemModal, "Proceed? - " & WScript.ScriptName) Then DeleteBatchFile : ReleaseObjectMemory : WScript.Quit
    End If
    sh.Run "cmd /c " & batchFile,, synchronous
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

Sub DeleteBatchFile
    On Error Resume Next
        batchStream.Close
    On Error Goto 0
    If fso.FileExists(batchFile) Then
        fso.DeleteFile batchFile
    End If
End Sub

'Remove the registry keys associated with script components;
'regsvr32.exe may show a success message on unregister
'without removing the registry keys.
Sub DeleteScriptletKeys
    Dim keys : keys = Array( _
    "CLSID\{ADCEC089-30E1-11D7-86BF-00606744568C}", _
    "Wow6432Node\CLSID\{ADCEC089-30E1-11D7-86BF-00606744568C}", _
    "VBScripting.EventExample", _
    "CLSID\{ADCEC089-30DE-11D7-86BF-00606744568C}", _
    "Wow6432Node\CLSID\{ADCEC089-30DE-11D7-86BF-00606744568C}", _
    "Includer", _
    "VBScripting.Includer", _
    "CLSID\{ADCEC089-30E2-11D7-86BF-00606744568C}", _
    "Wow6432Node\CLSID\{ADCEC089-30E2-11D7-86BF-00606744568C}", _
    "VBScripting.KeyDeleter", _
    "CLSID\{ADCEC089-30DF-11D7-86BF-00606744568C}", _
    "Wow6432Node\CLSID\{ADCEC089-30DF-11D7-86BF-00606744568C}", _
    "VBScripting.StringFormatter", _
    "CLSID\{ADCEC089-30E0-11D7-86BF-00606744568C}", _
    "Wow6432Node\CLSID\{ADCEC089-30E0-11D7-86BF-00606744568C}", _
    "VBScripting.VBSApp", _
"")
    Dim i : For i = 0 To UBound(keys) - 1
        keyDeleter.DeleteKey keyDeleter.HKCR, keys(i)
    Next
End Sub

Sub ReleaseObjectMemory
    Set sa = Nothing
    Set sh = Nothing
    Set fso = Nothing
End Sub

Const synchronous = True
Const batchFile = "Setup.bat"
Const configFile = "Setup.config"
Const ForAppending = 8
Const CreateNew = True
Dim batchStream
Dim projectFolder, buildFolder, componentFolder
Dim installing, uninstalling, registerVerb, setupNoun
Dim wscFlag, dllFlag
Dim sa, sh, fso, reg
Dim include, format, keyDeleter
Dim wow
Dim inspectBatchFile

Sub Initialize
    Set sa = CreateObject("Shell.Application")
    Set sh = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set reg = GetObject("winmgmts:\\.\root\default:StdRegProv")

    'get config data
    inspectBatchFile = False
    On Error Resume Next
        Execute fso.OpenTextFile(configFile).ReadAll
    On Error Goto 0

    'convert relative paths to absolute paths
    projectFolder = fso.GetParentFolderName(WScript.ScriptFullName)
    sh.CurrentDirectory = projectFolder
    buildFolder = fso.GetAbsolutePathName(buildFolder_)
    componentFolder = fso.GetAbsolutePathName(componentFolder_)

    'initialize required project classes
    Set include = GetObject("script:" & componentFolder & "\Includer.wsc")
    Set format = GetObject("script:" & componentFolder & "\StringFormatter.wsc")
    Set keyDeleter = GetObject("script:" & componentFolder & "\KeyDeleter.wsc")
    With include
        .SetLibraryPath fso.GetAbsolutePathName("class")
        Execute .Read("PrivilegeChecker")
        Dim pc : Set pc = New PrivilegeChecker
        Execute .Read("WoWChecker")
        Set wow = New WoWChecker
    End With

    'look for /u on the command line
    With WScript.Arguments
        uninstalling = False
        Dim i : For i = 0 To .Count - 1
            If "/u" = LCase(.item(i)) Then
                uninstalling = True
            End If
        Next
        Dim setupFlag 'flag for restarting this script
        If uninstalling Then
            setupFlag = "/u"
            registerVerb = "Unregistering"
            setupNoun = "uninstalling"
            wscFlag = "/u"
            dllFlag = "/unregister"
            installing = False
        Else 'installing
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

    'prepare batchStream
    If fso.FileExists(batchFile) Then fso.DeleteFile batchFile
    Set batchStream = fso.OpenTextFile(batchFile, ForAppending, CreateNew)
    batchStream.WriteLine "@echo off & echo."
End Sub
