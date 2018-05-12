
'Setup the VBScript utilities

'Registers the windows script component, Includer.wsc,
'other project components, and builds the VBScript
'extension libraries (.dll files).

'The User Account Control dialog will open
'to verify elevation of privileges.

'Use /u to uninstall

Option Explicit

componentFolder_ = "class\wsc"
buildFolder_ = ".Net\build"

Initialize
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
        ProgramsAndFeaturesEntry
    ElseIf uninstalling Then
        If Not silent Then
            If vbCancel = MsgBox("Uninstall VBScripting utility classes and extensions?", vbOKCancel + vbInformation + vbSystemModal + vbDefaultButton2, WScript.ScriptName) Then
                DeleteBatchFile
                Exit Sub
            End If
        End If
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
        If vbCancel = MsgBox(format(Array("Click OK to proceed with %s the VBScript Utilities after inspecting the batch file.", setupNoun)), vbInformation + vbOKCancel + vbSystemModal, "Proceed? - " & WScript.ScriptName) Then DeleteBatchFile : ReleaseObjectMemory : WScript.Quit
    End If
    sh.Run format(Array("cmd /c %s", batchFile)),, synchronous
End Sub

Sub CreateEventLogSource
    On Error Resume Next
        Dim va : Set va = CreateObject("VBScripting.Admin")
        va.CreateEventSource va.EventSource
        Set va = Nothing
    On Error Goto 0
End Sub

Sub ProgramsAndFeaturesEntry
    Const HKLM = &H80000002
    Const uninstKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\VBScripting"
    Dim InstallLocation : InstallLocation = fso.GetParentFolderName(WScript.ScriptFullName)
    reg.CreateKey HKLM, uninstKey
    reg.SetStringValue HKLM, uninstKey, "DisplayName", "VBScripting utility classes and extensions"
    reg.SetDWORDValue HKLM, uninstKey, "NoRemove", 0
    reg.SetStringValue HKLM, uninstKey, "UninstallString", format(Array("wscript ""%s\Setup.vbs"" /u", InstallLocation))
    reg.SetDWORDValue HKLM, uninstKey, "NoModify", 1
    reg.SetStringValue HKLM, uninstKey, "ModifyPath", ""
    reg.SetDWORDValue HKLM, uninstKey, "NoRepair", 1
    reg.SetStringValue HKLM, uninstKey, "HelpLink", "https://github.com/koswald/VBScript"
    reg.SetStringValue HKLM, uninstKey, "InstallLocation", InstallLocation
    reg.SetDWORDValue HKLM, uninstKey, "EstimatedSize", 1500 'kilobytes
    reg.SetExpandedStringValue HKLM, uninstKey, "DisplayIcon", "%SystemRoot%\System32\wscript.exe,1"
    reg.SetStringValue HKLM, uninstKey, "Publisher", "Karl Oswald"
    reg.SetStringValue HKLM, uninstKey, "HelpTelephone", ""
    reg.SetStringValue HKLM, uninstKey, "Contact", ""
    reg.SetStringValue HKLM, uninstKey, "UrlInfoAbout", ""
    reg.SetStringValue HKLM, uninstKey, "DisplayVersion", ""
    reg.SetStringValue HKLM, uninstKey, "Comments", ""
    reg.SetStringValue HKLM, uninstKey, "Readme", InstallLocation & "\ReadMe.md"
    reg.SetStringValue HKLM, uninstKey, "InstallDate", "" ' [YYYYMMDD]
    reg.SetDWORDValue HKLM, uninstKey, "Version", 0
    reg.SetDWORDValue HKLM, uninstKey, "VersionMajor", 0
    reg.SetDWORDValue HKLM, uninstKey, "VersionMinor", 0
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
    Dim keys : keys = Array("" _
        , "Software\Classes\CLSID\{ADCEC089-30E1-11D7-86BF-00606744568C}" _
        , "Software\Classes\Wow6432Node\CLSID\{ADCEC089-30E1-11D7-86BF-00606744568C}" _
        , "Software\Classes\VBScripting.EventExample" _
        , "Software\Classes\CLSID\{ADCEC089-30DE-11D7-86BF-00606744568C}" _
        , "Software\Classes\Wow6432Node\CLSID\{ADCEC089-30DE-11D7-86BF-00606744568C}" _
        , "Software\Classes\Includer" _
        , "Software\Classes\VBScripting.Includer" _
        , "Software\Classes\CLSID\{ADCEC089-30E2-11D7-86BF-00606744568C}" _
        , "Software\Classes\Wow6432Node\CLSID\{ADCEC089-30E2-11D7-86BF-00606744568C}" _
        , "Software\Classes\VBScripting.KeyDeleter" _
        , "Software\Classes\CLSID\{ADCEC089-30DF-11D7-86BF-00606744568C}" _
        , "Software\Classes\Wow6432Node\CLSID\{ADCEC089-30DF-11D7-86BF-00606744568C}" _
        , "Software\Classes\VBScripting.StringFormatter" _
        , "Software\Classes\CLSID\{ADCEC089-30E0-11D7-86BF-00606744568C}" _
        , "Software\Classes\Wow6432Node\CLSID\{ADCEC089-30E0-11D7-86BF-00606744568C}" _
        , "Software\Classes\VBScripting.VBSApp" _
        , "Software\Microsoft\Windows\CurrentVersion\Uninstall\VBScripting" _
    )
    Dim i : For i = 1 To UBound(keys)
        keyDeleter.DeleteKey keyDeleter.HKLM, keys(i)
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
Dim installing, uninstalling, registerVerb, setupNoun, silent
Dim wscFlag, dllFlag
Dim sa, sh, fso, reg
Dim include, format, keyDeleter
Dim wow
Dim inspectBatchFile
Dim componentFolder_, buildFolder_

Sub Initialize
    Set sa = CreateObject("Shell.Application")
    Set sh = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set reg = GetObject("winmgmts:\\.\root\default:StdRegProv")

    'convert relative paths to absolute paths
    projectFolder = fso.GetParentFolderName(WScript.ScriptFullName)
    sh.CurrentDirectory = projectFolder
    buildFolder = fso.GetAbsolutePathName(buildFolder_)
    componentFolder = fso.GetAbsolutePathName(componentFolder_)

    'get config data
    inspectBatchFile = False
    On Error Resume Next
        Execute fso.OpenTextFile(configFile).ReadAll
    On Error Goto 0

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

    'look for arguments on the command line
    With WScript.Arguments
        uninstalling = False
        silent = False
        Dim silentFlag : silentFlag = ""
        Dim i : For i = 0 To .Count - 1
            If "/u" = LCase(.item(i)) Then
                uninstalling = True
            ElseIf "/s" = LCase(.item(i)) Then
                silent = True
                silentFlag = "/s"
            End If
        Next
        Dim setupFlag
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
            setupNoun = "setting up"
            wscFlag = ""
            dllFlag = ""
            installing = True
        End If
    End With
    If Not pc Then

        'restart this script to elevate privileges
        Dim restartArgs : restartArgs = format(Array( _
            "/c cd ""%s"" & start wscript ""%s"" %s %s", _
            projectFolder, WScript.ScriptFullName, setupFlag, silentFlag _
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
