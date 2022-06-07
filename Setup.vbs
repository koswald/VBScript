'Setup the VBScripting utilities

'Registers or unregisters the windows script component (.wsc) files.
'Compiles and registers or unregisters the VBScripting extension .dll libraries.
'Creates or removes the VBScripting event log source.

'Use /u to unregister/remove.

'The User Account Control dialog will open
'to verify elevation of privileges.

Option Explicit
Dim batchFile 'string: name of the setup batch file
Dim batchStream 'text stream object for creating and writing to the setup batch file
Dim configFile 'string: name of the .config file associated with this file
Dim installing, uninstalling 'booleans
Dim registerVerb ' "Registering" or "Unregistering"
Dim setupVerbal ' "setting up" or "uninstalling"
Dim silent 'boolean: True for non-interactive/silent setup
Dim visibility 'integer: for the Run method: hidden (0) for non-interactive setup
Dim wscFlag 'regsvr32.exe argument(s) for un/registering a .wsc
Dim dllFlag '.bat file command-line argument for un/registering a .dll
Dim sa 'Shell.Application object
Dim sh 'WScript.Shell object
Dim fso 'Scripting.FileSystemObject object
Dim include 'Includer object for dependency management
Dim format 'a string formatter object
Dim keyDeleter 'an object that can delete registry keys
Dim wow 'object for checking system bitness
Dim inspectBatchFile 'boolean
Dim projectFolder 'VBScripting project root folder
Dim buildFolder 'path to scripts that compile the extension library (.dll) files
Dim componentFolder 'path to Windows Script Component (.wsc) files
Dim projInfo 'object containing project information
Const synchronous = True 'for the Run method
Const ForWriting = 2 'for the OpenTextFile method
Const CreateNew = True 'for the OpenTextFile method
Const hidden = 0, normal = 1 'for the Run method

Initialize
Main
ReleaseObjectMemory

Sub Initialize
    Dim pc 'a PrivilegeChecker object
    Dim i 'integer
    Dim silentFlag 'string: command-line argument: "/s" when re/starting this script non-interactively
    Dim restartArgs 'string: a series of command-line arguments 
    Dim appData 'string: the project folder within %AppData%
    Dim setupFlag '/u if uninstalling
    
    Set sa = CreateObject( "Shell.Application" )
    Set sh = CreateObject( "WScript.Shell" )
    Set fso = CreateObject( "Scripting.FileSystemObject" )

    'relative paths => absolute paths
    projectFolder = fso.GetParentFolderName(WScript.ScriptFullName)
    sh.CurrentDirectory = projectFolder
    buildFolder = fso.GetAbsolutePathName(".Net\build")
    componentFolder = fso.GetAbsolutePathName("class\wsc")

    'get config data
    configFile = "Setup.config"
    inspectBatchFile = False 'in case the .config file can't be read
    On Error Resume Next
        Execute fso.OpenTextFile(configFile).ReadAll
    On Error Goto 0

    'instantiate Windows Script Components
    Set include = GetObject("script:" & componentFolder & "\Includer.wsc")
    Set format = GetObject("script:" & componentFolder & "\StringFormatter.wsc")
    Set keyDeleter = GetObject("script:" & componentFolder & "\KeyDeleter.wsc")

    'instantiate other project objects
    With include
        .SetLibraryPath fso.GetAbsolutePathName( "class" )
        Execute .Read( "PrivilegeChecker" )
        Set pc = New PrivilegeChecker
        Execute .Read( "WoWChecker" )
        Set wow = New WoWChecker
        Execute .Read("..\ProjectInfo.vbs")
        Set projInfo = New ProjectInfo
    End With

    'get command line arguments
    uninstalling = False
    silent = False
    visibility = normal
    silentFlag = ""
    With WScript.Arguments
        For i = 0 To .Count - 1
            If "/u" = LCase( .item( i )) Then
                uninstalling = True
            ElseIf "/s" = LCase( .item( i )) Then
                silent = True
                silentFlag = "/s"
                visibility = hidden
            End If
        Next
    End With
    If uninstalling Then
        setupFlag = "/u"
        registerVerb = "Unregistering"
        setupVerbal = "uninstalling"
        wscFlag = "/u /n"
        dllFlag = "/unregister"
        installing = False
    Else 'installing
        setupFlag = ""
        registerVerb = "Registering"
        setupVerbal = "setting up"
        wscFlag = ""
        dllFlag = ""
        installing = True
    End If

    If Not pc Then
        'Privileges are not elevated, so...
        'Start another instance of this script but with elevated privileges, retaining the command-line arguments, and then exit/quit the first instance. The User Account Control dialog will open.
        restartArgs = format( Array( _
            """%s"" %s %s", _
            WScript.ScriptFullName, setupFlag, silentFlag _
        ))
        sa.ShellExecute "wscript", restartArgs,, "runas"
        ReleaseObjectMemory
        WScript.Quit
    End If

    'create the batch file and open it for writing
    batchFile = "Setup.bat"
    If fso.FileExists( batchFile ) Then fso.DeleteFile batchFile
    Set batchStream = fso.OpenTextFile( batchFile, ForWriting, CreateNew )
    batchStream.WriteLine "@echo off & echo."

    '%AppData%
    appData = "%AppData%\VBScripting"
    appData = sh.ExpandEnvironmentStrings( appData )
    If Not fso.FolderExists( appData ) Then
        fso.CreateFolder appData
    End If
End Sub

Sub Main
    Dim m, i, s 'MsgBox args
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
            m = "Uninstall VBScripting utility classes and extensions?"
            i = vbOKCancel + vbInformation + vbSystemModal + vbDefaultButton2
            s = WScript.ScriptName
            If vbCancel = MsgBox( m, i, s ) Then
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
        DeleteSelectedKeys
    End If
    DeleteBatchFile
End Sub

'prepare to register .wsc files for 32-bit or 64-bit, according to system bitness
Sub PrepWscRegistrationSystem32
    Dim file 'file object
    batchStream.WriteLine "echo."
    For Each file In fso.GetFolder( componentFolder ).Files
        If "wsc" = LCase(fso.GetExtensionName( file.Name )) Then
            batchStream.WriteLine format( Array( _
                "echo %s %s" & vbCrLf & _
                "%SystemRoot%\System32\regsvr32 %s /s /i:""%s"" scrobj.dll", _
                registerVerb, file.Name, _
                wscFlag, file.Path _
            ))
        End If
    Next
End Sub

'prepare to register .wsc files for 32-bit apps on 64-bit systems
Sub PrepWscRegistrationSysWoW64
    Dim file 'file object
    If wow Then Exit Sub 'not applicable to 32-bit systems
    batchStream.WriteLine "echo."
    For Each file In fso.GetFolder( componentFolder ).Files
        If "wsc" = LCase( fso.GetExtensionName( file.Name )) Then
            batchStream.WriteLine format( Array( _
                "echo %s %s for 32-bit apps %s" & _
                "%SystemRoot%\SysWow64\regsvr32 %s /s /i:""%s"" scrobj.dll", _
                registerVerb, file.Name, vbCrLf, wscFlag, file _
            ))
        End If
    Next
End Sub

'prepare to compile and register .dll files
Sub PrepDllRegistration
    Dim file 'file object
    batchStream.WriteLine "echo."
    batchStream.WriteLine format( Array( _
        "cd ""%s""", buildFolder _
    ))
    For Each file In fso.GetFolder( buildFolder ).Files
        If "bat" = fso.GetExtensionName( file.Name ) Then
            batchStream.WriteLine format( Array( _
                "call ""%s"" %s", file.Name, dllFlag _
            ))
        End If
    Next
End Sub

Sub PrepFinalInstruction
    batchStream.WriteLine "echo."
    If silent Then Exit Sub
    batchStream.WriteLine format( Array( _
        "echo Close this window to finish %s. & pause > nul", _
        setupVerbal _
    ))
End Sub

Sub RunBatchFile
    Dim m, i, s 'MsgBox args
    batchStream.Close
    If inspectBatchFile And Not silent Then
        sh.Run "notepad """ & batchFile & """"
        'another opt out
        m = "Click OK to proceed with %s the "
        m = m & "VBScripting Utilities after "
        m = m & "inspecting the batch file."
        m = format( Array( m, setupVerbal ))
        i = vbInformation + vbOKCancel + vbSystemModal
        s = WScript.ScriptName
        If vbCancel = MsgBox( m, i, s ) Then
            DeleteBatchFile
            ReleaseObjectMemory
            WScript.Quit
        End If
    End If
    sh.Run format( Array( _
        "cmd /c %s", batchFile _
    )), visibility, synchronous
End Sub

Sub CreateEventLogSource
    On Error Resume Next
        With CreateObject( "VBScripting.Admin" )
            .CreateEventSource .EventSource
        End With
    On Error Goto 0
End Sub

Sub ProgramsAndFeaturesEntry
    Dim InstallLocation 'string: the project root folder
    Dim now_ 'variant subtype Date: a moment in time
    Dim size 'size in Kb; Windows GUIs typically convert this to Mb
    Dim reg 'StdRegProv object
    Const HKLM = &H80000002
    Const uninstKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\VBScripting"
    InstallLocation = fso.GetParentFolderName(WScript.ScriptFullName)
    now_ = Now
    size = fso.GetFolder( projectFolder ).Size/1024 'bytes ==> Kb
    Set reg = GetObject( "winmgmts:\\.\root\default:StdRegProv" )
    reg.CreateKey HKLM, uninstKey
    reg.SetStringValue HKLM, uninstKey, "DisplayName", "VBScripting Utility Classes and Extensions"
    reg.SetDWORDValue HKLM, uninstKey, "NoRemove", 0
    reg.SetStringValue HKLM, uninstKey, "UninstallString", format( Array( _
        "wscript ""%s\Setup.vbs"" /u", InstallLocation _
    ))
    reg.SetDWORDValue HKLM, uninstKey, "NoModify", 1
    reg.SetStringValue HKLM, uninstKey, "ModifyPath", ""
    reg.SetDWORDValue HKLM, uninstKey, "NoRepair", 0
    reg.SetStringValue HKLM, uninstKey, "RepairPath", format( Array("wscript ""%s\Setup.vbs""", InstallLocation )) '""
    reg.SetStringValue HKLM, uninstKey, "HelpLink", "https://github.com/koswald/VBScript"
    reg.SetStringValue HKLM, uninstKey, "InstallLocation", InstallLocation
    reg.SetDWORDValue HKLM, uninstKey, "EstimatedSize", size
    reg.SetExpandedStringValue HKLM, uninstKey, "DisplayIcon", "%SystemRoot%\System32\wscript.exe,2"
    reg.SetStringValue HKLM, uninstKey, "Publisher", "Karl Oswald"
    reg.SetStringValue HKLM, uninstKey, "HelpTelephone", ""
    reg.SetStringValue HKLM, uninstKey, "Contact", ""
    reg.SetStringValue HKLM, uninstKey, "UrlInfoAbout", ""
    reg.SetStringValue HKLM, uninstKey, "Comments", ""
    reg.SetStringValue HKLM, uninstKey, "Readme", InstallLocation & "\ReadMe.md"
    reg.SetStringValue HKLM, uninstKey, "InstallDate", 10000 * Year( now_ ) + 100 * Month( now_ ) + Day( now_ ) 'YYYYMMDD
    reg.SetStringValue HKLM, uninstKey, "DisplayVersion", projInfo.MajorVersion & "." & projInfo.MinorVersion & "." & projInfo.MicroVersion
    reg.SetDWORDValue HKLM, uninstKey, "VersionMajor", projInfo.MajorVersion
    reg.SetDWORDValue HKLM, uninstKey, "VersionMinor", projInfo.MinorVersion
End Sub

Sub DeleteEventLogSource
    On Error Resume Next
        With CreateObject( "VBScripting.Admin" )
            .DeleteEventSource .EventSource
        End With
    On Error Goto 0
End Sub

Sub DeleteBatchFile
    On Error Resume Next
        batchStream.Close
    On Error Goto 0
    If fso.FileExists( batchFile ) Then
        fso.DeleteFile batchFile
    End If
End Sub

Sub DeleteSelectedKeys
    Dim keys 'array of strings: registry keys
    Dim i 'integer
    keys = Array( "" _
        , "Software\Microsoft\Windows\CurrentVersion\Uninstall\VBScripting" _
    )
    For i = 1 To UBound( keys )
        keyDeleter.DeleteKey keyDeleter.HKLM, keys( i )
    Next
End Sub

Sub ReleaseObjectMemory
    Set sa = Nothing
    Set sh = Nothing
    Set fso = Nothing
End Sub
