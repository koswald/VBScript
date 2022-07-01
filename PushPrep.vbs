'Main script included by PushPrep.hta

'If this .vbs file is started independently of the .hta, then launch the .hta
If IsEmpty( document ) Then
    Set sh = CreateObject( "WScript.Shell" )
    Set fso = CreateObject( "Scripting.FileSystemObject" )
    sh.CurrentDirectory = fso.GetParentFolderName( _
        WScript.ScriptFullName _
    )
    sh.Run "PushPrep.hta"
    WScript.Quit
End If

s = "Run Setup.vbs before launching PushPrep.hta."
On Error Resume Next
    Set incl = CreateObject("VBScripting.Includer")
    If Err Then
        MsgBox _
            "Err.Description: " & vbTab & Err.Description & vbLf & _
            "Hex(Err.Number): " & vbTab & Err.Number & _
            vbLf & vbLf & s, _
            vbInformation, "PushPrep.hta"
        Self.Close
    End If
On Error Goto 0

Dim sh 'WScript.Shell object
Dim fso 'Scripting.FileSystemObject
Dim format 'StringFormatter object
Dim suiteFolder 'string: folder where test suite scripts are located
Dim projectFolder 'string: root folder for this project
Dim suiteFilter 'string: filename filter for selecting integration test suites.
Dim caption 'string: MsgBox/PopUp title bar text.
Dim aDocGens 'array of strings: filespecs for code-comment-based documentation generators.
Dim aGits 'array of strings: common filespecs for Git bash and Git GUI executables.
Dim aDocs 'array of strings: filespecs for last-minute docs to update before a push.
Dim nextItem 'integer: current index of the prepItems array.
Dim settings 'integer: controls MsgBox/PopUp behaviour.
Dim prepItems 'array: list of prcedure (Sub) names to be called by window.SetTimeout.
Dim flagFile 'string: filename of a temp file used by Setup.vbs.
Dim versionLink 'web page with version info
Dim editor 'document editor
Dim powershell 'filespec of a pwsh.exe, if available; or just "powershell"
Const CreateNew = True 'for the OpenTextFile method.
Const Enter = 13 'window.event.keyCode for the Enter key
Const Esc = 27 'window.event.keyCode for the Esc key
Const synchronous = True 'for the Run method
Const hidden = 0 'for the Run method
Const VBScript = "VBScript" 'for the SetTimeout method
Const uninstallKey = "HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall"
Const bitMatch = 1, bitNoMatch = 0 'for bitwise comparisons
Const bitYes = 8, bitNo = 4, bitCancel = 2
Const gitFound = 1, gitLost = 0

Sub Window_OnLoad
    Dim app 'HTML application object reference
    Dim defaultSuiteFolder, defaultSuiteFilter, defaultDocGens, defaultGits, defaultDocs, defaultEditor 'strings
    Dim candidate 'a pwsh.exe filespec

    Set sh = CreateObject( "WScript.Shell" )
    Set fso = CreateObject( "Scripting.FileSystemObject" )
    Set app = document.getElementsByTagName( "application" )(0)
    document.Title = app.applicationName
    projectFolder = Replace(fso.GetParentFolderName(app.CommandLine), """", "")
    Set format = New StringFormatter

    defaultSuiteFolder = "spec\suite"
    defaultSuiteFilter = "TestLauncher"
    defaultDocGens = "examples\Generate-the-CSharp-docs.vbs | examples\Generate-the-VBScript-docs.vbs"
    defaultGits = "%ProgramFiles%\Git\cmd\git-gui.exe | %ProgramFiles%\Git\git-bash.exe | %LocalAppData%\Programs\Git\cmd\git-gui.exe | %LocalAppData%\Programs\Git\git-bash.exe"
    defaultDocs = "ChangeLog.md | ProjectInfo.vbs"
    defaultEditor = "notepad"

    With New Configurer
        powershell = .PowerShell

        If .Exists( "suite folder" ) Then
            suiteFolder = .Item( "suite folder" )
        Else suiteFolder = defaultSuiteFolder
        End If
        If .Exists( "suite filter" ) Then
            suiteFilter = .Item( "suite filter" )
        Else suiteFilter = defaultSuiteFilter
        End If
        If .Exists( "doc generators" ) Then
            aDocGens = .ToArray( .Item( "doc generators" ))
        Else aDocGens = .ToArray( defaultDocGens )
        End If
        If .Exists( "gits" ) Then
            aGits = .ToArray( .Item( "gits" ))
        Else aGits = .ToArray( defaultGits )
        End If
        If .Exists( "push docs" ) Then
            aDocs = .ToArray( .Item( "push docs" ))
        Else aDocs = .ToArray( defaultDocs )
        End If
        If .Exists( "editor" ) Then
            editor = .Item( "editor" )
        Else editor = defaultEditor
        End If

    End With

    prepItems = Array("" _
        , "UpdatePrePushDocs" _
        , "RunSetupUninstall" _
        , "StopScripts" _
        , "RunSetup" _
        , "RunTestSuites" _
        , "GenerateDocs" _
        , "OpenProgramsAndFeatures" _
        , "OpenGit" _
    )
    flagFile = "Setup.bat"

    caption = document.Title
    settings = vbYesNoCancel + vbInformation + vbDefaultButton2
    sh.CurrentDirectory = projectFolder
    versionLink = "https://github.com/koswald/VBScript/blob/master/ProjectInfo.vbs"
End Sub

Sub prepBtn_OnClick
    nextItem = 0
    AwaitNextItem
End Sub
Sub AwaitNextItem
    ClearFeedback
    nextItem = nextItem + 1
    If nextItem > UBound( prepItems ) Then
        nextItem = 0
        Exit Sub
    End If
    window.setTimeout prepItems( nextItem ), 1, VBScript
End Sub

Sub UpdatePrePushDocs
    Dim response 'integer: response to MsgBox
    Dim cmd 'string: Windows command
    Dim doc 'string: partial filespec
    If Not chkVersionChkBox.checked Then
        AwaitNextItem
        Exit Sub
    End If
    If reqConfirmChkBox.checked Then
        response = MsgBox( "Open selected pre-push docs for editing?", settings, caption )
    Else response = vbYes
    End If
    If vbCancel = response Then
        Exit Sub
    ElseIf vbNo = response Then
        AwaitNextItem
        Exit Sub
    End If
    For Each doc In aDocs
        If reqConfirmChkBox.checked Then
            response = MsgBox( "Edit " & doc & "?", settings, caption )
        Else response = vbYes
        End If
        If vbCancel = response Then
            Exit Sub
        ElseIf vbYes = response Then
            cmd = format( Array( _
                """%s"" ""%s\%s""", _
                editor, projectFolder, doc _
            ))
            sh.Run cmd, hidden
        End If
    Next
    AwaitNextItem
End Sub

Sub RunSetupUninstall
    Dim response
    If Not uninstallChkBox.checked Then
        AwaitNextItem
        Exit Sub
    End If
    If reqConfirmChkBox.checked Then
        response = MsgBox("Uninstall the VBScripting components and libraries, etc.?", settings, caption)
    Else response = vbYes
    End If
    If vbYes = response Then
        CreateFlagFile
        If Not UninstallFromProgramsAndFeatures Then
            UninstallDirectly
        End If
        nextItem = nextItem + 1
        window.setTimeout "AwaitSetupCompletion", 1000, VBScript
    ElseIf vbNo = response Then
        AwaitNextItem
    End If
End Sub
Sub UninstallDirectly
    If reqConfirmChkBox.checked Then
        sh.Run "wscript Setup.vbs /u"
    Else sh.Run "wscript Setup.vbs /u /s"
    End If
End Sub
Sub AwaitSetupCompletion
    If fso.FileExists(flagFile) Then
        Feedback "Waiting for Setup/Uninstall to finish.<br><br>After Setup/uninstall has finished, and after inspecting for errors, close the console window."
        window.setTimeout "AwaitSetupCompletion", 2000, VBScript
    Else window.setTimeout prepItems(nextItem), 1, VBScript
        ClearFeedback
    End If
End Sub
Function UninstallFromProgramsAndFeatures
    Dim key : key = format( Array( _
        "%s\VBScripting\UninstallString", uninstallKey _
    ))
    If Not reqConfirmChkBox.checked Then
        UninstallFromProgramsAndFeatures = False
        Exit Function
    End If
    On Error Resume Next
        sh.Run sh.RegRead(key)
        If Err Then
            UninstallFromProgramsAndFeatures = False
        Else UninstallFromProgramsAndFeatures = True
        End If
    On Error Goto 0
End Function
Sub CreateFlagFile
    If Not fso.FileExists(flagFile) Then
        On Error Resume Next
            fso.CreateTextFile flagFile, CreateNew
        On Error Goto 0
    End If
End Sub

'When one of the .NET extension objects is in use, and it is desired to recompile the class file, it is necessary first to stop the instance of the script that is using the object.
Sub StopScripts
    Dim response
    If Not stopScriptsChkBox.checked Then
        AwaitNextItem
        Exit Sub
    End If
    If reqConfirmChkBox.checked Then
        response = MsgBox( _
            "Stop all instances of wscript.exe?" & vbLf & vbLf & _
            "If any processes are using the project module or library files, then the C# compiler will not be able to recreate those files.", _
            settings, caption)
    Else response = vbYes
    End If
    If vbYes = response Then
        KillProcessesByName( "wscript.exe" )
    ElseIf vbCancel = response Then
        Exit Sub
    End If
    AwaitNextItem
End Sub
Sub KillProcessesByName(processName)
    Dim id, IDs
    With New WMIUtility
        IDs = .GetProcessIDsByName(processName)
        For Each id In IDs
            .TerminateProcessById(id)
        Next
    End With
End Sub

Sub RunSetup
    Dim response
    If Not runSetupChkBox.checked Then
        AwaitNextItem
        Exit Sub
    End If
    If reqConfirmChkBox.checked Then
        response = MsgBox("Run Setup?", settings, caption)
    Else response = vbYes
    End If
    If vbYes = response Then
        CreateFlagFile
        sh.Run "Setup.vbs"
        nextItem = nextItem + 1
        window.setTimeout "AwaitSetupCompletion", 1000, VBScript
    ElseIf vbNo = response Then
        AwaitNextItem
    End If
End Sub

Sub RunTestSuites
    Dim file, path
    If Not runTestsChkBox.checked Then
        AwaitNextItem
        Exit Sub
    End If
    Feedback "Waiting for tests to complete.<br><br>After each test suite finishes, and after inspecting for errors, close the console window(s)."
    path = format( Array( _
        "%s\%s", projectFolder, suiteFolder _
    ))
    For Each file In fso.GetFolder( path ).Files
        If bitCancel And SuiteResult( file ) Then
            ClearFeedback
            Exit Sub
        End If
    Next
    ClearFeedback
    AwaitNextItem
End Sub
Function SuiteResult( suiteCandidate )
    Dim response 'integer: actual or implied user response
    Dim suite 'file object representing the suite script file
    If InStr( suiteCandidate.Name, suiteFilter ) Then
        Set suite = suiteCandidate
    Else SuiteResult = bitNoMatch
        Exit Function
    End If
    If reqConfirmChkBox.checked Then
        response = MsgBox(format(Array("Run %s?", suite.Name)), settings, caption)
    Else response = vbYes
    End If
    If vbYes = response Then
        sh.Run format( Array( """%s""", suite.Path )),, synchronous
        SuiteResult = bitYes Or bitMatch
    ElseIf vbCancel = response Then
        SuiteResult = bitCancel Or bitMatch
    Else SuiteResult = bitNo Or bitMatch
    End If
End Function

Sub GenerateDocs
    Dim i, response, item
    If Not generateDocsChkBox.checked Then
        AwaitNextItem
        Exit Sub
    End If
    For i = 0 To UBound(aDocGens)
        item = fso.GetAbsolutePathName(aDocGens(i))
        If reqConfirmChkBox.checked Then
            response = MsgBox(format(Array("Run %s?", item)), settings, caption)
        Else response = vbYes
        End If
        If vbYes = response Then
            sh.Run format(Array("""%s""", item)),, synchronous
        ElseIf vbCancel = response Then
            Exit Sub
        End If
    Next
    AwaitNextItem
End Sub

Sub OpenProgramsAndFeatures
    Dim response
    If Not openProgramsAndFeaturesChkBox.checked Then
        AwaitNextItem
        Exit Sub
    End If
    If reqConfirmChkBox.checked Then
        response = MsgBox("Open Programs and features (legacy GUI)?", settings, caption)
    Else response = vbYes
    End If
    If vbYes = response Then
        sh.Run "control /name Microsoft.ProgramsAndFeatures"
    ElseIf vbCancel = response Then
        Exit Sub
    End If
    If reqConfirmChkBox.checked Then
        response = MsgBox("Open Programs and features (Windows 10 GUI)?", settings, caption)
    Else response = vbYes
    End If
    If vbYes = response Then
        sh.Run "ms-settings:appsfeatures"
    ElseIf vbCancel = response Then
        Exit Sub
    End If
    AwaitNextItem
End Sub

Sub OpenGit
    Dim i 'integer
    Dim result 'integer: response to MsgBox
    Dim gitWasFound 'boolean: indicates whether any Git executables were found.
    gitWasFound = False
    If Not openGitChkBox.checked Then
        Exit Sub
    End If
    For i = 0 To UBound(aGits)
        result = GitResult(aGits(i))
        If (bitYes And result) _
        Or (bitCancel And result) Then
            Exit Sub
        End If
        If gitFound And result Then gitWasFound = True
    Next
    If gitWasFound Then
        Exit Sub
    End If
    MsgBox "Couldn't find a Git executable.", vbInformation, caption
End Sub
Function GitResult(git)
    Dim response
    If Not fso.FileExists(Expand(git)) Then
        GitResult = gitLost
        Exit Function
    End If
    If reqConfirmChkBox.checked Then
        response = MsgBox(format(Array("Run %s?", git)), settings, caption)
    Else response = vbYes
    End If
    If vbYes = response Then
        sh.Run format(Array("""%s""", git))
        GitResult = bitYes Or gitFound
    ElseIf vbCancel = response Then
        GitResult = bitCancel Or gitFound
    Else GitResult = bitNo Or gitFound
    End If
End Function

Function Expand(str)
    Expand = sh.ExpandEnvironmentStrings(str)
End Function

Sub selectAllChkBox_OnClick
    Dim input, inputs
    Set inputs = document.getElementsByTagName( "input" )
    For Each input In inputs
        CheckOrUncheckPrepItem input, selectAllChkBox.checked
    Next
End Sub
Sub CheckOrUncheckPrepItem(element, newStatus)
    If "selectAllChkBox" = element.id _
    Or "reqConfirmChkBox" = element.id _
    Or Not "checkbox" = element.type Then
        Exit Sub
    End If
    element.checked = newStatus
End Sub

Sub Document_OnKeyUp
    If Enter = window.event.keyCode Then
        prepBtn_OnClick
    ElseIf Esc = window.event.keyCode Then
        Self.Close
    End If
End Sub

Sub Feedback(str)
    Unhide info
    info.innerHTML = str
End Sub
Sub ClearFeedback
    Hide info
    info.innerHTML = ""
End Sub
Sub Hide(element)
    element.style.display = "none"
End Sub
Sub Unhide(element)
    element.style.display = "block"
End Sub

Sub Window_OnUnload
    Set sh = Nothing
    Set fso = Nothing
    Set format = Nothing
End Sub
