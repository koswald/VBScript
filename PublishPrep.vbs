'Semi-automate selected tasks commonly performed
'before pushing to the remote branch.
'This script must be in the root project folder.
Option Explicit
docGenerators = Array("" _
    , "examples\Generate-the-CSharp-docs.vbs" _
    , "examples\Generate-the-VBScript-docs.vbs" _
)
gits = Array("" _
    , "%ProgramFiles%\Git\cmd\git-gui.exe" _
    , "%ProgramFiles%\Git\git-bash.exe" _
)
componentFolder = "class\wsc"
suiteFolder = "spec\suite"
suiteFilter = "TestLauncher"

Initialize
Main
Quit

Sub Main
    RunSetupUninstall
    StopScripts
    RunSetup
    RunTestSuites
    GenerateDocs
    OpenGit
End Sub

Sub RunSetupUninstall
    response = MsgBox("Run Setup.vbs /u?", settings, caption)
    If vbYes = response Then
        sh.Run "Setup.vbs /u",, synchronous
    ElseIf vbCancel = response Then
        Quit
    End If
End Sub
Sub StopScripts
    response = MsgBox("Stop all instances of wscript.exe?" & vbLf & vbLf & "Some C# compiler errors may be prevented by stopping scripts that may be using the .NET libraries." & vbLf & "This script would need to be restarted.", settings, caption)
    If vbYes = response Then
        Execute includer.Read("WMIUtility")
        Dim wmi : Set wmi = New WMIUtility
        Dim ids : ids = wmi.GetProcessIDsByName("wscript.exe")
        Dim i
        For i = 0 To UBound(ids)
            wmi.TerminateProcessById(ids(i))
        Next
    ElseIf vbCancel = response Then
        Quit
    End If
End Sub
Sub RunSetup
    response = MsgBox("Run Setup.vbs?", settings, caption)
    If vbYes = response Then
        sh.Run "Setup.vbs",, synchronous
    ElseIf vbCancel = response Then
        Quit
    End If
End Sub
Sub RunTestSuites
    Dim file
    For Each file In fso.GetFolder(projectFolder & "\" & suiteFolder).Files
        If InStr(File.Name, suiteFilter) Then
            response = MsgBox("Run " & File.Name & "?", settings, caption)
            If vbYes = response Then
                sh.Run "cmd /k cscript """ & File.Path & """",, synchronous
            ElseIf vbCancel = response Then
                Quit
            End If
        End If
    Next
End Sub
Sub GenerateDocs
    Dim i
    For i = 1 To UBound(docGenerators)
        response = MsgBox("Run " & docGenerators(i) & "?", settings, caption)
        If vbYes = response Then
            sh.Run docGenerators(i),, synchronous
        ElseIf vbCancel = response Then
            Quit
        End If
    Next
End Sub
Sub OpenGit
    Dim i, gitFound : gitFound = False
    For i = 1 To UBound(gits)
        If fso.FileExists(sh.ExpandEnvironmentStrings(gits(i))) Then
            gitFound = True
            response = MsgBox("Run " & gits(i) & "?", settings, caption)
            If vbYes = response Then
                sh.Run """" & gits(i) & """"
                Exit Sub
            ElseIf vbCancel = response Then
                Quit
            End If
        End If
    Next
    If Not gitFound Then
        MsgBox "Couldn't find a Git executable.", vbInformation, caption
    End If
End Sub

Const withElevatedPrivileges = "runas"
Const synchronous = True
Const timedOut = -1
Dim sh, fso, sa, format, includer, pc
Dim suites, docGenerators, gits
Dim response
Dim caption, settings
Dim componentFolder, projectFolder
Dim suiteFolder, suiteFilter

Sub Initialize
    Set sh = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set sa = CreateObject("Shell.Application")
    projectFolder = fso.GetParentFolderName(WScript.ScriptFullName)
    Set includer = GetObject("script:" & projectFolder & "\" & componentFolder & "\Includer.wsc")
    includer.SetLibraryPath projectFolder & "\class"    
    caption = WScript.ScriptName
    settings = vbYesNoCancel + vbInformation + vbSystemModal + vbDefaultButton2
End Sub
Sub Quit
    Set sh = Nothing
    Set fso = Nothing
    Set sa = Nothing
    Set format = Nothing
    Set includer = Nothing
    Set pc = Nothing
    WScript.Quit
End Sub
