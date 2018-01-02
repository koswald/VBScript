
'Launch the test runner for 32-bit (x86) tests

Option Explicit : Initialize

Call Main

Sub Main

    'specify the file types
    testRunner.SetSpecPattern "*.spec.wow.vbs"
 
    'specify the folder containing the tests; path is relative to this script
    testRunner.SetSpecFolder "..\..\spec\wow"

    'handle command-line arguments, if any
    With WScript.Arguments
        If .Count Then

            'if it is desired to run just a single test file, pass it in on the
            'command line, using a relative path, relative to the spec folder
            testRunner.SetSpecFile .item(0)

            'get the runCount from the command-line, arg #2, if specified
            If .Count > 1 Then testRunner.SetRunCount .item(1)
       End If
    End With

    'run the tests
    testRunner.Run
End Sub

Const privilegesElevated = True
Const privilegesNotElevated = False
Dim testRunner

Sub Initialize
    With CreateObject("includer")
        Execute .read("VBSTestRunner")
        Execute .read("WoWChecker")
        Execute .read("VBSApp")
    End With
    Set testRunner = New VBSTestRunner
    Dim wow : Set wow = New WoWChecker
    Dim app : Set app = New VBSApp

    'if required, restart the script with the x86 cscript.exe
    If (Not wow.isWoW) Or (Not "cscript.exe" = app.GetExe) Then
        app.SetUserInteractive False
        app.RestartWith "%SystemRoot%\SysWoW64\cscript.exe", "/k", privilegesNotElevated
    End If
End Sub
