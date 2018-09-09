'Launch the test runner for .dll libraries
Option Explicit
Initialize
Main
Sub Main
    testRunner.SetSpecPattern "*.spec.vbs"
    testRunner.SetSpecFolder "..\dll"
    With WScript.Arguments
        If .Count Then
            'if it is desired to run just a single test file, pass it in on the
            'command line, using a relative path, relative to the spec folder
            testRunner.SetSpecFile .item(0)
            'get the runCount from the command-line, arg #2, if specified
            If .Count > 1 Then testRunner.SetRunCount .item(1)
       End If
    End With
    testRunner.Run
End Sub

Const privilegesElevated = True
Const privilegesNotElevated = False
Dim testRunner
Sub Initialize
    With CreateObject("VBScripting.Includer")
        Execute .read("VBSTestRunner")
        Execute .read("VBSApp")
        Execute .read("WoWChecker")
    End With
    Set testRunner = New VBSTestRunner
    Set app = New VBSApp
    Set wow = New WoWChecker
    If Not "cscript.exe" = app.GetHost Or Not wow.IsWoW Then
        app.SetUserInteractive False
        app.RestartWith "%SystemRoot%\SysWoW64\cscript.exe", "/k", privilegesNotElevated
    End If

    Dim app, wow
End Sub
