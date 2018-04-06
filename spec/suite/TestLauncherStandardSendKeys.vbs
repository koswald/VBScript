'Launch the test runner for standard tests
Option Explicit
Initialize
Main
Sub Main
    testRunner.SetSpecPattern "*.spec.sk.vbs"
    testRunner.SetSpecFolder ".."
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
    End With
    Set testRunner = New VBSTestRunner
    Dim app : Set app = New VBSApp
    If Not "cscript.exe" = app.GetHost Then
        app.SetUserInteractive False
        app.RestartWith "cscript.exe", "/k", privilegesNotElevated
    End If
End Sub
