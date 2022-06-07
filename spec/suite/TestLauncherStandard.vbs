'Launch the test runner for standard tests
Option Explicit
Dim testRunner

Initialize
Main

Sub Initialize
    Dim incl
    Set incl = CreateObject( "VBScripting.Includer" )
    Execute incl.Read( "VBSTestRunner" )
    Set testRunner = New VBSTestRunner
    Execute incl.Read( "VBSApp" )
    With New VBSApp
        .RestartUsing "cscript.exe", .DoNotExit, .DoNotElevate
    End With
End Sub

Sub Main
    testRunner.SetSpecPattern "*.spec.vbs | *.spec.elev+std.vbs | *.spec.wsf"
    testRunner.SetSpecFolder ".."

    'If it is desired to run just a single test file, pass it in on the command line, using a relative path, relative to the spec folder. Also, get the runCount from the command-line, arg #2, if specified

    With WScript.Arguments
        If .Count > 0 Then
            testRunner.SetSpecFile .item(0)
        End If
        If .Count > 1 Then
            testRunner.SetRunCount .item(1)
        End If
    End With

    'Run the test suite

    testRunner.Run
End Sub
