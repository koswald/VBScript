'Launch the test runner for .dll libraries
Option Explicit
Dim testRunner

Initialize
Main

Sub Initialize
    Dim wow
    Dim incl
    Dim cscriptX86
    cscriptX86 = "%SystemRoot%\SysWoW64\cscript.exe"
    Set incl = CreateObject( "VBScripting.Includer" )
    Execute incl.Read( "VBSTestRunner" )
    Set testRunner = New VBSTestRunner
    Execute incl.Read( "WoWChecker" )
    Set wow = New WoWChecker
    Execute incl.Read( "VBSApp" )
    With New VBSApp
        If "cscript.exe" = .GetExe _
        And wow.IsWoW Then
            WScript.StdOut.WriteLine "Using the 32-bit cscript.exe..."
            Exit Sub
        End If
        .RestartUsing cscriptX86, .DoNotExit, .DoNotElevate
    End With
End Sub

Sub Main
    testRunner.SetSpecPattern "*.spec.vbs"
    testRunner.SetSpecFolder "..\dll"

    'if it is desired to run just a single test file, pass it in on the command line, using a relative path, relative to the spec folder. Also get the runCount from the command-line, arg #2, if specified

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
