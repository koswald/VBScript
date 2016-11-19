
'Launch the test runner for tests to be run with elevated privileges

Option Explicit
Main
Sub Main
    With CreateObject("includer")
        ExecuteGlobal(.read("VBSTestRunner"))
    End With
    Dim testRunner : Set testRunner = New VBSTestRunner
    With WScript.Arguments
        If .Count Then

            'if it is desired to run just a single test file, pass it in on the
            'command line, using a relative path, relative to the spec folder

            testRunner.SetSpecFile .item(0)
        End If
    End With

    'specify the folder containing the tests; path is relative to this script

    testRunner.SetSpecFolder "../spec"

    'specify elevated tests

    testRunner.SetSpecPattern ".*\.elevSpec\.vbs" 'default: .*\.spec\.vbs

    'run the tests

    testRunner.Run
End Sub
