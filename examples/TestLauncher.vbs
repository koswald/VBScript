
'Launch the test runner

Main

Sub Main

    With CreateObject("includer")
        ExecuteGlobal(.read("VBSTestRunner"))
        ExecuteGlobal(.read("VBSHoster"))
    End With

    With New VBSHoster

        'restart this script, if necessary, hosted with cscript.exe
        'if restarting, opens in a new window

        .EnsureCScriptHost
    End With

    Dim testRunner : Set testRunner = New VBSTestRunner

    With WScript.Arguments
        If .Count Then

            'if it is desired to run just a single test file, pass it in on the 
            'command line, using a relative path, relative to the spec folder

            testRunner.SetSpecFile .item(0)
        End If
    End With

    'the spec folder contains the test files; path is relative to this script

    testRunner.SetSpecFolder "../spec"
    testRunner.Run

End Sub