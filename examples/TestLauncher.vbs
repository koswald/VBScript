
'Launch the test runner

With CreateObject("includer")
	Execute(.read("VBSTestRunner"))
	Execute(.read("VBSHoster"))
End With

'restart this script, if necessary, hosted with cscript.exe; if restarting, opens in a new window

With New VBSHoster
    .EnsureCScriptHost
End With

Dim testRunner : Set testRunner = New VBSTestRunner

With WScript.Arguments
    If .Count Then
        testRunner.SetSpecFile .item(0) 'spec file is a file or a relative path/file, relative to the spec folder
    End If
End With

testRunner.SetSpecFolder "../spec" 'spec folder contains the test files; path is relative to this script
testRunner.Run
