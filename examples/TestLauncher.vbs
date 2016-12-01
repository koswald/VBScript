
'Launch the test runner

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

    'specify the number of times to run the test(s)

    testRunner.SetRunCount 1

    'specify the folder containing the tests; path is relative to this script

    testRunner.SetSpecFolder "../spec"

    'run the tests

    On Error Resume Next
        testRunner.Run
        If Err.Number = &H80070006 Then MsgBox "Start this script from a command window with CScript.exe:" & vbLf & vbLf & "cscript //nologo " & WScript.ScriptName, vbInformation, WScript.ScriptName
    On Error Goto 0
End Sub
