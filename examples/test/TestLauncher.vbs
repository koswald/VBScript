
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

            'get the runCount from the command-line, arg #2, if specified

            If .Count >= 2 Then testRunner.SetRunCount .item(1)
       End If
    End With

    'specify the folder containing the tests; path is relative to this script

    testRunner.SetSpecFolder "../../spec"

    'specify the reg ex pattern to match file types

    testRunner.SetSpecPattern ".*\.spec\.vbs|.*\.spec\.elev\+std\.vbs"

    'specify the time allotted for each test file to complete all of its specs, in seconds

    testRunner.SetTimeout 1 'default is 0; 0 => indefinite

    'run the tests

    On Error Resume Next
        testRunner.Run
        If &H80070006 = Err Then
            MsgBox "Start this script from a command window with CScript.exe:" & vbLf & vbLf & "cscript //nologo " & WScript.ScriptName, vbInformation, WScript.ScriptName
        ElseIf Err Then
            WScript.StdOut.Writeline Err.Description
        End If
    On Error Goto 0
End Sub
