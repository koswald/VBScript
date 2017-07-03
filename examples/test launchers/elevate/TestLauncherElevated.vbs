
'Launch the test runner

Option Explicit
Main
Sub Main
    With CreateObject("includer")
        Execute(.read("VBSTestRunner"))
    End With
    Dim testRunner : Set testRunner = New VBSTestRunner

    'specify the reg ex pattern to match file types

    testRunner.SetSpecPattern ".*\.spec\.elev\.vbs|.*\.spec\.elev\+std\.vbs"

    With WScript.Arguments
        If .Count > 1 Then

            'get the runCount and/or spec from the command-line

            SetCountOrPattern testRunner, 1
            If .Count > 2 Then SetCountOrPattern testRunner, 2
       End If
    End With

    'specify the folder containing the tests; path is relative to this script

    testRunner.SetSpecFolder "../../../spec"

    'specify the time allotted for each test file to complete all of its specs, in seconds

    testRunner.SetTimeout 10 'default is 0; 0 => indefinite

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

Sub SetCountOrPattern(runner, argIndex)
    With WScript.Arguments
        If IsInteger(.item(argIndex)) Then
            'arg specifies how many times to repeat the test(s)
            runner.SetRunCount .item(argIndex)
        Else
            'arg is a spec file: convert to regular expression: replace . with \. and + with \+
            runner.SetSpecPattern Replace(Replace(.item(argIndex), ".", "\."), "+", "\+")
        End If
    End With
End Sub

Function IsInteger(var)
    On Error Resume Next
        Dim CIntVar : CIntVar = CInt(var)
        IsInteger = Not CBool(Err)
    On Error Goto 0
End Function
