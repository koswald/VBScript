Option Explicit : Setup
With New TestingFramework

    .describe "VBScripting.Timer object 32-bit bug test"
        Dim csTimer : Set csTimer = CreateObject("VBScripting.Timer")

    .it "should reproduce error/bug"
        On Error Resume Next
            x = csTimer.Interval/1
            .AssertEqual Err.Description, _
               "Variable uses an Automation type not supported in VBScript"
        On Error Goto 0

    .it "should demonstrate workaround syntax: convert to CLng"
        On Error Resume Next
            x = CLng(csTimer.Interval)/1
            .AssertEqual Err.Description, ""
        On Error Goto 0

End With
Teardown

Dim x
Dim sh, fso, includer, wow

Sub Setup
    Set sh = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set includer = CreateObject("VBScripting.Includer")
    ExecuteGlobal includer.Read("TestingFramework")
    Execute includer.Read("WowChecker")
    Set wow = New WowChecker
    If wow.OSIs64Bit And Not wow.IsWow Then
        Teardown
        Err.Raise 1,, WScript.ScriptName & errMsg
    End If

    Const errMsg = " is intended to be run from %SystemRoot%\SysWOW64\cscript.exe"
End Sub
Sub Teardown
    Set sh = Nothing
    Set fso = Nothing
    Set wow = Nothing
    Set includer = Nothing
    Set csTimer = Nothing
End Sub
