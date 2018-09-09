Option Explicit : Setup
With New TestingFramework

    .describe "VBScripting.Timer class"
        Set vt = CreateObject("VBScripting.Timer")
    .it "should return a Single type for IntervalInHours"
        vt.IntervalInHours = 3
        .AssertEqual TypeName(vt.IntervalInHours), "Single"
    .it "should return the expected Single for IntervalInHours"
        vt.IntervalInHours = 1.54
        .AssertEqual vt.IntervalInHours, 1.54
    .it "should return a (VBScript) Long for Interval"
        .AssertEqual TypeName(vt.Interval), "Long"
    .it "should accept the max value for Interval (2147483647)"
        On Error Resume Next
            vt.Interval = 2147483647
            .AssertEqual Err.Description, ""
        On Error Goto 0
    .it "should not accept the max value +1 for Interval"
        On Error Resume Next
            vt.Interval = 2147483648
            .AssertEqual Err.Description, "Overflow"
        On Error Goto 0

End With
Teardown

Dim vt, sh, fso, includer
Sub Setup
    Set sh = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set includer = CreateObject("VBScripting.Includer")
    ExecuteGlobal includer.Read("TestingFramework")
End Sub
Sub Teardown
    Set vt = Nothing
    Set sh = Nothing
    Set fso = Nothing
    Set includer = Nothing
End Sub
