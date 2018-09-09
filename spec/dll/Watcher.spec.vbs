Option Explicit : Setup
With New TestingFramework

    .describe "Watcher.dll/Watcher class"
        Set watcher = CreateObject("VBScripting.Watcher")
    .it "should initialize the default reset period"
        .AssertEqual watcher.ResetPeriod, 30000
    .it "should initialize Watch"
        .AssertEqual watcher.Watch, False
    .it "should return a (VBScript) Long for ResetPeriod"
        .AssertEqual TypeName(watcher.ResetPeriod), "Long"
    .it "should return a Byte for CurrentState"
        .AssertEqual TypeName(watcher.CurrentState), "Byte"
    .it "should accept the max value for ResetPeriod (2147483647)"
        On Error Resume Next
            watcher.ResetPeriod = 2147483647
            .AssertEqual Err.Description, ""
        On Error Goto 0
    .it "should not accept the max value +1 for ResetPeriod"
        On Error Resume Next
            watcher.ResetPeriod = 2147483648
            .AssertEqual Err.Description, "Overflow"
        On Error Goto 0

End With
Teardown

Dim watcher, includer
Sub Setup
    Set includer = CreateObject("VBScripting.Includer")
    ExecuteGlobal includer.read("TestingFramework")
End Sub
Sub Teardown
    watcher.Dispose
    Set watcher = Nothing
    Set includer = Nothing
End Sub
