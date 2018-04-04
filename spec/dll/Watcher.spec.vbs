Option Explicit : Initialize
With New TestingFramework
    .describe "Watcher.dll/Watcher class"
        Dim watcher : Set watcher = CreateObject("VBScripting.Watcher")
    .it "should initialize the default reset period"
        .AssertEqual watcher.ResetPeriod, 30000
    .it "should initialize Watch"
        .AssertEqual watcher.Watch, False
    .it "should return a Long for ResetPeriod"
        .AssertEqual TypeName(watcher.ResetPeriod), "Long"
    .it "should return a Long for CurrentState"
        .AssertEqual TypeName(watcher.CurrentState), "Long"
End With

watcher.Dispose
Set watcher = Nothing

Sub Initialize
    With CreateObject("VBScripting.Includer")
        ExecuteGlobal .read("TestingFramework")
    End With
End Sub
