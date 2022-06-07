Option Explicit
Dim watcher 'VBScripting.Watcher: object in test
With CreateObject( "VBScripting.Includer" )
    Execute .Read( "TestingFramework" )
End With
With New TestingFramework

    .Describe "VBScripting.Watcher object"
        Set watcher = CreateObject( "VBScripting.Watcher" )

    .It "should initialize the default reset period"
        .AssertEqual watcher.ResetPeriod, 30000

    .It "should initialize Watch"
        .AssertEqual watcher.Watch, False

    .It "should return a (VBScript) Long for ResetPeriod"
        .AssertEqual TypeName(watcher.ResetPeriod), "Long"

    .It "should return a Byte for CurrentState"
        .AssertEqual TypeName(watcher.CurrentState), "Byte"

    .It "should accept the max value for ResetPeriod (2147483647)"
        On Error Resume Next
            watcher.ResetPeriod = 2147483647
            .AssertEqual Err.Description, ""
        On Error Goto 0

    .It "should not accept the max value +1 for ResetPeriod"
        On Error Resume Next
            watcher.ResetPeriod = 2147483648
            .AssertEqual Err.Description, "Overflow"
        On Error Goto 0

End With

watcher.Dispose
Set watcher = Nothing

