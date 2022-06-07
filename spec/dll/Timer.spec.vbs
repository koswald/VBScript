Option Explicit : Setup
With New TestingFramework

    .Describe "VBScripting.Timer object"
        Set vt = CreateObject( "VBScripting.Timer" )

    .It "should return a Single type for IntervalInHours"
        vt.IntervalInHours = 3
        .AssertEqual TypeName(vt.IntervalInHours), "Single"

    .It "should return the expected Single for IntervalInHours"
        vt.IntervalInHours = 1.54
        .AssertEqual vt.IntervalInHours, 1.54

    .It "should return a (VBScript) Long for Interval"
        .AssertEqual TypeName(vt.Interval), "Long"

    .It "should accept the max value for Interval (2147483647)"
        On Error Resume Next
            vt.Interval = 2147483647
            .AssertEqual Err.Description, ""
        On Error Goto 0

    .It "should not accept the max value +1 for Interval"
        On Error Resume Next
            vt.Interval = 2147483648
            .AssertEqual Err.Description, "Overflow"
        On Error Goto 0

End With

Teardown

Dim vt 'VBScripting.Timer

Sub Setup
    With CreateObject( "VBScripting.Includer" )
        ExecuteGlobal .Read( "TestingFramework" )
    End With
End Sub

Sub Teardown
    Set vt = Nothing
End Sub
