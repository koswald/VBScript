'Test VBSStopwatch.vbs

Option Explicit
Dim stopwatch 'VBStopwatch object: what is to be tested
Dim max 'float: approx. maximum error in split time, in seconds
Dim before, after 'float: split times

With CreateObject( "VBScripting.Includer" )
    Execute .Read( "VBSStopwatch" )
    Execute .Read( "TestingFramework" )
End With

With New TestingFramework

    .Describe "VBSStopwatch class"
        Set stopwatch = New VBSStopwatch
        stopwatch.SetPrecision 2
        max = 0.03 'seconds: based on ~1.1 * max observed overhead

    .It "should get a split time"
        .AssertEqual stopwatch.Split < max, True

    .It "should have a default property"
        .AssertEqual IsNumeric(stopwatch), True

    .It "should have adjustable precision"
        stopwatch.SetPrecision 1
        .AssertEqual stopwatch.GetPrecision, 1

    .It "should be resettable"
        stopwatch.SetPrecision 2
        WScript.Sleep 10 'millisecond(s) wait
        before = stopwatch.Split
        stopwatch.Reset
        after = stopwatch.Split
        .AssertEqual before > after, True

End With
