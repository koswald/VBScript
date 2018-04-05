'test VBSStopwatch.vbs
With CreateObject("VBScripting.Includer")
    Execute .read("VBSStopwatch")
    Execute .read("TestingFramework")
End With
With New TestingFramework
    .describe "VBSStopwatch class"
        Dim stopwatch : Set stopwatch = New VBSStopwatch
    'setup
        stopwatch.SetPrecision 2
        Const max = 0.02 'based on ~1.1 * max observed overhead
    .it "should get a split time"
        .AssertEqual stopwatch.Split < max, True
    .it "should have a default property"
        .AssertEqual IsNumeric(stopwatch), True
    .it "should have adjustable precision"
        stopwatch.SetPrecision 1
        .AssertEqual stopwatch.GetPrecision, 1
    .it "should be resettable"
        stopwatch.SetPrecision 2
        WScript.Sleep 10 'millisecond(s) wait
        Dim before : before = stopwatch.Split
        stopwatch.Reset
        Dim after : after = stopwatch.Split
        .AssertEqual before > after, True
End With
