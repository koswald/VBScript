
'test VBSTimer.vbs

With CreateObject("VBScripting.Includer")
    Execute .read("VBSTimer")
    Execute .read("TestingFramework")
End With

With New TestingFramework

    .describe "VBSTimer class"
        Dim tmr : Set tmr = New VBSTimer

    'setup
        tmr.SetPrecision 2
        Const max = 0.02 'based on ~1.1 * max observed overhead

    .it "should get a split time"
        .AssertEqual tmr.Split < max, True

    .it "should have a default property"
        .AssertEqual IsNumeric(tmr), True

    .it "should have adjustable precision"
        tmr.SetPrecision 1
        .AssertEqual tmr.GetPrecision, 1

    .it "should be resettable"
        tmr.SetPrecision 2
        WScript.Sleep 10 'millisecond(s) wait
        Dim before : before = tmr.Split
        tmr.Reset
        Dim after : after = tmr.Split
        .AssertEqual before > after, True

End With
