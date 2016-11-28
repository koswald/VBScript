
'test VBSTimer.vbs

With CreateObject("includer")
    Execute(.read("VBSTimer"))
    Execute(.read("TestingFramework"))
End With

With New TestingFramework

    'setup
        Const max = 0.02 'based on ~1.1 * max observed overhead
        Dim before, after
        
    .describe "VBSTimer class"
        Dim tmr : Set tmr = New VBSTimer
        
    .it "should get a split time"
        tmr.SetPrecision 2
        .AssertEqual tmr.Split < max, True
        
    .it "should have a default property"
        .AssertEqual tmr < max, True
        
    .it "should have adjustable precision"
        tmr.SetPrecision 1
        .AssertEqual tmr.GetPrecision, 1
        
    .it "should be resettable"
        tmr.SetPrecision 2
        WScript.Sleep 10 'millisecond(s) wait
        before = tmr.Split
        tmr.Reset
        after = tmr.Split
        .AssertEqual before > after, True

End With
