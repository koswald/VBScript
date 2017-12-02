
'test the VBSPower class

With CreateObject("includer")
    Execute .read("VBSPower")
    Execute .read("TestingFramework")
End With

With New TestingFramework

    .describe "VBSPower class"
        Dim pwr : Set pwr = New VBSPower

    'setup
        pwr.SetDebug True 'True => don't actually power down, restart, etc

    .it "should power down the computer"
        .AssertEqual TypeName(pwr.Shutdown), "Empty"

    .it "should restart the computer"
        .AssertEqual TypeName(pwr.Restart), "Empty"

    .it "should log off of the user session"
        .AssertEqual TypeName(pwr.Logoff), "Empty"

    .it "should put the computer to sleep"
        .AssertEqual TypeName(pwr.Sleep), "Empty"

    .it "should put the computer into hibernation"
        .AssertEqual TypeName(pwr.Hibernate), "Empty"

    .it "should enable hibernation"
        .AssertEqual TypeName(pwr.EnableHibernation), "Empty"

    .it "should disable hibernation"
        .AssertEqual TypeName(pwr.DisableHibernation), "Empty"

End With

