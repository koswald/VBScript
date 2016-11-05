
'test WoWChecker.vbs

With CreateObject("includer")
    Execute(.read("TestingFramework"))
    Execute(.read("WoWChecker"))
End With

With New TestingFramework

    .describe "WoWChecker class"

    .it "should return False, with a correctly configured test"

        Set chkr = New WoWChecker

        .AssertEqual chkr.isWoW, False

    .it "should return an obj self reference on ByCheckSum call"

        Set chkr = New WowChecker.ByCheckSum

        .AssertEqual chkr.isSysWoW64, False

    .it "should return an obj self reference on BySize call"

        Set chkr = New WoWChecker.BySize

        .AssertEqual chkr.isSystem32, True

    .it "should have a public default property"

        .AssertEqual chkr, False

End With

'garbage collection

Set chkr = Nothing
