With CreateObject("includer")
    Execute .read("TestingFramework")
    Execute .read("PrivilegeChecker")
End With

With New TestingFramework

    .describe "PrivilegeChecker class"
        Dim pc : Set pc = New PrivilegeChecker

    .it "should indicate that privileges are not elevated"
        .AssertEqual pc, False

End With