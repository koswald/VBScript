
'test the VBSEnvironment class

With CreateObject("includer")
    Execute(.read("VBSEnvironment"))
    Execute(.read("VBSNatives"))
    Execute(.read("TestingFramework"))
End With

Dim n : Set n = New VBSNatives

With New TestingFramework

    .describe "VBSEnvironment class"

        Dim env : Set env = New VBSEnvironment

    .it "should create a user variable"

        Dim varName : varName = n.fso.GetBaseName(n.fso.GetTempName)
        Dim varValue : varValue = n.fso.GetBaseName(n.fso.GetTempName)
        'with Android Studio running on emulator, creating a variable may be very slow
        Dim userEnv : Set userEnv = n.sh.Environment("user")
        env.CreateUserVar varName, varValue

        .AssertEqual userEnv(varName), varValue

    .it "should collapse a the user variable"

        .AssertEqual env.collapse(n.sh.ExpandEnvironmentStrings("%" & varName & "%")), "%" & varName & "%"

    .it "should remove a user variable"

        env.RemoveUserVar varName

        .AssertEqual userEnv(varName), ""

    .it "should create a process variable"

        varName = n.fso.GetBaseName(n.fso.GetTempName)
        varValue = n.fso.GetBaseName(n.fso.GetTempName)
        Dim proEnv : Set proEnv = n.sh.Environment("process")
        env.CreateProcessVar varName, varValue

        .AssertEqual proEnv(varName), varValue

    .it "should remove a process variable"

        env.RemoveProcessVar varName

        .AssertEqual proEnv(varName), ""

    .it "should expand an environment variable"

        .AssertEqual n.sh.ExpandEnvironmentStrings("%SystemRoot%"), env.Expand("%SystemRoot%")


    'garbage collection

        Set userEnv = Nothing
        Set proEnv = Nothing

End With
