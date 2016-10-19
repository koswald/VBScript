
'test the VBSEnvironment class

With CreateObject("includer")
    Execute(.read("VBSEnvironment"))
    Execute(.read("VBSNatives"))
    Execute(.read("TestingFramework"))
    Execute(.read("VBSLogger"))
End With

Dim n : Set n = New VBSNatives
Dim log : Set log = New VBSLogger

With New TestingFramework

    .describe "VBSEnvironment class"

        Dim env : Set env = New VBSEnvironment

    .it "should create a user variable"

        Dim varName : varName = n.fso.GetBaseName(n.fso.GetTempName)
        Dim varValue : varValue = n.fso.GetBaseName(n.fso.GetTempName)
        Dim userEnv : Set userEnv = n.sh.Environment("user")
        log "Creating a user variable" 'logging to measure time lag: when Android Studio is running an app, this step may be very slow
        env.CreateUserVar varName, varValue
        log "Done creating a user variable"

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
