
'test the VBSEnvironment class

With CreateObject("includer")
    Execute(.read("VBSEnvironment"))
    Execute(.read("VBSNatives"))
    Execute(.read("TestingFramework"))
End With

Dim n : Set n = New VBSNatives
Dim sh : Set sh = n.sh
Dim fso : Set fso = n.fso

With New TestingFramework

    .describe "VBSEnvironment class"

        Dim env : Set env = New VBSEnvironment

    .it "should create a user variable"

        Dim varName : varName = fso.GetBaseName(fso.GetTempName)
        Dim varValue : varValue = fso.GetBaseName(fso.GetTempName)
        Dim userEnv : Set userEnv = sh.Environment("user")
        env.CreateUserVar varName, varValue

        .AssertEqual userEnv(varName), varValue

    .it "should remove a user variable"

        env.RemoveUserVar varName

        .AssertEqual userEnv(varName), ""

    .it "should create a process variable"

        varName = fso.GetBaseName(fso.GetTempName)
        varValue = fso.GetBaseName(fso.GetTempName)
        Dim proEnv : Set proEnv = sh.Environment("process")
        env.CreateProcessVar varName, varValue

        .AssertEqual proEnv(varName), varValue

    .it "should remove a process variable"

        env.RemoveProcessVar varName

        .AssertEqual proEnv(varName), ""

    .it "should collapse a system variable"

        .AssertEqual env.collapse(sh.ExpandEnvironmentStrings("%UserProfile%\Dropbox")), "%drop%"

    .it "should collapse a the user variable"

        .AssertEqual env.collapse(sh.ExpandEnvironmentStrings("%UserProfile%\Google Drive\g")), "%g%"

    .it "should expand an environment variable"

        .AssertEqual sh.ExpandEnvironmentStrings("%SystemRoot%"), env.Expand("%SystemRoot%")

End With