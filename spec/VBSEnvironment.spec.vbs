
'test the VBSEnvironment class

With CreateObject( "VBScripting.Includer" )
    Execute .Read( "VBSEnvironment" )
    Execute .Read( "TestingFramework" )
End With
Dim sh : Set sh = CreateObject( "WScript.Shell" )
Dim fso : Set fso = CreateObject( "Scripting.FileSystemObject" )

With New TestingFramework

    .describe "VBSEnvironment class"

        Dim env : Set env = New VBSEnvironment

    .it "should create a user variable"

        Dim varName : varName = fso.GetBaseName(fso.GetTempName)
        Dim varValue : varValue = fso.GetBaseName(fso.GetTempName)
        'with Android Studio running on emulator,
        'creating a variable may be very slow
        Dim userEnv : Set userEnv = sh.Environment( "user" )
        env.CreateUserVar varName, varValue

        .AssertEqual userEnv(varName), varValue

    .it "should collapse a user variable"

        .AssertEqual env.collapse( sh.ExpandEnvironmentStrings("%" & varName & "%" )), "%" & varName & "%"

    .it "should remove a user variable"

        env.RemoveUserVar varName

        .AssertEqual userEnv(varName), ""

    .it "should create a process variable"

        varName = fso.GetBaseName(fso.GetTempName)
        varValue = fso.GetBaseName(fso.GetTempName)
        Dim proEnv : Set proEnv = sh.Environment( "process" )
        env.CreateProcessVar varName, varValue

        .AssertEqual proEnv(varName), varValue

    .it "should remove a process variable"

        env.RemoveProcessVar varName

        .AssertEqual proEnv(varName), ""

    .it "should expand an environment variable"

        .AssertEqual sh.ExpandEnvironmentStrings("%SystemRoot%"), env.Expand("%SystemRoot%")


    'garbage collection

        Set userEnv = Nothing
        Set proEnv = Nothing
        Set sh = Nothing
        Set fso = Nothing

End With
