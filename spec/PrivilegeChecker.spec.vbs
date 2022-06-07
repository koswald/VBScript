'Test the PrivilegeChecker object
'This test is valid only when privileges are not elevated.

Option Explicit
Dim pc 'PrivilegeChecker object being tested
Dim incl 'VBScripting.Includer object

Set incl = CreateObject( "VBScripting.Includer" )
Execute incl.Read( "TestingFramework" )

With New TestingFramework

    .describe "PrivilegeChecker class"
        Set pc = incl.LoadObject( "PrivilegeChecker" )

    .it "should indicate that privileges are not elevated"
        .AssertEqual pc, False

End With
