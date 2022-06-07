'PrivilegeChecker class/object integration test.
'This test is valid only when privileges are elevated. 

Option Explicit
Dim pc 'PrivilegeChecker object being tested
Dim incl 'VBScripting.Includer object

Set incl = CreateObject( "VBScripting.Includer" )
Set pc = incl.LoadObject( "PrivilegeChecker" )

Execute incl.Read( "VBSApp" )
With New VBSApp
    If Not pc Then
        'Restart the script with elevated privileges
        .SetUserInteractive False
        .RestartUsing "cscript.exe", .DoNotExit, .DoElevate
    End If
End With

Execute incl.Read( "TestingFramework" )
With New TestingFramework

    .Describe "PrivilegeChecker class"

    .It "should indicate that privileges are elevated"
        .AssertEqual pc, True

End With
