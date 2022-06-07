'MathConstants class integration test

Option Explicit
Dim mc 'MathConstants object: what is being tested
Dim incl 'VBScripting.Includer object

Set incl = CreateObject( "VBScripting.Includer" )

Execute incl.Read( "TestingFramework" )
With New TestingFramework

    .Describe "MathConstants class"
        Set mc = incl( "MathConstants" )

    .It "should return pi"
        .AssertEqual Round(mc.pi, 14), 3.14159265358979

    .It "should return pi/180, the degrees => radians converter"
        .AssertEqual Round(mc.DegRad, 14), 0.01745329251994

    .It "should return 180/pi, the radians => degrees converter"
        .AssertEqual Round(mc.RaDeg, 13), 57.2957795130824
        
    .It "should return 180/pi, the radians => degrees converter #2"
        .AssertEqual Round(mc.RadDeg, 13), 57.2957795130824
        
    .It "should return e"
        .AssertEqual Round(mc.e, 14), 2.71828182845905

End With
