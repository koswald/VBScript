
'test MathConstants.vbs

With CreateObject("includer")
    Execute(.read("MathConstants"))
    Execute(.read("TestingFramework"))
End With
Set c = New MathConstants

With New TestingFramework

    .describe "MathConstants class"

    .it "should return pi"

        .AssertEqual Round(c.pi, 14), 3.14159265358979

    .it "should return pi/180, the degrees => radians converter"

        .AssertEqual Round(c.DEGRAD, 14), 0.01745329251994

    .it "should return 180/pi, the radians => degrees converter"

        .AssertEqual Round(c.RADEG, 13), 57.2957795130823

End With
