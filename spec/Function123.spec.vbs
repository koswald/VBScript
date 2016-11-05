
'test Function123.vbs

With CreateObject("includer")
    Execute(.read("TestingFramework"))
    Execute(.read("Function123"))
End With
Dim f : Set f = New Function123

With New TestingFramework

    .describe "Function123 (birthing tub function)"

    .it "should calculate the function at 0"

        .AssertEqual f(0), 4

    .it "should calculate the function at pi/4" 'pi/4 radians = 45 degrees

        .AssertEqual Round(f(f.c.RADEG * f.c.pi/4), 12), Round(-9.51938674518403, 12)

End With
