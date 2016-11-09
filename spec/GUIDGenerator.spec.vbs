
'test GUIDGenerator.vbs

With CreateObject("includer")
    Execute(.read("GUIDGenerator"))
    Execute(.read("TestingFramework"))
End With

With New TestingFramework

    .describe "GUIDGenerator class"

        Dim gg : Set gg = New GUIDGenerator

    'before

        Dim re : Set re = New RegExp
        re.Pattern = "\{[A-F0-9]{8}-[A-F0-9]{4}-[A-F0-9]{4}-[A-F0-9]{4}-[A-F0-9]{12}}"

    .it "should return a valid GUID on Generate call"

        .AssertEqual re.Test(gg.Generate), True

    .it "should return a valid GUID on default property call"

        .AssertEqual re.Test(gg), True

End With

