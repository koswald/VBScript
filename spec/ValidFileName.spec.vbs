
'test ValidFileName.vbs

With CreateObject("includer")
    Execute .read("ValidFileName")
    Execute .read("TestingFramework")
End With

With New TestingFramework

    .describe "ValidFileName.vbs"

    .it "should return a string suitable for a filename"

        .AssertEqual GetValidFileName("\/:*?""<>|%20#"), "-----------"

End With
