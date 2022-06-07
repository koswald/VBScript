
'test the VBSExtracter class

With CreateObject( "VBScripting.Includer" )
    Execute .Read( "TestingFramework" )
    Execute .Read( "VBSExtracter" )
    Dim inputFile
    Execute .Read("..\spec\VBSExtracter.spec.config")
End With
Dim xtr : Set xtr = New VBSExtracter

With New TestingFramework

    .describe "VBSExtracter class"
    .it "should extract a string that spans multiple lines"
        xtr.SetFile inputFile
        xtr.SetPattern "<hta:application[^>]+id\s*=\s*""?[\w]+""?[^>]*>"
        .AssertEqual xtr.Extract, "<hta:application" & vbCrLf & vbTab & "id=""htaId"" >"

End With
