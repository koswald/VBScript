Option Explicit
Dim gg 'GUIDGenerator object; to be tested
Dim re 'RegExp object
Dim incl 'VBScripting.Includer object

Set incl = CreateObject( "VBScripting.Includer" )
Set re = New RegExp
re.Pattern = "\{[A-F0-9]{8}-[A-F0-9]{4}-[A-F0-9]{4}-[A-F0-9]{4}-[A-F0-9]{12}}"

Execute incl.Read( "TestingFramework" )
With New TestingFramework

    .describe "GUIDGenerator class"
        Set gg = incl.LoadObject( "GUIDGenerator" )

    .it "should return a valid GUID on Generate call"
        .AssertEqual re.Test(gg.Generate), True

    .it "should return a valid GUID on default property call"
        .AssertEqual re.Test(gg), True

End With
