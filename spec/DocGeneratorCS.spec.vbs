Option Explicit
Dim dgc 'DocGeneratorCS object; to be tested
Dim incl 'VBScripting.Includer object
Dim rawName 'string

Set incl = CreateObject( "VBScripting.Includer" )

Execute incl.Read( "TestingFramework" )
With New TestingFramework

    .describe "C# Class Documentation Generator"
        Set dgc = incl.LoadObject( "DocGeneratorCS" )

    .it "should parse rawName for the type"
        rawName = "M:VBScripting.IProgressBar.PBarSize(System.Int32,System.Int32)"
        .AssertEqual dgc.GetKind(rawName), "Method"

    .it "should change rawName to the expected value"
        .AssertEqual rawName, "VBScripting.IProgressBar.PBarSize(System.Int32,System.Int32)"

    .it "should parse the new rawName for the Namespace name"
        .AssertEqual dgc.GetNamespaceName(rawName), "VBScripting"

    .it "should change rawName to the expected value"
        .AssertEqual rawName, "IProgressBar.PBarSize(System.Int32,System.Int32)"

    .it "should parse the new rawName for the member type name"
        .AssertEqual dgc.GetTypeName(rawName), "IProgressBar"

    .it "should change rawName to the expected value"
        .AssertEqual rawName, "PBarSize(System.Int32,System.Int32)"

    .it "should parse the new rawName for the member type name for a Type"
        rawName = "ProgressBar"
        .AssertEqual dgc.GetTypeName(rawName), "ProgressBar"

    .it "should change rawName to the expected value"
        .AssertEqual rawName, ""

    .it "should parse the new rawName for the member name, without parameters"
        rawName = "PBarSize"
        .AssertEqual dgc.GetName(rawName), "PBarSize"

    .it "should parse the new rawName for the member name, with parameters"
        rawName = "PBarSize(System.Int32,System.Int32)"
        .AssertEqual dgc.GetName(rawName), "PBarSize"

End With
