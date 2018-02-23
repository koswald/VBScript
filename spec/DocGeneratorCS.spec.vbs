With CreateObject("VBScripting.Includer")
    Execute .Read("DocGeneratorCS")
    Execute .Read("TestingFramework")
End With
Set sh = CreateObject("WScript.Shell")
With New TestingFramework
    .describe "C# Class Documentation Generator"
        Dim dgc : Set dgc = New DocGeneratorCS

    .it "should parse rawName for the type"
        Dim rawName : rawName = "M:VBScripting.IProgressBar.PBarSize(System.Int32,System.Int32)"
        .AssertEqual dgc.GetType(rawName), "Method"

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
