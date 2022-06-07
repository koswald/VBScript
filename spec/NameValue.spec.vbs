'NameValue class integration test

Option Explicit
Dim nv 'NameValue object: what is to be tested
Dim incl 'VBScripting.Includer object
Dim actual, expected

Set incl = CreateObject( "VBScripting.Includer" )

Execute incl.Read( "TestingFramework" )
With New TestingFramework

    .Describe "NameValue class"
        Set nv = incl("NameValue").Init("City", "Reno")

    .It "should initialize on instantiation"
        actual = TypeName(nv) & nv.Name & nv.Value
        expected = "NameValueCityReno"
        .AssertEqual actual, expected

    .It "should initialize on instantiation #2"
        Execute incl.Read("NameValue")
        Set nv = New NameValue.Init("Dog","Beagle")
        actual = TypeName(nv) & nv.Name & nv.Value
        expected = "NameValueDogBeagle"
        .AssertEqual actual, expected

End With
