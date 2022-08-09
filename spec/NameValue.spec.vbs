'NameValue class integration test

Option Explicit
Dim nv 'NameValue object: what is to be tested
Dim incl 'VBScripting.Includer object
Dim actual, expected

Set incl = CreateObject( "VBScripting.Includer" )

Execute incl.Read( "TestingFramework" )
With New TestingFramework

    .Describe "NameValue class"

    .It "should initialize on instantiation - LoadObject #1"
        Set nv = incl.LoadObject("NameValue").Init("City", "Sprague")
        actual = TypeName(nv) & nv.Name & nv.Value
        expected = "NameValueCitySprague"
        .AssertEqual actual, expected

    .It "should initialize on instantiation - LoadObject #2"
        Set nv = incl("NameValue").Init("City", "Salem")
        actual = TypeName(nv) & nv.Name & nv.Value
        expected = "NameValueCitySalem"
        .AssertEqual actual, expected

    .It "should initialize on instantiation - LoadObject #3"
        Set nv = (incl("NameValue"))("City", "Portland")
        actual = TypeName(nv) & nv.Name & nv.Value
        expected = "NameValueCityPortland"
        .AssertEqual actual, expected

    Execute incl.Read("NameValue")

    .It "should initialize on instantiation - New object #1"
        Set nv = New NameValue.Init("Dog","Beagle")
        actual = TypeName(nv) & nv.Name & nv.Value
        expected = "NameValueDogBeagle"
        .AssertEqual actual, expected

    .It "should initialize on instantiation - New object #2"
        Set nv = (New NameValue)("Dog","Terrier")
        actual = TypeName(nv) & nv.Name & nv.Value
        expected = "NameValueDogTerrier"
        .AssertEqual actual, expected

End With
